use std::collections::{BTreeMap, HashMap};
use std::io::Cursor;

use formula_engine::date::{serial_to_ymd, ExcelDateSystem};
use formula_model::{sheet_name_eq_case_insensitive, CellRef, Range};
use quick_xml::events::Event;
use quick_xml::Reader;

use super::cache_definition::split_sheet_ref;
use crate::openxml::{local_name, parse_relationships, rels_part_name, resolve_target};
use crate::package::{XlsxError, XlsxPackage};
use crate::pivots::cache_records::PivotCacheValue;
use crate::pivots::PivotCacheSourceType;
use crate::shared_strings::parse_shared_strings_xml;
use crate::tables::{parse_table, TABLE_REL_TYPE};
use crate::xml::{QName, XmlElement, XmlNode};
use crate::DateSystem;

const WORKBOOK_PART: &str = "xl/workbook.xml";
const REL_TYPE_SHARED_STRINGS: &str =
    "http://schemas.openxmlformats.org/officeDocument/2006/relationships/sharedStrings";
const REL_TYPE_STYLES: &str =
    "http://schemas.openxmlformats.org/officeDocument/2006/relationships/styles";

fn text_eq_case_insensitive(a: &str, b: &str) -> bool {
    if a.is_ascii() && b.is_ascii() {
        return a.eq_ignore_ascii_case(b);
    }
    a.chars()
        .flat_map(|c| c.to_uppercase())
        .eq(b.chars().flat_map(|c| c.to_uppercase()))
}

#[derive(Debug, Clone, PartialEq, Eq)]
struct WorksheetSource {
    sheet: Option<String>,
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
        let (worksheet_part, range) = resolve_source_worksheet_and_range(self, &worksheet_source)?;
        let worksheet_xml = self
            .part(&worksheet_part)
            .ok_or_else(|| XlsxError::MissingPart(worksheet_part.clone()))?;

        let date_system = workbook_date_system(self)?;
        let field_date_flags = detect_date_fields(self, &cache_definition_bytes, range)?;

        let shared_strings = load_shared_strings(self)?;
        let cells = parse_worksheet_cells_in_range(worksheet_xml, range, &shared_strings)?;
        let (field_names, mut records) = build_cache_fields_and_records(range, &cells);

        coerce_date_fields(&mut records, &field_date_flags, date_system);

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

    /// Refresh every pivot cache whose `cacheSource` is worksheet-backed.
    ///
    /// This convenience wrapper iterates all `xl/pivotCache/pivotCacheDefinition*.xml` parts and
    /// refreshes any cache definition that declares `cacheSource type="worksheet"`.
    ///
    /// Non-worksheet cache sources (external, consolidation, scenario, unknown) are skipped.
    ///
    /// If any cache refresh fails, returns an error that includes the failing part name.
    pub fn refresh_all_pivot_caches_from_worksheets(&mut self) -> Result<(), XlsxError> {
        let defs = self.pivot_cache_definitions()?;
        for (part_name, def) in defs {
            if def.cache_source_type != PivotCacheSourceType::Worksheet {
                continue;
            }

            if let Err(err) = self.refresh_pivot_cache_from_worksheet(&part_name) {
                return Err(XlsxError::Invalid(format!(
                    "failed to refresh pivot cache {part_name:?} from worksheet source: {err}"
                )));
            }
        }

        Ok(())
    }
}

fn workbook_date_system(package: &XlsxPackage) -> Result<DateSystem, XlsxError> {
    let workbook_xml = package
        .part(WORKBOOK_PART)
        .ok_or_else(|| XlsxError::MissingPart(WORKBOOK_PART.to_string()))?;

    let mut reader = Reader::from_reader(Cursor::new(workbook_xml));
    reader.config_mut().trim_text(true);
    let mut buf = Vec::new();

    loop {
        match reader.read_event_into(&mut buf)? {
            Event::Start(e) | Event::Empty(e) if local_name(e.name().as_ref()) == b"workbookPr" => {
                for attr in e.attributes().with_checks(false) {
                    let attr = attr?;
                    if local_name(attr.key.as_ref()) == b"date1904" {
                        let val = attr.unescape_value()?.into_owned();
                        if val == "1" || val.eq_ignore_ascii_case("true") {
                            return Ok(DateSystem::V1904);
                        }
                    }
                }
                return Ok(DateSystem::V1900);
            }
            Event::Eof => break,
            _ => {}
        }
        buf.clear();
    }

    Ok(DateSystem::V1900)
}

fn parse_worksheet_source_from_cache_definition(xml: &[u8]) -> Result<WorksheetSource, XlsxError> {
    let mut reader = Reader::from_reader(Cursor::new(xml));
    reader.config_mut().trim_text(true);

    let mut buf = Vec::new();
    let mut sheet = None;
    let mut reference = None;
    let mut name = None;

    loop {
        match reader.read_event_into(&mut buf)? {
            Event::Start(e) | Event::Empty(e) => {
                if local_name(e.name().as_ref()) == b"worksheetSource" {
                    for attr in e.attributes().with_checks(false) {
                        let attr = attr?;
                        match local_name(attr.key.as_ref()) {
                            b"sheet" => sheet = Some(attr.unescape_value()?.into_owned()),
                            b"ref" => reference = Some(attr.unescape_value()?.into_owned()),
                            b"name" => name = Some(attr.unescape_value()?.into_owned()),
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

    let mut sheet = sheet;
    let mut reference = reference.or(name).ok_or(XlsxError::MissingAttr(
        "worksheetSource@ref or worksheetSource@name",
    ))?;

    // Some non-standard producers omit `worksheetSource@sheet` and embed the sheet name in the
    // ref (e.g. `Sheet1!A1:C5` or `'Sheet 1'!A1:C5`).
    //
    // Excel quotes sheet names with spaces/special characters using `'...'`, where any embedded
    // quote is escaped as `''`. Reuse the same parsing logic as the pivot cache definition parser.
    if let Some((parsed_sheet, parsed_ref)) = split_sheet_ref(&reference) {
        if sheet.is_none() {
            sheet = Some(parsed_sheet);
        }
        reference = parsed_ref;
    }
    Ok(WorksheetSource { sheet, reference })
}

fn resolve_source_worksheet_and_range(
    package: &XlsxPackage,
    worksheet_source: &WorksheetSource,
) -> Result<(String, Range), XlsxError> {
    // `worksheetSource` can point at:
    //   - a literal A1 range (e.g. `A1:D10` or `'Sheet 1'!$A$1:$D$10`)
    //   - a workbook/sheet-scoped defined name
    //   - an Excel Table (ListObject) name (`Table1`)
    //
    // Table sources are tricky because the `<table>` part doesn't encode the worksheet identity;
    // it's referenced from a worksheet's relationship file. Best-effort strategy:
    //   1) Prefer a sheet specified in the reference itself (`Sheet1!Table1`) or the
    //      `worksheetSource/@sheet` attribute.
    //   2) Otherwise, scan worksheets for relationships that target the matching table part.

    let (sheet_from_ref, token) = split_sheet_qualified_reference(&worksheet_source.reference);

    // 1) Literal A1 range.
    if let Ok(range) = Range::from_a1(token) {
        let sheet_name = sheet_from_ref
            .as_deref()
            .or_else(|| worksheet_source.sheet.as_deref())
            .ok_or_else(|| {
                XlsxError::Invalid(format!(
                    "worksheetSource reference {ref_:?} is an A1 range but no sheet was provided",
                    ref_ = worksheet_source.reference
                ))
            })?;
        let worksheet_part = resolve_worksheet_part(package, sheet_name)?;
        return Ok((worksheet_part, range));
    }

    // 2) Defined name.
    if let Some((defined_sheet, range)) = resolve_defined_name_reference(
        package,
        token,
        sheet_from_ref
            .as_deref()
            .or_else(|| worksheet_source.sheet.as_deref()),
    )? {
        let sheet_name = defined_sheet
            .as_deref()
            .or_else(|| sheet_from_ref.as_deref())
            .or_else(|| worksheet_source.sheet.as_deref())
            .ok_or_else(|| {
                XlsxError::Invalid(format!(
                    "defined name {name:?} does not specify a sheet and worksheetSource has no sheet attribute",
                    name = token
                ))
            })?;
        let worksheet_part = resolve_worksheet_part(package, sheet_name)?;
        return Ok((worksheet_part, range));
    }

    // 3) Table name.
    resolve_table_reference(
        package,
        token,
        sheet_from_ref.as_deref().or_else(|| worksheet_source.sheet.as_deref()),
    )?
    .ok_or_else(|| {
        XlsxError::Invalid(format!(
            "unable to resolve worksheetSource reference {ref_:?} as an A1 range, defined name, or table name",
            ref_ = worksheet_source.reference
        ))
    })
}

fn split_sheet_qualified_reference(input: &str) -> (Option<String>, &str) {
    let s = input.trim();
    if s.is_empty() {
        return (None, s);
    }
    let bytes = s.as_bytes();
    if bytes[0] == b'\'' {
        // Quoted sheet name: `'My Sheet'!A1:B2`. Inside the quotes, `''` escapes a literal `'`.
        let mut sheet = String::new();
        let mut i = 1usize;
        while i < bytes.len() {
            match bytes[i] {
                b'\'' => {
                    if bytes.get(i..).and_then(|s| s.get(1)) == Some(&b'\'') {
                        sheet.push('\'');
                        i += 2;
                        continue;
                    }
                    if bytes.get(i..).and_then(|s| s.get(1)) == Some(&b'!') {
                        let Some(rest_start) = i.checked_add(2) else {
                            debug_assert!(
                                false,
                                "expected i+2 in split_sheet_qualified_reference at i={i} len={} for {s:?}",
                                s.len()
                            );
                            return (None, s);
                        };
                        let Some(rest) = s.get(rest_start..) else {
                            debug_assert!(
                                false,
                                "expected utf-8 boundary at rest_start={rest_start} len={} for {s:?}",
                                s.len()
                            );
                            return (None, s);
                        };
                        return (Some(sheet), rest);
                    }
                    // Not actually a sheet-qualified reference; treat the whole string as the token.
                    return (None, s);
                }
                _ => {
                    let Some(tail) = s.get(i..) else {
                        debug_assert!(
                            false,
                            "expected char boundary at i={i} (len={}) for {s:?}",
                            s.len()
                        );
                        return (None, s);
                    };
                    let Some(ch) = tail.chars().next() else {
                        debug_assert!(
                            false,
                            "expected non-empty tail at i={i} (len={}) for {s:?}",
                            s.len()
                        );
                        return (None, s);
                    };
                    sheet.push(ch);
                    i += ch.len_utf8();
                    continue;
                }
            }
        }
        // Unterminated quote; treat as unqualified.
        return (None, s);
    }

    match s.find('!') {
        Some(idx) => {
            let Some(sheet) = s.get(..idx) else {
                debug_assert!(
                    false,
                    "expected utf-8 boundary at idx={idx} len={} for {s:?}",
                    s.len()
                );
                return (None, s);
            };
            let Some(rest_start) = idx.checked_add(1) else {
                debug_assert!(
                    false,
                    "expected idx+1 in split_sheet_qualified_reference at idx={idx} len={} for {s:?}",
                    s.len()
                );
                return (None, s);
            };
            let Some(rest) = s.get(rest_start..) else {
                debug_assert!(
                    false,
                    "expected utf-8 boundary at rest_start={rest_start} len={} for {s:?}",
                    s.len()
                );
                return (None, s);
            };
            (Some(sheet.to_string()), rest)
        }
        None => (None, s),
    }
}

#[derive(Debug, Clone)]
struct ParsedDefinedName {
    name: String,
    local_sheet_id: Option<u32>,
    value: String,
}

fn resolve_defined_name_reference(
    package: &XlsxPackage,
    name: &str,
    sheet_hint: Option<&str>,
) -> Result<Option<(Option<String>, Range)>, XlsxError> {
    let workbook_xml = package
        .part(WORKBOOK_PART)
        .ok_or_else(|| XlsxError::MissingPart(WORKBOOK_PART.to_string()))?;

    let mut reader = Reader::from_reader(Cursor::new(workbook_xml));
    reader.config_mut().trim_text(false);
    let mut buf = Vec::new();

    let mut matches: Vec<ParsedDefinedName> = Vec::new();
    let mut current: Option<ParsedDefinedName> = None;

    loop {
        match reader.read_event_into(&mut buf)? {
            Event::Start(e) if local_name(e.name().as_ref()) == b"definedName" => {
                let mut dn_name: Option<String> = None;
                let mut local_sheet_id: Option<u32> = None;
                for attr in e.attributes().with_checks(false) {
                    let attr = attr?;
                    match local_name(attr.key.as_ref()) {
                        b"name" => dn_name = Some(attr.unescape_value()?.into_owned()),
                        b"localSheetId" => {
                            local_sheet_id = attr.unescape_value()?.trim().parse::<u32>().ok();
                        }
                        _ => {}
                    }
                }

                let Some(dn_name) = dn_name else {
                    current = None;
                    continue;
                };
                if text_eq_case_insensitive(&dn_name, name) {
                    current = Some(ParsedDefinedName {
                        name: dn_name,
                        local_sheet_id,
                        value: String::new(),
                    });
                } else {
                    current = None;
                }
            }
            Event::Empty(e) if local_name(e.name().as_ref()) == b"definedName" => {
                let mut dn_name: Option<String> = None;
                let mut local_sheet_id: Option<u32> = None;
                for attr in e.attributes().with_checks(false) {
                    let attr = attr?;
                    match local_name(attr.key.as_ref()) {
                        b"name" => dn_name = Some(attr.unescape_value()?.into_owned()),
                        b"localSheetId" => {
                            local_sheet_id = attr.unescape_value()?.trim().parse::<u32>().ok();
                        }
                        _ => {}
                    }
                }
                let Some(dn_name) = dn_name else {
                    continue;
                };
                if text_eq_case_insensitive(&dn_name, name) {
                    matches.push(ParsedDefinedName {
                        name: dn_name,
                        local_sheet_id,
                        value: String::new(),
                    });
                }
            }
            Event::Text(e) if current.is_some() => {
                if let Some(ref mut dn) = current {
                    dn.value.push_str(&e.unescape()?.into_owned());
                }
            }
            Event::CData(e) if current.is_some() => {
                if let Some(ref mut dn) = current {
                    dn.value
                        .push_str(std::str::from_utf8(e.as_ref()).map_err(|err| {
                            XlsxError::Invalid(format!(
                                "defined name {name:?} contains invalid utf-8: {err}",
                                name = dn.name
                            ))
                        })?);
                }
            }
            Event::End(e) if local_name(e.name().as_ref()) == b"definedName" => {
                if let Some(dn) = current.take() {
                    matches.push(dn);
                }
            }
            Event::Eof => break,
            _ => {}
        }
        buf.clear();
    }

    if matches.is_empty() {
        return Ok(None);
    }

    // If multiple definitions exist (sheet-scoped names), try selecting one based on the
    // `worksheetSource/@sheet` hint.
    let selected = if matches.len() == 1 {
        matches.remove(0)
    } else {
        let Some(sheet_hint) = sheet_hint else {
            return Err(XlsxError::Invalid(format!(
                "defined name {name:?} is ambiguous (found {count} matches) and worksheetSource did not specify a sheet",
                count = matches.len()
            )));
        };

        let sheets = package.workbook_sheets()?;
        let hint_idx = sheets
            .iter()
            .position(|s| sheet_name_eq_case_insensitive(&s.name, sheet_hint))
            .map(|idx| idx as u32);

        let mut filtered: Vec<ParsedDefinedName> = matches
            .into_iter()
            .filter(|dn| match (dn.local_sheet_id, hint_idx) {
                (Some(local), Some(hint)) => local == hint,
                _ => false,
            })
            .collect();

        if filtered.len() == 1 {
            filtered.remove(0)
        } else {
            return Err(XlsxError::Invalid(format!(
                "defined name {name:?} is ambiguous for sheet {sheet_hint:?}",
            )));
        }
    };

    let mut formula = selected.value.trim();
    if let Some(stripped) = formula.strip_prefix('=') {
        formula = stripped.trim();
    }
    if formula.contains(',') {
        return Err(XlsxError::Invalid(format!(
            "defined name {name:?} refers to multiple areas ({formula:?}); only single-area ranges are supported",
            name = selected.name
        )));
    }

    let (sheet_from_value, ref_str) = split_sheet_qualified_reference(formula);
    let range = Range::from_a1(ref_str).map_err(|e| {
        XlsxError::Invalid(format!(
            "defined name {name:?} does not resolve to an A1 range ({ref_str:?}): {e}",
            name = selected.name
        ))
    })?;

    let sheet = if sheet_from_value.is_some() {
        sheet_from_value
    } else if let Some(local_id) = selected.local_sheet_id {
        let sheets = package.workbook_sheets()?;
        sheets.get(local_id as usize).map(|s| s.name.clone())
    } else {
        None
    };

    Ok(Some((sheet, range)))
}

fn resolve_table_reference(
    package: &XlsxPackage,
    name: &str,
    sheet_hint: Option<&str>,
) -> Result<Option<(String, Range)>, XlsxError> {
    let mut matched: Option<(String, Range)> = None;
    for (part_name, bytes) in package.parts() {
        // ZIP entry names in valid OOXML packages should use `/`, but some producers emit `\` or
        // vary casing; normalize to make table discovery best-effort.
        let canonical = part_name.trim_start_matches(|c| c == '/' || c == '\\');
        let canonical = if canonical.contains('\\') {
            canonical.replace('\\', "/")
        } else {
            canonical.to_string()
        };
        if !crate::ascii::starts_with_ignore_case(&canonical, "xl/tables/table")
            || !crate::ascii::ends_with_ignore_case(&canonical, ".xml")
        {
            continue;
        }
        let Ok(xml) = std::str::from_utf8(bytes) else {
            continue;
        };
        let Ok(table) = parse_table(xml) else {
            continue;
        };
        if !table.name.eq_ignore_ascii_case(name) && !table.display_name.eq_ignore_ascii_case(name)
        {
            continue;
        }

        let worksheet_part = if let Some(sheet_name) = sheet_hint {
            resolve_worksheet_part(package, sheet_name)?
        } else {
            resolve_worksheet_part_for_table(package, &canonical)?.ok_or_else(|| {
                XlsxError::Invalid(format!(
                    "table {name:?} found in {part:?} but worksheetSource did not specify a sheet and no worksheet relationship targets the table part",
                    name = table.name,
                    part = canonical
                ))
            })?
        };

        matched = Some((worksheet_part, table.range));
        break;
    }
    Ok(matched)
}

fn resolve_worksheet_part_for_table(
    package: &XlsxPackage,
    table_part: &str,
) -> Result<Option<String>, XlsxError> {
    let mut candidates = Vec::new();
    for part_name in package.part_names() {
        let worksheet_part = part_name.trim_start_matches(|c| c == '/' || c == '\\');
        let worksheet_part = if worksheet_part.contains('\\') {
            worksheet_part.replace('\\', "/")
        } else {
            worksheet_part.to_string()
        };
        if !crate::ascii::starts_with_ignore_case(&worksheet_part, "xl/worksheets/")
            || !crate::ascii::ends_with_ignore_case(&worksheet_part, ".xml")
        {
            continue;
        }
        let rels_part = rels_part_name(&worksheet_part);
        let Some(rels_bytes) = package.part(&rels_part) else {
            continue;
        };
        let rels = match parse_relationships(rels_bytes) {
            Ok(rels) => rels,
            Err(_) => continue,
        };
        for rel in rels {
            if rel.type_uri != TABLE_REL_TYPE {
                continue;
            }
            if rel
                .target_mode
                .as_deref()
                .is_some_and(|mode| mode.trim().eq_ignore_ascii_case("External"))
            {
                continue;
            }
            let target = resolve_target(&worksheet_part, &rel.target);
            let target = target.strip_prefix('/').unwrap_or(target.as_str());
            if crate::zip_util::zip_part_names_equivalent(target, table_part) {
                candidates.push(worksheet_part);
                break;
            }
        }
    }

    match candidates.len() {
        0 => Ok(None),
        1 => Ok(Some(candidates.remove(0))),
        _ => Err(XlsxError::Invalid(format!(
            "table part {table_part:?} is referenced from multiple worksheets ({candidates:?})",
        ))),
    }
}

fn resolve_worksheet_part(package: &XlsxPackage, sheet_name: &str) -> Result<String, XlsxError> {
    let sheets = package.workbook_sheets()?;
    let sheet = sheets
        .iter()
        // Sheet names use Excel's Unicode-aware, NFKC + case-insensitive comparison semantics.
        .find(|s| sheet_name_eq_case_insensitive(&s.name, sheet_name))
        .ok_or_else(|| XlsxError::Invalid(format!("sheet {sheet_name:?} not found in workbook")))?;

    let rels_part = rels_part_name("xl/workbook.xml");
    let guess = format!("xl/worksheets/sheet{}.xml", sheet.sheet_id);

    // Primary: resolve the sheet part through `/xl/_rels/workbook.xml.rels`.
    //
    // Some producers omit/mangle the relationship part; in that case fall back to the conventional
    // `xl/worksheets/sheet{sheetId}.xml` naming pattern when that part exists.
    if let Some(rels_bytes) = package.part(&rels_part) {
        if let Ok(rels) = parse_relationships(rels_bytes) {
            if let Some(rel) = rels.iter().find(|r| r.id == sheet.rel_id) {
                if !rel
                    .target_mode
                    .as_deref()
                    .is_some_and(|mode| mode.trim().eq_ignore_ascii_case("External"))
                {
                    let target = resolve_target("xl/workbook.xml", &rel.target);
                    // Best-effort: some producers emit workbook relationships that point to missing
                    // worksheet parts (or otherwise non-canonical names). If the resolved target is
                    // missing, fall back to the conventional `sheet{sheetId}.xml` filename when
                    // present. If neither part exists, still return the relationship target so
                    // downstream callers can surface a MissingPart error with the expected part
                    // name (and tests can validate the name-resolution logic in isolation).
                    if package.part(&target).is_some() {
                        return Ok(target);
                    }
                    if package.part(&guess).is_some() {
                        return Ok(guess);
                    }
                    return Ok(target);
                }
            }
        }
    }

    if package.part(&guess).is_some() {
        return Ok(guess);
    }

    Err(XlsxError::Invalid(format!(
        "unable to resolve worksheet part for sheet {sheet_name:?}: expected workbook relationship {rid:?} in {rels_part:?} or fallback worksheet part {guess:?}",
        rid = sheet.rel_id
    )))
}

fn resolve_cache_records_part(
    package: &XlsxPackage,
    cache_definition_part: &str,
) -> Result<String, XlsxError> {
    let rels_part = rels_part_name(cache_definition_part);
    let rels_error = if let Some(rels_bytes) = package.part(&rels_part) {
        match parse_relationships(rels_bytes) {
            Ok(rels) => {
                let mut pivot_rel_seen = false;
                let mut pivot_rel_error: Option<String> = None;

                for rel in rels {
                    if !rel.type_uri.ends_with("/pivotCacheRecords") {
                        continue;
                    }
                    pivot_rel_seen = true;

                    if rel
                        .target_mode
                        .as_deref()
                        .is_some_and(|mode| mode.trim().eq_ignore_ascii_case("External"))
                    {
                        pivot_rel_error.get_or_insert_with(|| {
                            format!(
                                "pivotCacheRecords relationship in {rels_part:?} is marked External"
                            )
                        });
                        continue;
                    }

                    let target = resolve_target(cache_definition_part, &rel.target);
                    if package.part(&target).is_some() {
                        return Ok(target);
                    }

                    pivot_rel_error.get_or_insert_with(|| {
                        format!(
                            "pivotCacheRecords relationship in {rels_part:?} targets missing part {target:?}"
                        )
                    });
                }

                if pivot_rel_seen {
                    pivot_rel_error.unwrap_or_else(|| {
                        format!("{rels_part:?} does not contain a usable pivotCacheRecords relationship")
                    })
                } else {
                    format!("{rels_part:?} does not contain a pivotCacheRecords relationship")
                }
            }
            Err(err) => format!("failed to parse {rels_part:?}: {err}"),
        }
    } else {
        format!("missing relationship part {rels_part:?}")
    };

    // Fallback: infer `pivotCacheRecordsN.xml` from the cache definition naming convention when
    // the relationship part is missing/malformed.
    let normalized = cache_definition_part
        .strip_prefix('/')
        .unwrap_or(cache_definition_part);
    let Some(def_idx) = normalized
        .strip_prefix("xl/pivotCache/pivotCacheDefinition")
        .and_then(|rest| rest.strip_suffix(".xml"))
        .filter(|idx| !idx.is_empty() && idx.as_bytes().iter().all(|b| b.is_ascii_digit()))
    else {
        return Err(XlsxError::Invalid(format!(
            "unable to resolve pivot cache records part for {cache_definition_part:?}: {rels_error}. \
cache definition part name does not match the conventional `xl/pivotCache/pivotCacheDefinitionN.xml` scheme",
            rels_error = rels_error,
        )));
    };

    let candidate = format!("xl/pivotCache/pivotCacheRecords{def_idx}.xml");
    if package.part(&candidate).is_some() {
        return Ok(candidate);
    }

    Err(XlsxError::Invalid(format!(
        "unable to resolve pivot cache records part for {cache_definition_part:?}: {rels_error}. \
guessed {candidate:?} from the cache definition index, but that part is missing",
        rels_error = rels_error,
    )))
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

fn detect_date_fields(
    package: &XlsxPackage,
    cache_definition_xml: &[u8],
    range: Range,
) -> Result<Vec<bool>, XlsxError> {
    let field_count = range.width() as usize;
    if field_count == 0 {
        return Ok(Vec::new());
    }

    let num_fmt_ids = parse_cache_definition_num_fmt_ids(cache_definition_xml)?;
    let num_fmts = load_styles_num_fmts(package)?;

    let mut flags = Vec::new();
    if flags.try_reserve_exact(field_count).is_err() {
        return Err(XlsxError::AllocationFailure("detect_date_fields output"));
    }
    for idx in 0..field_count {
        let num_fmt_id = num_fmt_ids.get(idx).copied().unwrap_or(0);
        flags.push(is_datetime_num_fmt_id(num_fmt_id, &num_fmts));
    }
    Ok(flags)
}

fn parse_cache_definition_num_fmt_ids(xml: &[u8]) -> Result<Vec<u16>, XlsxError> {
    let root = XmlElement::parse(xml).map_err(|e| {
        XlsxError::Invalid(format!("failed to parse pivot cache definition xml: {e}"))
    })?;
    let Some(cache_fields) = root.child("cacheFields") else {
        return Ok(Vec::new());
    };

    let mut out = Vec::new();
    for cache_field in cache_fields.children_by_local("cacheField") {
        let id = cache_field
            .attr("numFmtId")
            .and_then(|v| v.parse::<u16>().ok())
            .unwrap_or(0);
        out.push(id);
    }

    Ok(out)
}

fn load_styles_num_fmts(package: &XlsxPackage) -> Result<HashMap<u16, String>, XlsxError> {
    let styles_part = resolve_styles_part_name(package)?.or_else(|| {
        package
            .part("xl/styles.xml")
            .map(|_| "xl/styles.xml".to_string())
    });
    let Some(styles_part) = styles_part else {
        return Ok(HashMap::new());
    };
    let Some(bytes) = package.part(&styles_part) else {
        return Ok(HashMap::new());
    };

    parse_styles_num_fmts_xml(bytes)
        .map_err(|e| XlsxError::Invalid(format!("failed to parse {styles_part:?}: {e}")))
}

fn resolve_styles_part_name(package: &XlsxPackage) -> Result<Option<String>, XlsxError> {
    let rels_part = rels_part_name(WORKBOOK_PART);
    let rels_bytes = match package.part(&rels_part) {
        Some(bytes) => bytes,
        None => return Ok(None),
    };
    let rels = match parse_relationships(rels_bytes) {
        Ok(rels) => rels,
        Err(_) => return Ok(None),
    };
    for rel in rels {
        if rel.type_uri != REL_TYPE_STYLES {
            continue;
        }
        if rel
            .target_mode
            .as_deref()
            .is_some_and(|mode| mode.trim().eq_ignore_ascii_case("External"))
        {
            continue;
        }
        let target = resolve_target(WORKBOOK_PART, &rel.target);
        if package.part(&target).is_some() {
            return Ok(Some(target));
        }
    }
    Ok(None)
}

fn parse_styles_num_fmts_xml(xml: &[u8]) -> Result<HashMap<u16, String>, quick_xml::Error> {
    let mut reader = Reader::from_reader(Cursor::new(xml));
    reader.config_mut().trim_text(true);

    let mut buf = Vec::new();
    let mut in_num_fmts = false;
    let mut out: HashMap<u16, String> = HashMap::new();

    loop {
        match reader.read_event_into(&mut buf)? {
            Event::Start(e) if local_name(e.name().as_ref()) == b"numFmts" => {
                in_num_fmts = true;
            }
            Event::End(e) if local_name(e.name().as_ref()) == b"numFmts" => {
                in_num_fmts = false;
            }
            Event::Empty(e) if in_num_fmts && local_name(e.name().as_ref()) == b"numFmt" => {
                if let Some((id, code)) = parse_num_fmt_attrs(&e)? {
                    out.insert(id, code);
                }
            }
            Event::Start(e) if in_num_fmts && local_name(e.name().as_ref()) == b"numFmt" => {
                if let Some((id, code)) = parse_num_fmt_attrs(&e)? {
                    out.insert(id, code);
                }
                reader.read_to_end_into(e.name(), &mut Vec::new())?;
            }
            Event::Eof => break,
            _ => {}
        }
        buf.clear();
    }

    Ok(out)
}

fn parse_num_fmt_attrs(
    e: &quick_xml::events::BytesStart<'_>,
) -> Result<Option<(u16, String)>, quick_xml::Error> {
    let mut id: Option<u16> = None;
    let mut code: Option<String> = None;
    for attr in e.attributes().with_checks(false) {
        let attr = attr?;
        match local_name(attr.key.as_ref()) {
            b"numFmtId" => id = attr.unescape_value()?.trim().parse::<u16>().ok(),
            b"formatCode" => code = Some(attr.unescape_value()?.into_owned()),
            _ => {}
        }
    }
    match (id, code) {
        (Some(id), Some(code)) => Ok(Some((id, code))),
        _ => Ok(None),
    }
}

fn is_datetime_num_fmt_id(num_fmt_id: u16, num_fmts: &HashMap<u16, String>) -> bool {
    if num_fmt_id == 0 {
        return false;
    }

    if let Some(code) = num_fmts.get(&num_fmt_id) {
        return format_code_looks_like_datetime(code);
    }

    if let Some(code) = formula_format::builtin_format_code(num_fmt_id) {
        return format_code_looks_like_datetime(code);
    }

    // Excel reserves additional built-in numFmtId slots for locale-specific date/time patterns
    // (commonly 50-58). These frequently represent dates in real-world files even when the
    // corresponding format code is not present in styles.xml.
    (50..=58).contains(&num_fmt_id)
}

fn format_code_looks_like_datetime(code: &str) -> bool {
    let mut in_quotes = false;
    let mut escape = false;
    let mut chars = code.chars().peekable();

    while let Some(ch) = chars.next() {
        if escape {
            escape = false;
            continue;
        }
        if in_quotes {
            if ch == '"' {
                in_quotes = false;
            }
            continue;
        }

        match ch {
            '"' => in_quotes = true,
            '\\' => escape = true,
            '[' => {
                // Elapsed time: [h], [m], [s]
                let mut content = String::new();
                while let Some(c) = chars.next() {
                    if c == ']' {
                        break;
                    }
                    content.push(c);
                }
                if !content.is_empty()
                    && (content.chars().all(|c| matches!(c, 'h' | 'H'))
                        || content.chars().all(|c| matches!(c, 'm' | 'M'))
                        || content.chars().all(|c| matches!(c, 's' | 'S')))
                {
                    return true;
                }
            }
            'y' | 'Y' | 'd' | 'D' | 'h' | 'H' | 's' | 'S' => return true,
            'm' | 'M' => return true,
            'a' | 'A' => {
                // AM/PM or A/P markers (case-insensitive).
                let mut probe = String::new();
                probe.push(ch);
                let mut clone = chars.clone();
                for _ in 0..4 {
                    if let Some(c) = clone.next() {
                        probe.push(c);
                    } else {
                        break;
                    }
                }
                if crate::ascii::starts_with_ignore_case(&probe, "am/pm")
                    || crate::ascii::starts_with_ignore_case(&probe, "a/p")
                {
                    return true;
                }
            }
            _ => {}
        }
    }

    false
}

fn coerce_date_fields(
    records: &mut [Vec<PivotCacheValue>],
    date_fields: &[bool],
    date_system: DateSystem,
) {
    if date_fields.is_empty() {
        return;
    }

    for record in records {
        for (idx, value) in record.iter_mut().enumerate() {
            if !date_fields.get(idx).copied().unwrap_or(false) {
                continue;
            }

            if let PivotCacheValue::Number(n) = value {
                if let Some(dt) = excel_serial_to_pivot_datetime(*n, date_system) {
                    *value = PivotCacheValue::DateTime(dt);
                }
            }
        }
    }
}

fn excel_serial_to_pivot_datetime(serial: f64, date_system: DateSystem) -> Option<String> {
    if !serial.is_finite() || serial < 0.0 {
        return None;
    }

    let mut days = serial.floor() as i64;
    let frac = serial - (days as f64);
    let mut seconds = (frac * 86_400.0).round() as i64;

    if seconds >= 86_400 {
        seconds = 0;
        days = days.checked_add(1)?;
    }

    let serial_days: i32 = days.try_into().ok()?;
    let system = match date_system {
        DateSystem::V1900 => ExcelDateSystem::EXCEL_1900,
        DateSystem::V1904 => ExcelDateSystem::Excel1904,
    };
    let date = serial_to_ymd(serial_days, system).ok()?;

    let hour = (seconds / 3_600) as u32;
    let minute = ((seconds % 3_600) / 60) as u32;
    let second = (seconds % 60) as u32;

    Some(format!(
        "{:04}-{:02}-{:02}T{:02}:{:02}:{:02}Z",
        date.year, date.month, date.day, hour, minute, second
    ))
}

fn resolve_shared_strings_part_name(package: &XlsxPackage) -> Result<Option<String>, XlsxError> {
    let rels_part = rels_part_name(WORKBOOK_PART);
    let rels_bytes = match package.part(&rels_part) {
        Some(bytes) => bytes,
        None => return Ok(None),
    };
    let rels = match parse_relationships(rels_bytes) {
        Ok(rels) => rels,
        Err(_) => return Ok(None),
    };
    for rel in rels {
        if rel.type_uri != REL_TYPE_SHARED_STRINGS {
            continue;
        }
        if rel
            .target_mode
            .as_deref()
            .is_some_and(|mode| mode.trim().eq_ignore_ascii_case("External"))
        {
            continue;
        }
        let target = resolve_target(WORKBOOK_PART, &rel.target);
        if package.part(&target).is_some() {
            return Ok(Some(target));
        }
    }
    Ok(None)
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
        Some("e") => PivotCacheValue::Error(v_text.unwrap_or_default().to_string()),
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
            PivotCacheValue::DateTime(dt) => {
                build_value_element(ns.clone(), "d", Some(dt.as_str()))
            }
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
    use roxmltree::Document;
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

    #[test]
    fn load_shared_strings_falls_back_when_workbook_rels_target_missing() {
        // Workbook relationship points to a missing part, but the canonical `xl/sharedStrings.xml`
        // exists. We should still load it.
        let workbook_rels = r#"<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
  <Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/sharedStrings" Target="missingSharedStrings.xml"/>
</Relationships>"#;

        let shared_strings_xml = r#"<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<sst xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" count="1" uniqueCount="1">
  <si><t>Alpha</t></si>
</sst>"#;

        let cursor = Cursor::new(Vec::new());
        let mut zip = zip::ZipWriter::new(cursor);
        let options = zip::write::FileOptions::<()>::default()
            .compression_method(zip::CompressionMethod::Deflated);

        zip.start_file("xl/_rels/workbook.xml.rels", options)
            .unwrap();
        zip.write_all(workbook_rels.as_bytes()).unwrap();

        zip.start_file("xl/sharedStrings.xml", options).unwrap();
        zip.write_all(shared_strings_xml.as_bytes()).unwrap();

        let bytes = zip.finish().unwrap().into_inner();
        let pkg = XlsxPackage::from_bytes(&bytes).expect("read pkg");

        let strings = load_shared_strings(&pkg).expect("load shared strings");
        assert_eq!(strings, vec!["Alpha".to_string()]);
    }

    #[test]
    fn load_styles_num_fmts_falls_back_when_workbook_rels_target_missing() {
        // Workbook relationship points to a missing styles part, but the canonical `xl/styles.xml`
        // exists. We should still load it so date detection works.
        let workbook_rels = r#"<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
  <Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/styles" Target="missingStyles.xml"/>
</Relationships>"#;

        let styles_xml = r#"<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<styleSheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main">
  <numFmts count="1">
    <numFmt numFmtId="164" formatCode="m/d/yyyy"/>
  </numFmts>
</styleSheet>"#;

        let cursor = Cursor::new(Vec::new());
        let mut zip = zip::ZipWriter::new(cursor);
        let options = zip::write::FileOptions::<()>::default()
            .compression_method(zip::CompressionMethod::Deflated);

        zip.start_file("xl/_rels/workbook.xml.rels", options)
            .unwrap();
        zip.write_all(workbook_rels.as_bytes()).unwrap();

        zip.start_file("xl/styles.xml", options).unwrap();
        zip.write_all(styles_xml.as_bytes()).unwrap();

        let bytes = zip.finish().unwrap().into_inner();
        let pkg = XlsxPackage::from_bytes(&bytes).expect("read pkg");

        let num_fmts = load_styles_num_fmts(&pkg).expect("load num formats");
        assert_eq!(num_fmts.get(&164).map(String::as_str), Some("m/d/yyyy"));
    }

    #[test]
    fn refresh_pivot_cache_resolves_table_name_source() {
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
</Relationships>"#;

        // Table range is A1:B3 (header + 2 records).
        let worksheet_xml = r#"<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main"
 xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">
  <sheetData>
    <row r="1">
      <c r="A1" t="inlineStr"><is><t>Header1</t></is></c>
      <c r="B1" t="inlineStr"><is><t>Header2</t></is></c>
    </row>
    <row r="2">
      <c r="A2"><v>1</v></c>
      <c r="B2" t="inlineStr"><is><t>Alpha</t></is></c>
    </row>
    <row r="3">
      <c r="A3"><v>2</v></c>
      <c r="B3" t="inlineStr"><is><t>Beta</t></is></c>
    </row>
  </sheetData>
  <tableParts count="1">
    <tablePart r:id="rIdTable1"/>
  </tableParts>
</worksheet>"#;

        let worksheet_rels = r#"<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
  <Relationship Id="rIdTable1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/table" Target="../tables/table1.xml"/>
</Relationships>"#;

        let table_xml = r#"<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<table xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" id="1" name="Table1" displayName="Table1" ref="A1:B3" headerRowCount="1" totalsRowCount="0">
  <tableColumns count="2">
    <tableColumn id="1" name="Header1"/>
    <tableColumn id="2" name="Header2"/>
  </tableColumns>
</table>"#;

        // `worksheetSource/@ref` points at the table name (not an A1 range).
        let cache_definition_xml = r#"<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<pivotCacheDefinition xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" recordCount="0">
  <cacheSource type="worksheet">
    <worksheetSource ref="Table1"/>
  </cacheSource>
  <cacheFields count="0"/>
</pivotCacheDefinition>"#;

        let cache_definition_rels = r#"<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
  <Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/pivotCacheRecords" Target="pivotCacheRecords1.xml"/>
</Relationships>"#;

        let cache_records_xml = r#"<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<pivotCacheRecords xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" count="0"/>"#;

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

        zip.start_file("xl/worksheets/_rels/sheet1.xml.rels", options)
            .unwrap();
        zip.write_all(worksheet_rels.as_bytes()).unwrap();

        zip.start_file("xl/tables/table1.xml", options).unwrap();
        zip.write_all(table_xml.as_bytes()).unwrap();

        zip.start_file("xl/pivotCache/pivotCacheDefinition1.xml", options)
            .unwrap();
        zip.write_all(cache_definition_xml.as_bytes()).unwrap();

        zip.start_file(
            "xl/pivotCache/_rels/pivotCacheDefinition1.xml.rels",
            options,
        )
        .unwrap();
        zip.write_all(cache_definition_rels.as_bytes()).unwrap();

        zip.start_file("xl/pivotCache/pivotCacheRecords1.xml", options)
            .unwrap();
        zip.write_all(cache_records_xml.as_bytes()).unwrap();

        let bytes = zip.finish().unwrap().into_inner();
        let mut pkg = XlsxPackage::from_bytes(&bytes).expect("read pkg");
        pkg.refresh_pivot_cache_from_worksheet("xl/pivotCache/pivotCacheDefinition1.xml")
            .expect("refresh");

        let updated_def =
            std::str::from_utf8(pkg.part("xl/pivotCache/pivotCacheDefinition1.xml").unwrap())
                .unwrap();
        let doc = Document::parse(updated_def).expect("parse updated cache definition");
        let root = doc.root_element();
        assert_eq!(root.attribute("recordCount"), Some("2"));
        let cache_fields = root
            .descendants()
            .find(|n| n.is_element() && n.tag_name().name() == "cacheFields")
            .expect("cacheFields");
        assert_eq!(cache_fields.attribute("count"), Some("2"));
        let field_names: Vec<_> = cache_fields
            .children()
            .filter(|n| n.is_element() && n.tag_name().name() == "cacheField")
            .filter_map(|n| n.attribute("name"))
            .collect();
        assert_eq!(field_names, vec!["Header1", "Header2"]);

        let updated_records =
            std::str::from_utf8(pkg.part("xl/pivotCache/pivotCacheRecords1.xml").unwrap()).unwrap();
        let doc = Document::parse(updated_records).expect("parse updated cache records");
        let root = doc.root_element();
        assert_eq!(root.attribute("count"), Some("2"));
        let record_count = root
            .children()
            .filter(|n| n.is_element() && n.tag_name().name() == "r")
            .count();
        assert_eq!(record_count, 2);
    }

    #[test]
    fn resolve_worksheet_part_for_table_matches_case_insensitive_rels_targets() {
        // Some producers differ in how they case relationship targets vs ZIP entry names. Ensure we
        // still detect the worksheet that owns a table even when the relationship target casing
        // doesn't match the table part name.
        let worksheet_xml = r#"<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main">
  <sheetData/>
</worksheet>"#;

        let worksheet_rels = r#"<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
  <Relationship Id="rIdTable1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/table" Target="../tables/Table1.xml"/>
</Relationships>"#;

        let cursor = Cursor::new(Vec::new());
        let mut zip = zip::ZipWriter::new(cursor);
        let options = zip::write::FileOptions::<()>::default()
            .compression_method(zip::CompressionMethod::Deflated);

        zip.start_file("xl/worksheets/sheet1.xml", options).unwrap();
        zip.write_all(worksheet_xml.as_bytes()).unwrap();

        zip.start_file("xl/worksheets/_rels/sheet1.xml.rels", options)
            .unwrap();
        zip.write_all(worksheet_rels.as_bytes()).unwrap();

        zip.start_file("xl/tables/table1.xml", options).unwrap();
        zip.write_all(b"<table/>").unwrap();

        let bytes = zip.finish().unwrap().into_inner();
        let pkg = XlsxPackage::from_bytes(&bytes).expect("read pkg");

        let worksheet_part =
            resolve_worksheet_part_for_table(&pkg, "xl/tables/table1.xml").expect("resolve");
        assert_eq!(worksheet_part.as_deref(), Some("xl/worksheets/sheet1.xml"));
    }

    #[test]
    fn resolve_table_reference_discovers_tables_with_backslash_and_case_entry_names() {
        let worksheet_xml = r#"<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main">
  <sheetData/>
</worksheet>"#;

        let worksheet_rels = r#"<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
  <Relationship Id="rIdTable1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/table" Target="..\\tables\\table1.xml"/>
</Relationships>"#;

        let table_xml = r#"<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<table xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" id="1" name="Table1" displayName="Table1" ref="A1:B3" headerRowCount="1" totalsRowCount="0">
  <tableColumns count="2">
    <tableColumn id="1" name="Header1"/>
    <tableColumn id="2" name="Header2"/>
  </tableColumns>
</table>"#;

        let cursor = Cursor::new(Vec::new());
        let mut zip = zip::ZipWriter::new(cursor);
        let options = zip::write::FileOptions::<()>::default()
            .compression_method(zip::CompressionMethod::Deflated);

        zip.start_file("xl/worksheets/sheet1.xml", options).unwrap();
        zip.write_all(worksheet_xml.as_bytes()).unwrap();

        zip.start_file("xl/worksheets/_rels/sheet1.xml.rels", options)
            .unwrap();
        zip.write_all(worksheet_rels.as_bytes()).unwrap();

        // Backslash separators and case differences are invalid in canonical XLSX ZIP entry names,
        // but are tolerated by `XlsxPackage::part`. Ensure table discovery normalizes them too.
        zip.start_file("XL\\Tables\\Table1.xml", options).unwrap();
        zip.write_all(table_xml.as_bytes()).unwrap();

        let bytes = zip.finish().unwrap().into_inner();
        let pkg = XlsxPackage::from_bytes(&bytes).expect("read pkg");

        let resolved = resolve_table_reference(&pkg, "Table1", None).expect("resolve");
        let (sheet_part, range) = resolved.expect("expected match");
        assert_eq!(sheet_part, "xl/worksheets/sheet1.xml");
        assert_eq!(range, Range::from_a1("A1:B3").unwrap());
    }

    #[test]
    fn refresh_pivot_cache_supports_named_range_worksheet_source() {
        let workbook_xml = r#"<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<workbook xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main"
 xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">
  <sheets>
    <sheet name="Sheet1" sheetId="1" r:id="rId1"/>
  </sheets>
  <definedNames>
    <definedName name="MyNamedRange">Sheet1!$A$1:$B$3</definedName>
  </definedNames>
</workbook>"#;

        let workbook_rels = r#"<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
  <Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/worksheet" Target="worksheets/sheet1.xml"/>
</Relationships>"#;

        // Named range points at A1:B3 (header + 2 records).
        let worksheet_xml = r#"<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main">
  <sheetData>
    <row r="1">
      <c r="A1" t="inlineStr"><is><t>Name</t></is></c>
      <c r="B1" t="inlineStr"><is><t>Age</t></is></c>
    </row>
    <row r="2">
      <c r="A2" t="inlineStr"><is><t>Alice</t></is></c>
      <c r="B2"><v>30</v></c>
    </row>
    <row r="3">
      <c r="A3" t="inlineStr"><is><t>Bob</t></is></c>
      <c r="B3"><v>25</v></c>
    </row>
  </sheetData>
</worksheet>"#;

        // `worksheetSource/@name` points at the defined name instead of an A1 range.
        let cache_definition_xml = r#"<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<pivotCacheDefinition xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" recordCount="0">
  <cacheSource type="worksheet">
    <worksheetSource name="MyNamedRange"/>
  </cacheSource>
  <cacheFields count="1">
    <cacheField name="OldField" numFmtId="0"/>
  </cacheFields>
</pivotCacheDefinition>"#;

        let cache_definition_rels = r#"<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
  <Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/pivotCacheRecords" Target="pivotCacheRecords1.xml"/>
</Relationships>"#;

        let cache_records_xml = r#"<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<pivotCacheRecords xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" count="0"/>"#;

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

        zip.start_file("xl/pivotCache/pivotCacheDefinition1.xml", options)
            .unwrap();
        zip.write_all(cache_definition_xml.as_bytes()).unwrap();

        zip.start_file(
            "xl/pivotCache/_rels/pivotCacheDefinition1.xml.rels",
            options,
        )
        .unwrap();
        zip.write_all(cache_definition_rels.as_bytes()).unwrap();

        zip.start_file("xl/pivotCache/pivotCacheRecords1.xml", options)
            .unwrap();
        zip.write_all(cache_records_xml.as_bytes()).unwrap();

        let bytes = zip.finish().unwrap().into_inner();
        let mut pkg = XlsxPackage::from_bytes(&bytes).expect("read pkg");
        pkg.refresh_pivot_cache_from_worksheet("xl/pivotCache/pivotCacheDefinition1.xml")
            .expect("refresh");

        let updated_def =
            std::str::from_utf8(pkg.part("xl/pivotCache/pivotCacheDefinition1.xml").unwrap())
                .unwrap();
        let doc = Document::parse(updated_def).expect("parse updated cache definition");
        let root = doc.root_element();
        assert_eq!(root.attribute("recordCount"), Some("2"));
        let cache_fields = root
            .descendants()
            .find(|n| n.is_element() && n.tag_name().name() == "cacheFields")
            .expect("cacheFields");
        assert_eq!(cache_fields.attribute("count"), Some("2"));
        let field_names: Vec<_> = cache_fields
            .children()
            .filter(|n| n.is_element() && n.tag_name().name() == "cacheField")
            .filter_map(|n| n.attribute("name"))
            .collect();
        assert_eq!(field_names, vec!["Name", "Age"]);

        let updated_records =
            std::str::from_utf8(pkg.part("xl/pivotCache/pivotCacheRecords1.xml").unwrap()).unwrap();
        let doc = Document::parse(updated_records).expect("parse updated cache records");
        let root = doc.root_element();
        assert_eq!(root.attribute("count"), Some("2"));

        let record_values: Vec<Vec<(String, String)>> = root
            .children()
            .filter(|n| n.is_element() && n.tag_name().name() == "r")
            .map(|r| {
                r.children()
                    .filter(|n| n.is_element())
                    .map(|n| {
                        (
                            n.tag_name().name().to_string(),
                            n.attribute("v").unwrap_or_default().to_string(),
                        )
                    })
                    .collect::<Vec<_>>()
            })
            .collect();
        assert_eq!(
            record_values,
            vec![
                vec![
                    ("s".to_string(), "Alice".to_string()),
                    ("n".to_string(), "30".to_string()),
                ],
                vec![
                    ("s".to_string(), "Bob".to_string()),
                    ("n".to_string(), "25".to_string()),
                ],
            ]
        );
    }

    #[test]
    fn refresh_pivot_cache_falls_back_to_conventional_sheet_part_when_workbook_rels_missing() {
        let workbook_xml = r#"<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<workbook xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main"
 xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">
  <sheets>
    <sheet name="Sheet1" sheetId="1" r:id="rId1"/>
  </sheets>
</workbook>"#;

        // Source range is A1:B2 (header + 1 record).
        let worksheet_xml = r#"<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main">
  <sheetData>
    <row r="1">
      <c r="A1" t="inlineStr"><is><t>Header1</t></is></c>
      <c r="B1" t="inlineStr"><is><t>Header2</t></is></c>
    </row>
    <row r="2">
      <c r="A2"><v>1</v></c>
      <c r="B2" t="inlineStr"><is><t>Alpha</t></is></c>
    </row>
  </sheetData>
</worksheet>"#;

        let cache_definition_xml = r#"<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<pivotCacheDefinition xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" recordCount="0">
  <cacheSource type="worksheet">
    <worksheetSource sheet="Sheet1" ref="A1:B2"/>
  </cacheSource>
  <cacheFields count="0"/>
</pivotCacheDefinition>"#;

        let cache_definition_rels = r#"<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
  <Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/pivotCacheRecords" Target="pivotCacheRecords1.xml"/>
</Relationships>"#;

        let cache_records_xml = r#"<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<pivotCacheRecords xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" count="0"/>"#;

        let cursor = Cursor::new(Vec::new());
        let mut zip = zip::ZipWriter::new(cursor);
        let options = zip::write::FileOptions::<()>::default()
            .compression_method(zip::CompressionMethod::Deflated);

        zip.start_file("xl/workbook.xml", options).unwrap();
        zip.write_all(workbook_xml.as_bytes()).unwrap();

        // Intentionally omit `xl/_rels/workbook.xml.rels` to exercise the worksheet part fallback.
        zip.start_file("xl/worksheets/sheet1.xml", options).unwrap();
        zip.write_all(worksheet_xml.as_bytes()).unwrap();

        zip.start_file("xl/pivotCache/pivotCacheDefinition1.xml", options)
            .unwrap();
        zip.write_all(cache_definition_xml.as_bytes()).unwrap();

        zip.start_file(
            "xl/pivotCache/_rels/pivotCacheDefinition1.xml.rels",
            options,
        )
        .unwrap();
        zip.write_all(cache_definition_rels.as_bytes()).unwrap();

        zip.start_file("xl/pivotCache/pivotCacheRecords1.xml", options)
            .unwrap();
        zip.write_all(cache_records_xml.as_bytes()).unwrap();

        let bytes = zip.finish().unwrap().into_inner();
        let mut pkg = XlsxPackage::from_bytes(&bytes).expect("read pkg");
        pkg.refresh_pivot_cache_from_worksheet("xl/pivotCache/pivotCacheDefinition1.xml")
            .expect("refresh");

        let updated_def =
            std::str::from_utf8(pkg.part("xl/pivotCache/pivotCacheDefinition1.xml").unwrap())
                .unwrap();
        let doc = Document::parse(updated_def).expect("parse updated cache definition");
        let root = doc.root_element();
        assert_eq!(root.attribute("recordCount"), Some("1"));

        let updated_records =
            std::str::from_utf8(pkg.part("xl/pivotCache/pivotCacheRecords1.xml").unwrap()).unwrap();
        let doc = Document::parse(updated_records).expect("parse updated cache records");
        let root = doc.root_element();
        assert_eq!(root.attribute("count"), Some("1"));

        let record = root
            .children()
            .find(|n| n.is_element() && n.tag_name().name() == "r")
            .expect("record row");

        let values: Vec<(String, String)> = record
            .children()
            .filter(|n| n.is_element())
            .map(|n| {
                (
                    n.tag_name().name().to_string(),
                    n.attribute("v").unwrap_or_default().to_string(),
                )
            })
            .collect();
        assert_eq!(
            values,
            vec![
                ("n".to_string(), "1".to_string()),
                ("s".to_string(), "Alpha".to_string())
            ]
        );
    }

    #[test]
    fn refresh_pivot_cache_falls_back_to_default_styles_and_shared_strings_when_workbook_rels_malformed(
    ) {
        let workbook_xml = r#"<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<workbook xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main"
 xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">
  <sheets>
    <sheet name="Sheet1" sheetId="1" r:id="rId1"/>
  </sheets>
</workbook>"#;

        // Malformed: triggers `parse_relationships` failure.
        let workbook_rels = "<Relationships";

        let shared_strings_xml = r#"<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<sst xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" count="4" uniqueCount="4">
  <si><t>Date</t></si>
  <si><t>Name</t></si>
  <si><t>Alpha</t></si>
  <si><t>Beta</t></si>
</sst>"#;

        // Custom numFmtId so date detection depends on styles.xml (not built-in formats).
        let styles_xml = r#"<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<styleSheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main">
  <numFmts count="1">
    <numFmt numFmtId="164" formatCode="m/d/yyyy"/>
  </numFmts>
</styleSheet>"#;

        // Table range is A1:B3 (header + 2 records).
        // Header + string values use shared strings (`t="s"`), so refresh requires sharedStrings.xml.
        let worksheet_xml = r#"<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main"
 xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">
  <sheetData>
    <row r="1">
      <c r="A1" t="s"><v>0</v></c>
      <c r="B1" t="s"><v>1</v></c>
    </row>
    <row r="2">
      <c r="A2"><v>45123</v></c>
      <c r="B2" t="s"><v>2</v></c>
    </row>
    <row r="3">
      <c r="A3"><v>45124</v></c>
      <c r="B3" t="s"><v>3</v></c>
    </row>
  </sheetData>
  <tableParts count="1">
    <tablePart r:id="rIdTable1"/>
  </tableParts>
</worksheet>"#;

        let worksheet_rels = r#"<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
  <Relationship Id="rIdTable1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/table" Target="../tables/table1.xml"/>
</Relationships>"#;

        let table_xml = r#"<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<table xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" id="1" name="Table1" displayName="Table1" ref="A1:B3" headerRowCount="1" totalsRowCount="0">
  <tableColumns count="2">
    <tableColumn id="1" name="Date"/>
    <tableColumn id="2" name="Name"/>
  </tableColumns>
</table>"#;

        // `worksheetSource/@ref` points at the table name (not an A1 range) so worksheet lookup
        // does not depend on workbook.xml.rels.
        let cache_definition_xml = r#"<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<pivotCacheDefinition xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" recordCount="0">
  <cacheSource type="worksheet">
    <worksheetSource ref="Table1"/>
  </cacheSource>
  <cacheFields count="2">
    <cacheField name="OldDate" numFmtId="164"/>
    <cacheField name="OldName" numFmtId="0"/>
  </cacheFields>
</pivotCacheDefinition>"#;

        let cache_definition_rels = r#"<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
  <Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/pivotCacheRecords" Target="pivotCacheRecords1.xml"/>
</Relationships>"#;

        let cache_records_xml = r#"<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<pivotCacheRecords xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" count="0"/>"#;

        let cursor = Cursor::new(Vec::new());
        let mut zip = zip::ZipWriter::new(cursor);
        let options = zip::write::FileOptions::<()>::default()
            .compression_method(zip::CompressionMethod::Deflated);

        zip.start_file("xl/workbook.xml", options).unwrap();
        zip.write_all(workbook_xml.as_bytes()).unwrap();

        zip.start_file("xl/_rels/workbook.xml.rels", options)
            .unwrap();
        zip.write_all(workbook_rels.as_bytes()).unwrap();

        zip.start_file("xl/sharedStrings.xml", options).unwrap();
        zip.write_all(shared_strings_xml.as_bytes()).unwrap();

        zip.start_file("xl/styles.xml", options).unwrap();
        zip.write_all(styles_xml.as_bytes()).unwrap();

        zip.start_file("xl/worksheets/sheet1.xml", options).unwrap();
        zip.write_all(worksheet_xml.as_bytes()).unwrap();

        zip.start_file("xl/worksheets/_rels/sheet1.xml.rels", options)
            .unwrap();
        zip.write_all(worksheet_rels.as_bytes()).unwrap();

        zip.start_file("xl/tables/table1.xml", options).unwrap();
        zip.write_all(table_xml.as_bytes()).unwrap();

        zip.start_file("xl/pivotCache/pivotCacheDefinition1.xml", options)
            .unwrap();
        zip.write_all(cache_definition_xml.as_bytes()).unwrap();

        zip.start_file(
            "xl/pivotCache/_rels/pivotCacheDefinition1.xml.rels",
            options,
        )
        .unwrap();
        zip.write_all(cache_definition_rels.as_bytes()).unwrap();

        zip.start_file("xl/pivotCache/pivotCacheRecords1.xml", options)
            .unwrap();
        zip.write_all(cache_records_xml.as_bytes()).unwrap();

        let bytes = zip.finish().unwrap().into_inner();
        let mut pkg = XlsxPackage::from_bytes(&bytes).expect("read pkg");

        pkg.refresh_pivot_cache_from_worksheet("xl/pivotCache/pivotCacheDefinition1.xml")
            .expect("refresh");

        let updated_def =
            std::str::from_utf8(pkg.part("xl/pivotCache/pivotCacheDefinition1.xml").unwrap())
                .unwrap();
        let doc = Document::parse(updated_def).expect("parse updated cache definition");
        let root = doc.root_element();
        assert_eq!(root.attribute("recordCount"), Some("2"));
        let cache_fields = root
            .descendants()
            .find(|n| n.is_element() && n.tag_name().name() == "cacheFields")
            .expect("cacheFields");
        let field_names: Vec<_> = cache_fields
            .children()
            .filter(|n| n.is_element() && n.tag_name().name() == "cacheField")
            .filter_map(|n| n.attribute("name"))
            .collect();
        assert_eq!(field_names, vec!["Date", "Name"]);

        let updated_records =
            std::str::from_utf8(pkg.part("xl/pivotCache/pivotCacheRecords1.xml").unwrap()).unwrap();
        let doc = Document::parse(updated_records).expect("parse updated cache records");
        let root = doc.root_element();
        assert_eq!(root.attribute("count"), Some("2"));
        let records: Vec<_> = root
            .children()
            .filter(|n| n.is_element() && n.tag_name().name() == "r")
            .collect();
        assert_eq!(records.len(), 2);

        let expected_dt =
            excel_serial_to_pivot_datetime(45123.0, DateSystem::V1900).expect("serial");
        let first_values: Vec<_> = records[0].children().filter(|n| n.is_element()).collect();
        assert_eq!(first_values.len(), 2);
        assert_eq!(first_values[0].tag_name().name(), "d");
        assert_eq!(first_values[0].attribute("v"), Some(expected_dt.as_str()));
        assert_eq!(first_values[1].tag_name().name(), "s");
        assert_eq!(first_values[1].attribute("v"), Some("Alpha"));
    }

    #[test]
    fn split_sheet_qualified_reference_handles_unicode_quoted_sheet_names() {
        let (sheet, rest) = split_sheet_qualified_reference("'Straße'!$A$1");
        assert_eq!(sheet.as_deref(), Some("Straße"));
        assert_eq!(rest, "$A$1");
    }

    #[test]
    fn resolve_defined_name_reference_matches_unicode_case_insensitive_name_and_sheet_hint() {
        // Two sheet-scoped defined names share the same name; Excel selects the one matching the
        // worksheetSource/@sheet hint (case-insensitive across Unicode).
        let workbook_xml = r#"<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<workbook xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main"
 xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">
  <sheets>
    <sheet name="Sheet1" sheetId="1" r:id="rId1"/>
    <sheet name="Straße" sheetId="2" r:id="rId2"/>
  </sheets>
  <definedNames>
    <definedName name="Straße" localSheetId="0">A1</definedName>
    <definedName name="Straße" localSheetId="1">B2</definedName>
  </definedNames>
</workbook>"#;

        let cursor = Cursor::new(Vec::new());
        let mut zip = zip::ZipWriter::new(cursor);
        let options = zip::write::FileOptions::<()>::default()
            .compression_method(zip::CompressionMethod::Deflated);

        zip.start_file("xl/workbook.xml", options).unwrap();
        zip.write_all(workbook_xml.as_bytes()).unwrap();

        let bytes = zip.finish().unwrap().into_inner();
        let pkg = XlsxPackage::from_bytes(&bytes).expect("read pkg");

        let resolved =
            resolve_defined_name_reference(&pkg, "STRASSE", Some("STRASSE")).expect("resolve");
        let (sheet, range) = resolved.expect("expected match");
        assert_eq!(sheet.as_deref(), Some("Straße"));
        assert_eq!(range, Range::from_a1("B2").unwrap());
    }

    #[test]
    fn resolve_worksheet_part_matches_unicode_sheet_names_case_insensitive_like_excel() {
        let workbook_xml = r#"<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<workbook xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main"
 xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">
  <sheets>
    <sheet name="Straße" sheetId="1" r:id="rId1"/>
  </sheets>
</workbook>"#;

        let workbook_rels = r#"<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
  <Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/worksheet" Target="worksheets/sheet1.xml"/>
</Relationships>"#;

        let worksheet_xml = r#"<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main">
  <sheetData/>
</worksheet>"#;

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

        let bytes = zip.finish().unwrap().into_inner();
        let pkg = XlsxPackage::from_bytes(&bytes).expect("read pkg");

        let part = resolve_worksheet_part(&pkg, "STRASSE").expect("resolve worksheet");
        assert_eq!(part, "xl/worksheets/sheet1.xml");
    }

    #[test]
    fn resolve_worksheet_part_falls_back_to_guess_when_workbook_rels_targets_missing_part() {
        let workbook_xml = r#"<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<workbook xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main"
 xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">
  <sheets>
    <sheet name="Sheet1" sheetId="1" r:id="rId1"/>
  </sheets>
</workbook>"#;

        // Relationship is present but points to a missing worksheet part.
        let workbook_rels = r#"<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
  <Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/worksheet" Target="worksheets/missing.xml"/>
</Relationships>"#;

        // The conventional `sheet{sheetId}.xml` part exists and should be used as a fallback.
        let worksheet_xml = r#"<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main">
  <sheetData/>
</worksheet>"#;

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

        let bytes = zip.finish().unwrap().into_inner();
        let pkg = XlsxPackage::from_bytes(&bytes).expect("read pkg");

        let part = resolve_worksheet_part(&pkg, "Sheet1").expect("resolve worksheet");
        assert_eq!(part, "xl/worksheets/sheet1.xml");
    }

    #[test]
    fn refresh_pivot_cache_falls_back_to_records_filename_when_rels_missing() {
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
</Relationships>"#;

        // Cache source range is A1:B3 (header + 2 records).
        let worksheet_xml = r#"<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main">
  <sheetData>
    <row r="1">
      <c r="A1" t="inlineStr"><is><t>Header1</t></is></c>
      <c r="B1" t="inlineStr"><is><t>Header2</t></is></c>
    </row>
    <row r="2">
      <c r="A2"><v>1</v></c>
      <c r="B2" t="inlineStr"><is><t>Alpha</t></is></c>
    </row>
    <row r="3">
      <c r="A3"><v>2</v></c>
      <c r="B3" t="inlineStr"><is><t>Beta</t></is></c>
    </row>
  </sheetData>
</worksheet>"#;

        let cache_definition_xml = r#"<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<pivotCacheDefinition xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" recordCount="0">
  <cacheSource type="worksheet">
    <worksheetSource sheet="Sheet1" ref="A1:B3"/>
  </cacheSource>
  <cacheFields count="0"/>
</pivotCacheDefinition>"#;

        let cache_records_xml = r#"<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<pivotCacheRecords xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" count="0"/>"#;

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

        zip.start_file("xl/pivotCache/pivotCacheDefinition1.xml", options)
            .unwrap();
        zip.write_all(cache_definition_xml.as_bytes()).unwrap();

        // Intentionally omit `xl/pivotCache/_rels/pivotCacheDefinition1.xml.rels` to ensure
        // `resolve_cache_records_part()` falls back to the conventional records filename.

        zip.start_file("xl/pivotCache/pivotCacheRecords1.xml", options)
            .unwrap();
        zip.write_all(cache_records_xml.as_bytes()).unwrap();

        let bytes = zip.finish().unwrap().into_inner();
        let mut pkg = XlsxPackage::from_bytes(&bytes).expect("read pkg");
        pkg.refresh_pivot_cache_from_worksheet("xl/pivotCache/pivotCacheDefinition1.xml")
            .expect("refresh");

        let updated_records =
            std::str::from_utf8(pkg.part("xl/pivotCache/pivotCacheRecords1.xml").unwrap()).unwrap();
        let doc = Document::parse(updated_records).expect("parse updated cache records");
        let root = doc.root_element();
        assert_eq!(root.attribute("count"), Some("2"));
        let record_count = root
            .children()
            .filter(|n| n.is_element() && n.tag_name().name() == "r")
            .count();
        assert_eq!(record_count, 2);
    }

    #[test]
    fn refresh_pivot_cache_falls_back_to_records_filename_when_rels_malformed() {
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
</Relationships>"#;

        // Cache source range is A1:B3 (header + 2 records).
        let worksheet_xml = r#"<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main">
  <sheetData>
    <row r="1">
      <c r="A1" t="inlineStr"><is><t>Header1</t></is></c>
      <c r="B1" t="inlineStr"><is><t>Header2</t></is></c>
    </row>
    <row r="2">
      <c r="A2"><v>1</v></c>
      <c r="B2" t="inlineStr"><is><t>Alpha</t></is></c>
    </row>
    <row r="3">
      <c r="A3"><v>2</v></c>
      <c r="B3" t="inlineStr"><is><t>Beta</t></is></c>
    </row>
  </sheetData>
</worksheet>"#;

        let cache_definition_xml = r#"<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<pivotCacheDefinition xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" recordCount="0">
  <cacheSource type="worksheet">
    <worksheetSource sheet="Sheet1" ref="A1:B3"/>
  </cacheSource>
  <cacheFields count="0"/>
</pivotCacheDefinition>"#;

        let cache_records_xml = r#"<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<pivotCacheRecords xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" count="0"/>"#;

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

        zip.start_file("xl/pivotCache/pivotCacheDefinition1.xml", options)
            .unwrap();
        zip.write_all(cache_definition_xml.as_bytes()).unwrap();

        zip.start_file(
            "xl/pivotCache/_rels/pivotCacheDefinition1.xml.rels",
            options,
        )
        .unwrap();
        // Intentionally malformed.
        zip.write_all(b"this is not xml").unwrap();

        zip.start_file("xl/pivotCache/pivotCacheRecords1.xml", options)
            .unwrap();
        zip.write_all(cache_records_xml.as_bytes()).unwrap();

        let bytes = zip.finish().unwrap().into_inner();
        let mut pkg = XlsxPackage::from_bytes(&bytes).expect("read pkg");
        pkg.refresh_pivot_cache_from_worksheet("xl/pivotCache/pivotCacheDefinition1.xml")
            .expect("refresh");

        let updated_records =
            std::str::from_utf8(pkg.part("xl/pivotCache/pivotCacheRecords1.xml").unwrap()).unwrap();
        let doc = Document::parse(updated_records).expect("parse updated cache records");
        let root = doc.root_element();
        assert_eq!(root.attribute("count"), Some("2"));
        let record_count = root
            .children()
            .filter(|n| n.is_element() && n.tag_name().name() == "r")
            .count();
        assert_eq!(record_count, 2);
    }

    #[test]
    fn refresh_pivot_cache_falls_back_to_records_filename_when_rels_target_missing() {
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
</Relationships>"#;

        // Cache source range is A1:B3 (header + 2 records).
        let worksheet_xml = r#"<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main">
  <sheetData>
    <row r="1">
      <c r="A1" t="inlineStr"><is><t>Header1</t></is></c>
      <c r="B1" t="inlineStr"><is><t>Header2</t></is></c>
    </row>
    <row r="2">
      <c r="A2"><v>1</v></c>
      <c r="B2" t="inlineStr"><is><t>Alpha</t></is></c>
    </row>
    <row r="3">
      <c r="A3"><v>2</v></c>
      <c r="B3" t="inlineStr"><is><t>Beta</t></is></c>
    </row>
  </sheetData>
</worksheet>"#;

        let cache_definition_xml = r#"<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<pivotCacheDefinition xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" recordCount="0">
  <cacheSource type="worksheet">
    <worksheetSource sheet="Sheet1" ref="A1:B3"/>
  </cacheSource>
  <cacheFields count="0"/>
</pivotCacheDefinition>"#;

        // Relationship part exists, but points to a non-existent records part.
        let cache_definition_rels = r#"<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
  <Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/pivotCacheRecords" Target="pivotCacheRecords999.xml"/>
</Relationships>"#;

        let cache_records_xml = r#"<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<pivotCacheRecords xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" count="0"/>"#;

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

        zip.start_file("xl/pivotCache/pivotCacheDefinition1.xml", options)
            .unwrap();
        zip.write_all(cache_definition_xml.as_bytes()).unwrap();

        zip.start_file(
            "xl/pivotCache/_rels/pivotCacheDefinition1.xml.rels",
            options,
        )
        .unwrap();
        zip.write_all(cache_definition_rels.as_bytes()).unwrap();

        // Store the records file under the conventional name (the rels points elsewhere).
        zip.start_file("xl/pivotCache/pivotCacheRecords1.xml", options)
            .unwrap();
        zip.write_all(cache_records_xml.as_bytes()).unwrap();

        let bytes = zip.finish().unwrap().into_inner();
        let mut pkg = XlsxPackage::from_bytes(&bytes).expect("read pkg");
        pkg.refresh_pivot_cache_from_worksheet("xl/pivotCache/pivotCacheDefinition1.xml")
            .expect("refresh");

        let updated_records =
            std::str::from_utf8(pkg.part("xl/pivotCache/pivotCacheRecords1.xml").unwrap()).unwrap();
        let doc = Document::parse(updated_records).expect("parse updated cache records");
        let root = doc.root_element();
        assert_eq!(root.attribute("count"), Some("2"));
        let record_count = root
            .children()
            .filter(|n| n.is_element() && n.tag_name().name() == "r")
            .count();
        assert_eq!(record_count, 2);
    }
}
