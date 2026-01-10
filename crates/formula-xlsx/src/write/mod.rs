use std::collections::{BTreeMap, HashMap};
use std::io::{Cursor, Write};

use formula_model::{CellRef, CellValue, ErrorValue, Range, Worksheet};
use quick_xml::events::Event;
use quick_xml::events::attributes::AttrError;
use quick_xml::Reader;
use quick_xml::Writer;
use thiserror::Error;
use zip::write::FileOptions;
use zip::ZipWriter;

use crate::{CellValueKind, DateSystem, SheetMeta, XlsxDocument};

#[derive(Debug, Error)]
pub enum WriteError {
    #[error("io error: {0}")]
    Io(#[from] std::io::Error),
    #[error("zip error: {0}")]
    Zip(#[from] zip::result::ZipError),
    #[error("xml error: {0}")]
    Xml(#[from] quick_xml::Error),
    #[error("xml attribute error: {0}")]
    XmlAttr(#[from] AttrError),
}

pub fn write_to_vec(doc: &XlsxDocument) -> Result<Vec<u8>, WriteError> {
    let mut parts = build_parts(doc)?;

    // Deterministic ordering helps debugging and makes fixtures stable.
    let cursor = Cursor::new(Vec::new());
    let mut zip = ZipWriter::new(cursor);
    let options = FileOptions::<()>::default().compression_method(zip::CompressionMethod::Deflated);

    for (name, bytes) in parts.iter_mut() {
        zip.start_file(name, options)?;
        zip.write_all(bytes)?;
    }

    let cursor = zip.finish()?;
    Ok(cursor.into_inner())
}

fn build_parts(doc: &XlsxDocument) -> Result<BTreeMap<String, Vec<u8>>, WriteError> {
    let mut parts = doc.parts.clone();
    let is_new = parts.is_empty();
    if is_new {
        parts = generate_minimal_package(doc)?;
    }

    let (shared_strings_xml, shared_string_lookup) = build_shared_strings_xml(doc)?;
    if is_new || !shared_string_lookup.is_empty() || parts.contains_key("xl/sharedStrings.xml") {
        parts.insert("xl/sharedStrings.xml".to_string(), shared_strings_xml);
    }

    // styles.xml is required by most files even if we don't model it.
    if !parts.contains_key("xl/styles.xml") {
        parts.insert("xl/styles.xml".to_string(), minimal_styles_xml());
    }

    // Ensure core relationship/content types metadata exists when we synthesize new
    // parts for existing packages. For existing relationships we preserve IDs by
    // only adding missing entries with a new `rIdN`.
    if parts.contains_key("xl/sharedStrings.xml") {
        ensure_content_types_override(
            &mut parts,
            "/xl/sharedStrings.xml",
            "application/vnd.openxmlformats-officedocument.spreadsheetml.sharedStrings+xml",
        )?;
        ensure_workbook_rels_has_relationship(
            &mut parts,
            "http://schemas.openxmlformats.org/officeDocument/2006/relationships/sharedStrings",
            "sharedStrings.xml",
        )?;
    }
    if parts.contains_key("xl/styles.xml") {
        ensure_content_types_override(
            &mut parts,
            "/xl/styles.xml",
            "application/vnd.openxmlformats-officedocument.spreadsheetml.styles+xml",
        )?;
        ensure_workbook_rels_has_relationship(
            &mut parts,
            "http://schemas.openxmlformats.org/officeDocument/2006/relationships/styles",
            "styles.xml",
        )?;
    }

    let workbook_orig = parts.get("xl/workbook.xml").map(|b| b.as_slice());
    parts.insert(
        "xl/workbook.xml".to_string(),
        write_workbook_xml(doc, workbook_orig)?,
    );

    for sheet_meta in &doc.meta.sheets {
        let sheet = doc
            .workbook
            .sheet(sheet_meta.worksheet_id)
            .ok_or_else(|| WriteError::Io(std::io::Error::new(std::io::ErrorKind::NotFound, "worksheet not found")))?;
        let orig = parts.get(&sheet_meta.path).map(|b| b.as_slice());
        parts.insert(
            sheet_meta.path.clone(),
            write_worksheet_xml(doc, sheet_meta, sheet, orig, &shared_string_lookup)?,
        );
    }

    Ok(parts)
}

fn build_shared_strings_xml(
    doc: &XlsxDocument,
) -> Result<(Vec<u8>, HashMap<String, u32>), WriteError> {
    let mut table: Vec<String> = doc.shared_strings.clone();
    let mut lookup: HashMap<String, u32> = HashMap::new();
    for (idx, s) in table.iter().enumerate() {
        lookup.entry(s.clone()).or_insert(idx as u32);
    }

    let mut ref_count: u32 = 0;

    for sheet_meta in &doc.meta.sheets {
        let sheet = match doc.workbook.sheet(sheet_meta.worksheet_id) {
            Some(s) => s,
            None => continue,
        };

        let mut cells: Vec<(CellRef, &formula_model::Cell)> = sheet.iter_cells().collect();
        cells.sort_by_key(|(r, _)| (r.row, r.col));
        for (cell_ref, cell) in cells {
            let meta = doc.meta.cell_meta.get(&(sheet_meta.worksheet_id, cell_ref));
            let kind = meta.and_then(|m| m.value_kind.clone()).unwrap_or_else(|| infer_value_kind(cell));
            if let CellValueKind::SharedString { index } = kind {
                if let CellValue::String(text) = &cell.value {
                    ref_count += 1;
                    if table.get(index as usize).map(|s| s.as_str()) == Some(text.as_str()) {
                        lookup.entry(text.clone()).or_insert(index);
                        continue;
                    }
                    if !lookup.contains_key(text) {
                        let new_index = table.len() as u32;
                        table.push(text.clone());
                        lookup.insert(text.clone(), new_index);
                    }
                }
            }
        }
    }

    let mut xml = String::new();
    xml.push_str(r#"<?xml version="1.0" encoding="UTF-8" standalone="yes"?>"#);
    xml.push_str(r#"<sst xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main""#);
    xml.push_str(&format!(r#" count="{ref_count}" uniqueCount="{}">"#, table.len()));
    for s in &table {
        xml.push_str("<si><t");
        if needs_space_preserve(s) {
            xml.push_str(r#" xml:space="preserve""#);
        }
        xml.push('>');
        xml.push_str(&escape_text(s));
        xml.push_str("</t></si>");
    }
    xml.push_str("</sst>");

    Ok((xml.into_bytes(), lookup))
}

fn needs_space_preserve(s: &str) -> bool {
    s.starts_with(char::is_whitespace) || s.ends_with(char::is_whitespace)
}

fn escape_text(s: &str) -> String {
    s.replace('&', "&amp;")
        .replace('<', "&lt;")
        .replace('>', "&gt;")
}

fn escape_attr(s: &str) -> String {
    escape_text(s)
        .replace('\"', "&quot;")
        .replace('\'', "&apos;")
}

fn write_workbook_xml(doc: &XlsxDocument, original: Option<&[u8]>) -> Result<Vec<u8>, WriteError> {
    if let Some(original) = original {
        return patch_workbook_xml(doc, original);
    }

    let mut xml = String::new();
    xml.push_str(r#"<?xml version="1.0" encoding="UTF-8" standalone="yes"?>"#);
    xml.push_str(
        r#"<workbook xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">"#,
    );
    xml.push_str("<workbookPr");
    if doc.meta.date_system == DateSystem::V1904 {
        xml.push_str(r#" date1904="1""#);
    }
    xml.push_str("/>");
    xml.push_str("<sheets>");
    for sheet_meta in &doc.meta.sheets {
        let name = doc
            .workbook
            .sheet(sheet_meta.worksheet_id)
            .map(|s| s.name.as_str())
            .unwrap_or("Sheet");
        xml.push_str("<sheet");
        xml.push_str(&format!(r#" name="{}""#, escape_attr(name)));
        xml.push_str(&format!(r#" sheetId="{}""#, sheet_meta.sheet_id));
        xml.push_str(&format!(
            r#" r:id="{}""#,
            escape_attr(&sheet_meta.relationship_id)
        ));
        if let Some(state) = &sheet_meta.state {
            xml.push_str(&format!(r#" state="{}""#, escape_attr(state)));
        }
        xml.push_str("/>");
    }
    xml.push_str("</sheets>");
    xml.push_str("</workbook>");
    Ok(xml.into_bytes())
}

fn patch_workbook_xml(doc: &XlsxDocument, original: &[u8]) -> Result<Vec<u8>, WriteError> {
    let mut reader = Reader::from_reader(original);
    reader.config_mut().trim_text(false);
    let mut buf = Vec::new();
    let mut writer = Writer::new(Vec::with_capacity(original.len()));

    let mut skipping_sheets = false;
    let mut skipping_workbook_pr = false;
    let mut skipping_calc_pr = false;
    loop {
        match reader.read_event_into(&mut buf)? {
            Event::Start(e) if e.name().as_ref() == b"workbookPr" => {
                skipping_workbook_pr = true;
                let empty = Event::Empty(e.into_owned());
                match empty {
                    Event::Empty(e) => write_workbook_pr(doc, &mut writer, &e)?,
                    _ => unreachable!(),
                }
            }
            Event::Empty(e) if e.name().as_ref() == b"workbookPr" => write_workbook_pr(doc, &mut writer, &e)?,
            Event::End(e) if e.name().as_ref() == b"workbookPr" => {
                if skipping_workbook_pr {
                    skipping_workbook_pr = false;
                } else {
                    writer.write_event(Event::End(e.into_owned()))?;
                }
            }

            Event::Start(e) if e.name().as_ref() == b"calcPr" => {
                skipping_calc_pr = true;
                let empty = Event::Empty(e.into_owned());
                match empty {
                    Event::Empty(e) => write_calc_pr(doc, &mut writer, &e)?,
                    _ => unreachable!(),
                }
            }
            Event::Empty(e) if e.name().as_ref() == b"calcPr" => write_calc_pr(doc, &mut writer, &e)?,
            Event::End(e) if e.name().as_ref() == b"calcPr" => {
                if skipping_calc_pr {
                    skipping_calc_pr = false;
                } else {
                    writer.write_event(Event::End(e.into_owned()))?;
                }
            }

            Event::Start(e) if e.name().as_ref() == b"sheets" => {
                skipping_sheets = true;
                writer.get_mut().extend_from_slice(b"<sheets");
                for attr in e.attributes() {
                    let attr = attr?;
                    writer.get_mut().push(b' ');
                    writer.get_mut().extend_from_slice(attr.key.as_ref());
                    writer.get_mut().extend_from_slice(b"=\"");
                    writer.get_mut().extend_from_slice(
                        escape_attr(&attr.unescape_value()?.into_owned()).as_bytes(),
                    );
                    writer.get_mut().push(b'"');
                }
                writer.get_mut().push(b'>');

                for sheet_meta in &doc.meta.sheets {
                    let name = doc
                        .workbook
                        .sheet(sheet_meta.worksheet_id)
                        .map(|s| s.name.as_str())
                        .unwrap_or("Sheet");
                    writer.get_mut().extend_from_slice(b"<sheet");
                    writer.get_mut().extend_from_slice(b" name=\"");
                    writer.get_mut().extend_from_slice(escape_attr(name).as_bytes());
                    writer.get_mut().push(b'"');
                    writer.get_mut().extend_from_slice(b" sheetId=\"");
                    writer
                        .get_mut()
                        .extend_from_slice(sheet_meta.sheet_id.to_string().as_bytes());
                    writer.get_mut().push(b'"');
                    writer.get_mut().extend_from_slice(b" r:id=\"");
                    writer.get_mut().extend_from_slice(
                        escape_attr(&sheet_meta.relationship_id).as_bytes(),
                    );
                    writer.get_mut().push(b'"');
                    if let Some(state) = &sheet_meta.state {
                        writer.get_mut().extend_from_slice(b" state=\"");
                        writer.get_mut().extend_from_slice(escape_attr(state).as_bytes());
                        writer.get_mut().push(b'"');
                    }
                    writer.get_mut().extend_from_slice(b"/>");
                }
            }
            Event::Empty(e) if e.name().as_ref() == b"sheets" => {
                // Replace `<sheets/>` with a full section.
                writer.get_mut().extend_from_slice(b"<sheets");
                for attr in e.attributes() {
                    let attr = attr?;
                    writer.get_mut().push(b' ');
                    writer.get_mut().extend_from_slice(attr.key.as_ref());
                    writer.get_mut().extend_from_slice(b"=\"");
                    writer.get_mut().extend_from_slice(
                        escape_attr(&attr.unescape_value()?.into_owned()).as_bytes(),
                    );
                    writer.get_mut().push(b'"');
                }
                writer.get_mut().push(b'>');
                for sheet_meta in &doc.meta.sheets {
                    let name = doc
                        .workbook
                        .sheet(sheet_meta.worksheet_id)
                        .map(|s| s.name.as_str())
                        .unwrap_or("Sheet");
                    writer.get_mut().extend_from_slice(b"<sheet");
                    writer.get_mut().extend_from_slice(b" name=\"");
                    writer.get_mut().extend_from_slice(escape_attr(name).as_bytes());
                    writer.get_mut().push(b'"');
                    writer.get_mut().extend_from_slice(b" sheetId=\"");
                    writer
                        .get_mut()
                        .extend_from_slice(sheet_meta.sheet_id.to_string().as_bytes());
                    writer.get_mut().push(b'"');
                    writer.get_mut().extend_from_slice(b" r:id=\"");
                    writer.get_mut().extend_from_slice(
                        escape_attr(&sheet_meta.relationship_id).as_bytes(),
                    );
                    writer.get_mut().push(b'"');
                    if let Some(state) = &sheet_meta.state {
                        writer.get_mut().extend_from_slice(b" state=\"");
                        writer.get_mut().extend_from_slice(escape_attr(state).as_bytes());
                        writer.get_mut().push(b'"');
                    }
                    writer.get_mut().extend_from_slice(b"/>");
                }
                writer.get_mut().extend_from_slice(b"</sheets>");
            }
            Event::End(e) if e.name().as_ref() == b"sheets" => {
                skipping_sheets = false;
                writer.get_mut().extend_from_slice(b"</sheets>");
            }

            Event::Eof => break,
            ev if skipping_workbook_pr || skipping_calc_pr => drop(ev),
            ev if skipping_sheets => drop(ev),
            ev => writer.write_event(ev.into_owned())?,
        }
        buf.clear();
    }

    Ok(writer.into_inner())
}

fn write_workbook_pr(
    doc: &XlsxDocument,
    writer: &mut Writer<Vec<u8>>,
    e: &quick_xml::events::BytesStart<'_>,
) -> Result<(), WriteError> {
    let had_date1904 = e
        .attributes()
        .flatten()
        .any(|a| a.key.as_ref() == b"date1904");

    writer.get_mut().extend_from_slice(b"<workbookPr");
    for attr in e.attributes() {
        let attr = attr?;
        if attr.key.as_ref() == b"date1904" {
            continue;
        }
        writer.get_mut().push(b' ');
        writer.get_mut().extend_from_slice(attr.key.as_ref());
        writer.get_mut().extend_from_slice(b"=\"");
        writer
            .get_mut()
            .extend_from_slice(escape_attr(&attr.unescape_value()?.into_owned()).as_bytes());
        writer.get_mut().push(b'"');
    }

    if doc.meta.date_system == DateSystem::V1904 {
        writer.get_mut().extend_from_slice(b" date1904=\"1\"");
    } else if had_date1904 {
        writer.get_mut().extend_from_slice(b" date1904=\"0\"");
    }
    writer.get_mut().extend_from_slice(b"/>");
    Ok(())
}

fn write_calc_pr(
    doc: &XlsxDocument,
    writer: &mut Writer<Vec<u8>>,
    e: &quick_xml::events::BytesStart<'_>,
) -> Result<(), WriteError> {
    writer.get_mut().extend_from_slice(b"<calcPr");
    for attr in e.attributes() {
        let attr = attr?;
        match attr.key.as_ref() {
            b"calcId" | b"calcMode" | b"fullCalcOnLoad" => continue,
            _ => {}
        }
        writer.get_mut().push(b' ');
        writer.get_mut().extend_from_slice(attr.key.as_ref());
        writer.get_mut().extend_from_slice(b"=\"");
        writer
            .get_mut()
            .extend_from_slice(escape_attr(&attr.unescape_value()?.into_owned()).as_bytes());
        writer.get_mut().push(b'"');
    }

    if let Some(calc_id) = &doc.meta.calc_pr.calc_id {
        writer.get_mut().extend_from_slice(b" calcId=\"");
        writer.get_mut().extend_from_slice(escape_attr(calc_id).as_bytes());
        writer.get_mut().push(b'"');
    }
    if let Some(calc_mode) = &doc.meta.calc_pr.calc_mode {
        writer.get_mut().extend_from_slice(b" calcMode=\"");
        writer.get_mut().extend_from_slice(escape_attr(calc_mode).as_bytes());
        writer.get_mut().push(b'"');
    }
    if let Some(full) = doc.meta.calc_pr.full_calc_on_load {
        writer.get_mut().extend_from_slice(b" fullCalcOnLoad=\"");
        writer.get_mut().extend_from_slice(if full { b"1" } else { b"0" });
        writer.get_mut().push(b'"');
    }
    writer.get_mut().extend_from_slice(b"/>");
    Ok(())
}

fn write_worksheet_xml(
    doc: &XlsxDocument,
    sheet_meta: &SheetMeta,
    sheet: &Worksheet,
    original: Option<&[u8]>,
    shared_lookup: &HashMap<String, u32>,
) -> Result<Vec<u8>, WriteError> {
    if let Some(original) = original {
        return patch_worksheet_xml(doc, sheet_meta, sheet, original, shared_lookup);
    }

    let dimension = sheet
        .used_range()
        .unwrap_or(Range::new(CellRef::new(0, 0), CellRef::new(0, 0)))
        .to_string();
    let sheet_data_xml = render_sheet_data(doc, sheet_meta, sheet, shared_lookup);

    let mut xml = String::new();
    xml.push_str(r#"<?xml version="1.0" encoding="UTF-8" standalone="yes"?>"#);
    xml.push_str(
        r#"<worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">"#,
    );
    xml.push_str(&format!(r#"<dimension ref="{dimension}"/>"#));
    xml.push_str(&sheet_data_xml);
    xml.push_str("</worksheet>");
    Ok(xml.into_bytes())
}

fn patch_worksheet_xml(
    doc: &XlsxDocument,
    sheet_meta: &SheetMeta,
    sheet: &Worksheet,
    original: &[u8],
    shared_lookup: &HashMap<String, u32>,
) -> Result<Vec<u8>, WriteError> {
    let sheet_data_xml = render_sheet_data(doc, sheet_meta, sheet, shared_lookup);

    let mut reader = Reader::from_reader(original);
    reader.config_mut().trim_text(false);
    let mut buf = Vec::new();
    let mut writer = Writer::new(Vec::with_capacity(original.len() + sheet_data_xml.len()));

    let mut skipping_sheet_data = false;
    loop {
        match reader.read_event_into(&mut buf)? {
            Event::Start(e) if e.name().as_ref() == b"sheetData" => {
                skipping_sheet_data = true;
                writer.get_mut().extend_from_slice(sheet_data_xml.as_bytes());
            }
            Event::Empty(e) if e.name().as_ref() == b"sheetData" => {
                writer.get_mut().extend_from_slice(sheet_data_xml.as_bytes());
                drop(e);
            }
            Event::End(e) if e.name().as_ref() == b"sheetData" => {
                skipping_sheet_data = false;
                drop(e);
            }
            Event::Eof => break,
            ev if skipping_sheet_data => drop(ev),
            ev => writer.write_event(ev.into_owned())?,
        }
        buf.clear();
    }

    Ok(writer.into_inner())
}

fn render_sheet_data(
    doc: &XlsxDocument,
    sheet_meta: &SheetMeta,
    sheet: &Worksheet,
    shared_lookup: &HashMap<String, u32>,
) -> String {
    let mut out = String::new();
    out.push_str("<sheetData>");

    let mut cells: Vec<(CellRef, &formula_model::Cell)> = sheet.iter_cells().collect();
    cells.sort_by_key(|(r, _)| (r.row, r.col));

    let mut current_row: Option<u32> = None;
    for (cell_ref, cell) in cells {
        let row_1_based = cell_ref.row + 1;
        if current_row != Some(row_1_based) {
            if current_row.is_some() {
                out.push_str("</row>");
            }
            current_row = Some(row_1_based);
            out.push_str(&format!(r#"<row r="{row_1_based}">"#));
        }

        out.push_str(r#"<c r=""#);
        out.push_str(&cell_ref.to_a1());
        out.push('"');

        if cell.style_id != 0 {
            out.push_str(&format!(r#" s="{}""#, cell.style_id));
        }

        let meta = doc.meta.cell_meta.get(&(sheet_meta.worksheet_id, cell_ref));
        let value_kind = meta
            .and_then(|m| m.value_kind.clone())
            .unwrap_or_else(|| infer_value_kind(cell));

        if !matches!(cell.value, CellValue::Empty) {
            match value_kind {
                CellValueKind::SharedString { .. } => out.push_str(r#" t="s""#),
                CellValueKind::InlineString => out.push_str(r#" t="inlineStr""#),
                CellValueKind::Bool => out.push_str(r#" t="b""#),
                CellValueKind::Error => out.push_str(r#" t="e""#),
                CellValueKind::Str => out.push_str(r#" t="str""#),
                CellValueKind::Number => {}
            }
        }

        out.push('>');

        if let Some(formula_meta) = meta.and_then(|m| m.formula.clone()).or_else(|| {
            cell.formula
                .as_ref()
                .map(|f| crate::FormulaMeta { file_text: f.clone(), ..Default::default() })
        }) {
            out.push_str("<f");
            if let Some(t) = &formula_meta.t {
                out.push_str(&format!(r#" t="{}""#, escape_attr(t)));
            }
            if let Some(r) = &formula_meta.reference {
                out.push_str(&format!(r#" ref="{}""#, escape_attr(r)));
            }
            if let Some(si) = formula_meta.shared_index {
                out.push_str(&format!(r#" si="{si}""#));
            }
            if let Some(aca) = formula_meta.always_calc {
                out.push_str(&format!(r#" aca="{}""#, if aca { "1" } else { "0" }));
            }

            let file_text = formula_file_text(&formula_meta, cell.formula.as_deref());
            if file_text.is_empty() {
                out.push_str("/>");
            } else {
                out.push('>');
                out.push_str(&escape_text(&file_text));
                out.push_str("</f>");
            }
        }

        match &cell.value {
            CellValue::Empty => {}
            CellValue::Number(n) => {
                out.push_str("<v>");
                out.push_str(&escape_text(&raw_or_number(meta, *n)));
                out.push_str("</v>");
            }
            CellValue::Boolean(b) => {
                out.push_str("<v>");
                out.push_str(raw_or_bool(meta, *b));
                out.push_str("</v>");
            }
            CellValue::Error(err) => {
                out.push_str("<v>");
                out.push_str(&escape_text(&raw_or_error(meta, *err)));
                out.push_str("</v>");
            }
            CellValue::String(s) => match value_kind {
                CellValueKind::SharedString { index } => {
                    let idx = shared_string_index(doc, meta, s, index, shared_lookup);
                    out.push_str("<v>");
                    out.push_str(&idx.to_string());
                    out.push_str("</v>");
                }
                CellValueKind::InlineString => {
                    out.push_str("<is><t");
                    if needs_space_preserve(s) {
                        out.push_str(r#" xml:space="preserve""#);
                    }
                    out.push('>');
                    out.push_str(&escape_text(s));
                    out.push_str("</t></is>");
                }
                CellValueKind::Str => {
                    out.push_str("<v>");
                    out.push_str(&escape_text(&raw_or_str(meta, s)));
                    out.push_str("</v>");
                }
                _ => {
                    // Fallback: treat as shared string.
                    let idx = shared_string_index(doc, meta, s, 0, shared_lookup);
                    out.push_str("<v>");
                    out.push_str(&idx.to_string());
                    out.push_str("</v>");
                }
            },
            _ => {
                // TODO: RichText/Array/Spill not yet modeled for writing. Preserve as blank.
            }
        }

        out.push_str("</c>");
    }

    if current_row.is_some() {
        out.push_str("</row>");
    }
    out.push_str("</sheetData>");
    out
}

fn infer_value_kind(cell: &formula_model::Cell) -> CellValueKind {
    match &cell.value {
        CellValue::Boolean(_) => CellValueKind::Bool,
        CellValue::Error(_) => CellValueKind::Error,
        CellValue::Number(_) => CellValueKind::Number,
        CellValue::String(_) => CellValueKind::SharedString { index: 0 },
        CellValue::Empty => CellValueKind::Number,
        _ => CellValueKind::Number,
    }
}

fn formula_file_text(meta: &crate::FormulaMeta, display: Option<&str>) -> String {
    if meta.file_text.is_empty() {
        return String::new();
    }
    if let Some(display) = display {
        // Preserve stored file text if the model's display text matches.
        if strip_xlfn_prefixes(&meta.file_text) == display {
            return meta.file_text.clone();
        }
    }
    meta.file_text.clone()
}

fn strip_xlfn_prefixes(s: &str) -> String {
    s.replace("_xlfn.", "")
}

fn raw_or_number(meta: Option<&crate::CellMeta>, n: f64) -> String {
    if let Some(meta) = meta {
        if let Some(raw) = &meta.raw_value {
            if raw.parse::<f64>().ok() == Some(n) {
                return raw.clone();
            }
        }
    }
    // Default formatting is fine for deterministic output; raw_value is used to preserve
    // round-trip fidelity where available.
    n.to_string()
}

fn raw_or_bool(meta: Option<&crate::CellMeta>, b: bool) -> &'static str {
    if let Some(meta) = meta {
        if let Some(raw) = meta.raw_value.as_deref() {
            if (raw == "1" && b) || (raw == "0" && !b) {
                return if b { "1" } else { "0" };
            }
        }
    }
    if b { "1" } else { "0" }
}

fn raw_or_error(meta: Option<&crate::CellMeta>, err: ErrorValue) -> String {
    if let Some(meta) = meta {
        if let Some(raw) = &meta.raw_value {
            if raw == err.as_str() {
                return raw.clone();
            }
        }
    }
    err.as_str().to_string()
}

fn raw_or_str(meta: Option<&crate::CellMeta>, s: &str) -> String {
    if let Some(meta) = meta {
        if let Some(raw) = &meta.raw_value {
            if raw == s {
                return raw.clone();
            }
        }
    }
    s.to_string()
}

fn shared_string_index(
    doc: &XlsxDocument,
    meta: Option<&crate::CellMeta>,
    text: &str,
    fallback_index: u32,
    shared_lookup: &HashMap<String, u32>,
) -> u32 {
    if let Some(meta) = meta {
        if let Some(CellValueKind::SharedString { index }) = &meta.value_kind {
            if doc
                .shared_strings
                .get(*index as usize)
                .map(|s| s.as_str())
                == Some(text)
            {
                return *index;
            }
        }
    }
    shared_lookup.get(text).copied().unwrap_or(fallback_index)
}

fn minimal_styles_xml() -> Vec<u8> {
    br#"<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<styleSheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main">
  <fonts count="1"><font><sz val="11"/><color theme="1"/><name val="Calibri"/><family val="2"/><scheme val="minor"/></font></fonts>
  <fills count="2"><fill><patternFill patternType="none"/></fill><fill><patternFill patternType="gray125"/></fill></fills>
  <borders count="1"><border><left/><right/><top/><bottom/><diagonal/></border></borders>
  <cellStyleXfs count="1"><xf numFmtId="0" fontId="0" fillId="0" borderId="0"/></cellStyleXfs>
  <cellXfs count="1"><xf numFmtId="0" fontId="0" fillId="0" borderId="0" xfId="0"/></cellXfs>
  <cellStyles count="1"><cellStyle name="Normal" xfId="0" builtinId="0"/></cellStyles>
  <dxfs count="0"/>
  <tableStyles count="0" defaultTableStyle="TableStyleMedium9" defaultPivotStyle="PivotStyleLight16"/>
</styleSheet>
"#
    .to_vec()
}

fn generate_minimal_package(doc: &XlsxDocument) -> Result<BTreeMap<String, Vec<u8>>, WriteError> {
    let mut parts = BTreeMap::new();

    parts.insert(
        "_rels/.rels".to_string(),
        br#"<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
  <Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="xl/workbook.xml"/>
</Relationships>
"#
        .to_vec(),
    );

    // Minimal workbook relationships; existing packages preserve the original bytes.
    parts.insert(
        "xl/_rels/workbook.xml.rels".to_string(),
        minimal_workbook_rels_xml(doc).into_bytes(),
    );

    parts.insert(
        "[Content_Types].xml".to_string(),
        minimal_content_types_xml(doc).into_bytes(),
    );

    Ok(parts)
}

fn minimal_workbook_rels_xml(doc: &XlsxDocument) -> String {
    let mut xml = String::new();
    xml.push_str(r#"<?xml version="1.0" encoding="UTF-8" standalone="yes"?>"#);
    xml.push_str(r#"<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">"#);

    for sheet_meta in &doc.meta.sheets {
        let target = rels_target_from_part_path(&sheet_meta.path);
        xml.push_str(r#"<Relationship Id=""#);
        xml.push_str(&escape_attr(&sheet_meta.relationship_id));
        xml.push_str(r#"" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/worksheet" Target=""#);
        xml.push_str(&escape_attr(&target));
        xml.push_str(r#""/>"#);
    }

    let next = next_relationship_id(
        doc.meta
            .sheets
            .iter()
            .map(|s| s.relationship_id.as_str()),
    );
    xml.push_str(&format!(r#"<Relationship Id="rId{next}" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/styles" Target="styles.xml"/>"#));
    let next2 = next + 1;
    xml.push_str(&format!(r#"<Relationship Id="rId{next2}" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/sharedStrings" Target="sharedStrings.xml"/>"#));
    xml.push_str("</Relationships>");
    xml
}

fn rels_target_from_part_path(path: &str) -> String {
    // workbook.xml.rels is rooted at `xl/`, so worksheet targets are relative.
    path.strip_prefix("xl/")
        .or_else(|| path.strip_prefix("/xl/"))
        .unwrap_or(path)
        .to_string()
}

fn minimal_content_types_xml(doc: &XlsxDocument) -> String {
    let mut xml = String::new();
    xml.push_str(r#"<?xml version="1.0" encoding="UTF-8" standalone="yes"?>"#);
    xml.push_str(r#"<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">"#);
    xml.push_str(r#"<Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/>"#);
    xml.push_str(r#"<Default Extension="xml" ContentType="application/xml"/>"#);
    xml.push_str(r#"<Override PartName="/xl/workbook.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet.main+xml"/>"#);
    for sheet_meta in &doc.meta.sheets {
        xml.push_str(r#"<Override PartName="/"#);
        xml.push_str(&escape_attr(&sheet_meta.path));
        xml.push_str(r#"" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.worksheet+xml"/>"#);
    }
    xml.push_str(r#"<Override PartName="/xl/styles.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.styles+xml"/>"#);
    xml.push_str(r#"<Override PartName="/xl/sharedStrings.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.sharedStrings+xml"/>"#);
    xml.push_str("</Types>");
    xml
}

fn ensure_content_types_override(
    parts: &mut BTreeMap<String, Vec<u8>>,
    part_name: &str,
    content_type: &str,
) -> Result<(), WriteError> {
    let Some(existing) = parts.get("[Content_Types].xml").cloned() else {
        // Avoid synthesizing a full file for existing packages.
        return Ok(());
    };
    let mut xml = String::from_utf8(existing)
        .map_err(|e| WriteError::Io(std::io::Error::new(std::io::ErrorKind::InvalidData, e)))?;
    if xml.contains(&format!(r#"PartName="{part_name}""#)) {
        parts.insert("[Content_Types].xml".to_string(), xml.into_bytes());
        return Ok(());
    }
    if let Some(idx) = xml.rfind("</Types>") {
        let insert = format!(
            r#"<Override PartName="{part_name}" ContentType="{content_type}"/>"#
        );
        xml.insert_str(idx, &insert);
    }
    parts.insert("[Content_Types].xml".to_string(), xml.into_bytes());
    Ok(())
}

fn ensure_workbook_rels_has_relationship(
    parts: &mut BTreeMap<String, Vec<u8>>,
    rel_type: &str,
    target: &str,
) -> Result<(), WriteError> {
    let rels_name = "xl/_rels/workbook.xml.rels";
    let Some(existing) = parts.get(rels_name).cloned() else {
        return Ok(());
    };
    let mut xml = String::from_utf8(existing)
        .map_err(|e| WriteError::Io(std::io::Error::new(std::io::ErrorKind::InvalidData, e)))?;
    if xml.contains(rel_type) {
        parts.insert(rels_name.to_string(), xml.into_bytes());
        return Ok(());
    }
    let next = next_relationship_id_in_xml(&xml);
    let rel = format!(
        r#"<Relationship Id="rId{next}" Type="{rel_type}" Target="{target}"/>"#
    );
    if let Some(idx) = xml.rfind("</Relationships>") {
        xml.insert_str(idx, &rel);
    }
    parts.insert(rels_name.to_string(), xml.into_bytes());
    Ok(())
}

fn next_relationship_id<'a>(ids: impl Iterator<Item = &'a str>) -> u32 {
    let mut max_id = 0u32;
    for id in ids {
        if let Some(rest) = id.strip_prefix("rId") {
            if let Ok(n) = rest.parse::<u32>() {
                max_id = max_id.max(n);
            }
        }
    }
    max_id + 1
}

fn next_relationship_id_in_xml(xml: &str) -> u32 {
    let mut max_id = 0u32;
    let mut rest = xml;
    while let Some(idx) = rest.find("Id=\"rId") {
        let after = &rest[idx + "Id=\"rId".len()..];
        let mut digits = String::new();
        for ch in after.chars() {
            if ch.is_ascii_digit() {
                digits.push(ch);
            } else {
                break;
            }
        }
        if let Ok(n) = digits.parse::<u32>() {
            max_id = max_id.max(n);
        }
        rest = &after[digits.len()..];
    }
    max_id + 1
}
