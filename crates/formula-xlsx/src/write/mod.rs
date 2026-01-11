use std::collections::{BTreeMap, HashMap, HashSet};
use std::io::{Cursor, Write};

use formula_model::rich_text::{RichText, Underline};
use formula_model::{
    CellRef, CellValue, ErrorValue, Outline, OutlineEntry, Range, SheetVisibility, Worksheet,
    WorksheetId,
};
use quick_xml::events::Event;
use quick_xml::events::attributes::AttrError;
use quick_xml::Reader;
use quick_xml::Writer;
use thiserror::Error;
use zip::write::FileOptions;
use zip::ZipWriter;

use crate::path::resolve_target;
use crate::styles::StylesPart;
use crate::{CellValueKind, DateSystem, SheetMeta, XlsxDocument};

const WORKBOOK_PART: &str = "xl/workbook.xml";
const WORKBOOK_RELS_PART: &str = "xl/_rels/workbook.xml.rels";
const REL_TYPE_STYLES: &str =
    "http://schemas.openxmlformats.org/officeDocument/2006/relationships/styles";
const REL_TYPE_SHARED_STRINGS: &str =
    "http://schemas.openxmlformats.org/officeDocument/2006/relationships/sharedStrings";

mod dimension;

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
    #[error(transparent)]
    Styles(#[from] crate::styles::StylesPartError),
}

const WORKSHEET_REL_TYPE: &str =
    "http://schemas.openxmlformats.org/officeDocument/2006/relationships/worksheet";
const WORKSHEET_CONTENT_TYPE: &str =
    "application/vnd.openxmlformats-officedocument.spreadsheetml.worksheet+xml";

#[derive(Debug)]
struct SheetStructurePlan {
    sheets: Vec<SheetMeta>,
}

fn sheet_state_from_visibility(visibility: SheetVisibility) -> Option<String> {
    match visibility {
        SheetVisibility::Visible => None,
        SheetVisibility::Hidden => Some("hidden".to_string()),
        SheetVisibility::VeryHidden => Some("veryHidden".to_string()),
    }
}

fn sheet_part_number(path: &str) -> Option<u32> {
    let file = path.rsplit('/').next()?;
    if !file.starts_with("sheet") || !file.ends_with(".xml") {
        return None;
    }
    let digits = file.strip_prefix("sheet")?.strip_suffix(".xml")?;
    digits.parse::<u32>().ok()
}

fn next_sheet_part_number<'a>(paths: impl Iterator<Item = &'a str>) -> u32 {
    paths.filter_map(sheet_part_number).max().unwrap_or(0) + 1
}

fn plan_sheet_structure(
    doc: &XlsxDocument,
    parts: &mut BTreeMap<String, Vec<u8>>,
    is_new: bool,
) -> Result<SheetStructurePlan, WriteError> {
    let workbook_sheet_ids: HashSet<WorksheetId> = doc.workbook.sheets.iter().map(|s| s.id).collect();

    let existing_by_id: HashMap<WorksheetId, SheetMeta> = doc
        .meta
        .sheets
        .iter()
        .cloned()
        .map(|meta| (meta.worksheet_id, meta))
        .collect();

    let removed: Vec<SheetMeta> = doc
        .meta
        .sheets
        .iter()
        .filter(|meta| !workbook_sheet_ids.contains(&meta.worksheet_id))
        .cloned()
        .collect();

    let mut next_sheet_id = doc
        .meta
        .sheets
        .iter()
        .filter(|meta| workbook_sheet_ids.contains(&meta.worksheet_id))
        .map(|meta| meta.sheet_id)
        .max()
        .unwrap_or(0)
        + 1;

    let mut next_rel_id_num = if is_new {
        doc.meta
            .sheets
            .iter()
            .filter(|meta| workbook_sheet_ids.contains(&meta.worksheet_id))
            .filter_map(|meta| meta.relationship_id.strip_prefix("rId")?.parse::<u32>().ok())
            .max()
            .unwrap_or(0)
            + 1
    } else {
        parts
            .get("xl/_rels/workbook.xml.rels")
            .and_then(|b| std::str::from_utf8(b).ok())
            .map(next_relationship_id_in_xml)
            .unwrap_or(1)
    };

    let existing_paths = doc.meta.sheets.iter().map(|meta| meta.path.as_str());
    let part_paths = parts.keys().map(|p| p.as_str());
    let mut next_sheet_part = next_sheet_part_number(existing_paths.chain(part_paths));
    let mut used_paths: HashSet<String> = doc.meta.sheets.iter().map(|m| m.path.clone()).collect();

    let mut sheets: Vec<SheetMeta> = Vec::with_capacity(doc.workbook.sheets.len());
    let mut added: Vec<SheetMeta> = Vec::new();

    for sheet in &doc.workbook.sheets {
        if let Some(existing) = existing_by_id.get(&sheet.id) {
            let mut meta = existing.clone();
            meta.state = sheet_state_from_visibility(sheet.visibility);
            sheets.push(meta);
            continue;
        }

        let relationship_id = format!("rId{next_rel_id_num}");
        next_rel_id_num += 1;

        let mut path;
        loop {
            path = format!("xl/worksheets/sheet{next_sheet_part}.xml");
            next_sheet_part += 1;
            if !used_paths.contains(&path) && !parts.contains_key(&path) {
                break;
            }
        }
        used_paths.insert(path.clone());

        let meta = SheetMeta {
            worksheet_id: sheet.id,
            sheet_id: next_sheet_id,
            relationship_id,
            state: sheet_state_from_visibility(sheet.visibility),
            path,
        };
        next_sheet_id += 1;

        added.push(meta.clone());
        sheets.push(meta);
    }

    if !is_new && (!added.is_empty() || !removed.is_empty()) {
        for meta in &removed {
            parts.remove(&meta.path);
            parts.remove(&crate::openxml::rels_part_name(&meta.path));
        }

        patch_workbook_rels_for_sheet_edits(parts, &removed, &added)?;
        patch_content_types_for_sheet_edits(parts, &removed, &added)?;
    }

    Ok(SheetStructurePlan {
        sheets,
    })
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

    let sheet_plan = plan_sheet_structure(doc, &mut parts, is_new)?;
    if is_new {
        parts = generate_minimal_package(&sheet_plan.sheets)?;
    }

    let (mut styles_part_name, mut shared_strings_part_name) = (
        "xl/styles.xml".to_string(),
        "xl/sharedStrings.xml".to_string(),
    );
    if let Some(rels) = parts.get(WORKBOOK_RELS_PART).map(|b| b.as_slice()) {
        if let Some(target) = relationship_target_by_type(rels, REL_TYPE_STYLES)? {
            styles_part_name = resolve_target(WORKBOOK_PART, &target);
        }
        if let Some(target) = relationship_target_by_type(rels, REL_TYPE_SHARED_STRINGS)? {
            shared_strings_part_name = resolve_target(WORKBOOK_PART, &target);
        }
    }

    let (shared_strings_xml, shared_string_lookup) =
        build_shared_strings_xml(doc, &sheet_plan.sheets)?;
    if is_new || !shared_string_lookup.is_empty() || parts.contains_key(&shared_strings_part_name) {
        parts.insert(shared_strings_part_name.clone(), shared_strings_xml);
    }

    // Parse/update styles.xml (cellXfs) so cell `s` attributes refer to real xf indices.
    let mut style_table = doc.workbook.styles.clone();
    let mut styles_part = StylesPart::parse_or_default(
        parts.get(&styles_part_name).map(|b| b.as_slice()),
        &mut style_table,
    )?;
    let style_ids = doc
        .workbook
        .sheets
        .iter()
        .flat_map(|sheet| sheet.iter_cells().map(|(_, cell)| cell.style_id))
        .filter(|style_id| *style_id != 0);
    let style_to_xf = styles_part.xf_indices_for_style_ids(style_ids, &style_table)?;
    parts.insert(styles_part_name.clone(), styles_part.to_xml_bytes());

    // Ensure core relationship/content types metadata exists when we synthesize new
    // parts for existing packages. For existing relationships we preserve IDs by
    // only adding missing entries with a new `rIdN`.
    if parts.contains_key(&shared_strings_part_name) {
        ensure_content_types_override(
            &mut parts,
            &format!("/{shared_strings_part_name}"),
            "application/vnd.openxmlformats-officedocument.spreadsheetml.sharedStrings+xml",
        )?;
        ensure_workbook_rels_has_relationship(
            &mut parts,
            REL_TYPE_SHARED_STRINGS,
            &relationship_target_from_workbook(&shared_strings_part_name),
        )?;
    }
    if parts.contains_key(&styles_part_name) {
        ensure_content_types_override(
            &mut parts,
            &format!("/{styles_part_name}"),
            "application/vnd.openxmlformats-officedocument.spreadsheetml.styles+xml",
        )?;
        ensure_workbook_rels_has_relationship(
            &mut parts,
            REL_TYPE_STYLES,
            &relationship_target_from_workbook(&styles_part_name),
        )?;
    }

    let workbook_orig = parts.get("xl/workbook.xml").map(|b| b.as_slice());
    parts.insert(
        "xl/workbook.xml".to_string(),
        write_workbook_xml(doc, workbook_orig, &sheet_plan.sheets)?,
    );

    for sheet_meta in &sheet_plan.sheets {
        let sheet = doc
            .workbook
            .sheet(sheet_meta.worksheet_id)
            .ok_or_else(|| WriteError::Io(std::io::Error::new(std::io::ErrorKind::NotFound, "worksheet not found")))?;
        let orig = parts.get(&sheet_meta.path).map(|b| b.as_slice());
        parts.insert(
            sheet_meta.path.clone(),
            write_worksheet_xml(
                doc,
                sheet_meta,
                sheet,
                orig,
                &shared_string_lookup,
                &style_to_xf,
            )?,
        );
    }

    Ok(parts)
}

#[derive(Clone, Debug, PartialEq, Eq, Hash)]
struct SharedStringKey {
    text: String,
    runs: Vec<SharedStringRunKey>,
}

#[derive(Clone, Debug, PartialEq, Eq, Hash)]
struct SharedStringRunKey {
    start: usize,
    end: usize,
    style: SharedStringRunStyleKey,
}

#[derive(Clone, Debug, PartialEq, Eq, Hash)]
struct SharedStringRunStyleKey {
    bold: Option<bool>,
    italic: Option<bool>,
    underline: Option<u8>,
    color: Option<u32>,
    font: Option<String>,
    size_100pt: Option<u16>,
}

impl SharedStringKey {
    fn plain(text: &str) -> Self {
        Self {
            text: text.to_string(),
            runs: Vec::new(),
        }
    }

    fn from_rich_text(rich: &RichText) -> Self {
        let runs = rich
            .runs
            .iter()
            .map(|run| SharedStringRunKey {
                start: run.start,
                end: run.end,
                style: SharedStringRunStyleKey {
                    bold: run.style.bold,
                    italic: run.style.italic,
                    underline: run.style.underline.map(underline_key),
                    color: run.style.color.and_then(|c| c.argb()),
                    font: run.style.font.clone(),
                    size_100pt: run.style.size_100pt,
                },
            })
            .collect();
        Self {
            text: rich.text.clone(),
            runs,
        }
    }
}

fn underline_key(underline: Underline) -> u8 {
    match underline {
        Underline::None => 0,
        Underline::Single => 1,
        Underline::Double => 2,
        Underline::SingleAccounting => 3,
        Underline::DoubleAccounting => 4,
    }
}

fn build_shared_strings_xml(
    doc: &XlsxDocument,
    sheets: &[SheetMeta],
) -> Result<(Vec<u8>, HashMap<SharedStringKey, u32>), WriteError> {
    let mut table: Vec<RichText> = doc.shared_strings.clone();
    let mut lookup: HashMap<SharedStringKey, u32> = HashMap::new();
    for (idx, rich) in table.iter().enumerate() {
        lookup
            .entry(SharedStringKey::from_rich_text(rich))
            .or_insert(idx as u32);
    }

    let mut ref_count: u32 = 0;

    for sheet_meta in sheets {
        let sheet = match doc.workbook.sheet(sheet_meta.worksheet_id) {
            Some(s) => s,
            None => continue,
        };

        let mut cells: Vec<(CellRef, &formula_model::Cell)> = sheet.iter_cells().collect();
        cells.sort_by_key(|(r, _)| (r.row, r.col));
        for (cell_ref, cell) in cells {
            let meta = doc.meta.cell_meta.get(&(sheet_meta.worksheet_id, cell_ref));
            let kind = effective_value_kind(meta, cell);
            let CellValueKind::SharedString { .. } = kind else {
                continue;
            };

            match &cell.value {
                CellValue::String(text) => {
                    ref_count += 1;
                    if meta
                        .and_then(|m| m.value_kind.clone())
                        .and_then(|k| match k {
                            CellValueKind::SharedString { index } => Some(index),
                            _ => None,
                        })
                        .and_then(|idx| doc.shared_strings.get(idx as usize))
                        .map(|rt| rt.text.as_str() == text.as_str())
                        .unwrap_or(false)
                    {
                        // Preserve the original shared string index even if the entry
                        // contains rich formatting.
                        continue;
                    }

                    let key = SharedStringKey::plain(text);
                    if !lookup.contains_key(&key) {
                        let new_index = table.len() as u32;
                        table.push(RichText::new(text.clone()));
                        lookup.insert(key, new_index);
                    }
                }
                CellValue::RichText(rich) => {
                    ref_count += 1;
                    if meta
                        .and_then(|m| m.value_kind.clone())
                        .and_then(|k| match k {
                            CellValueKind::SharedString { index } => Some(index),
                            _ => None,
                        })
                        .and_then(|idx| doc.shared_strings.get(idx as usize))
                        .map(|rt| rt == rich)
                        .unwrap_or(false)
                    {
                        continue;
                    }

                    let key = SharedStringKey::from_rich_text(rich);
                    if !lookup.contains_key(&key) {
                        let new_index = table.len() as u32;
                        table.push(rich.clone());
                        lookup.insert(key, new_index);
                    }
                }
                _ => {
                    // Non-string values ignore shared string bookkeeping.
                }
            }
        }
    }

    let mut xml = String::new();
    xml.push_str(r#"<?xml version="1.0" encoding="UTF-8" standalone="yes"?>"#);
    xml.push_str(r#"<sst xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main""#);
    xml.push_str(&format!(r#" count="{ref_count}" uniqueCount="{}">"#, table.len()));
    for rich in &table {
        xml.push_str("<si>");
        if rich.runs.is_empty() {
            write_shared_string_t(&mut xml, &rich.text);
        } else {
            for run in &rich.runs {
                xml.push_str("<r>");
                if !run.style.is_empty() {
                    xml.push_str("<rPr>");
                    write_shared_string_rpr(&mut xml, &run.style);
                    xml.push_str("</rPr>");
                }
                let segment = rich.slice_run_text(run);
                write_shared_string_t(&mut xml, segment);
                xml.push_str("</r>");
            }
        }
        xml.push_str("</si>");
    }
    xml.push_str("</sst>");

    Ok((xml.into_bytes(), lookup))
}

fn write_shared_string_t(xml: &mut String, text: &str) {
    xml.push_str("<t");
    if needs_space_preserve(text) {
        xml.push_str(r#" xml:space="preserve""#);
    }
    xml.push('>');
    xml.push_str(&escape_text(text));
    xml.push_str("</t>");
}

fn write_shared_string_rpr(xml: &mut String, style: &formula_model::rich_text::RichTextRunStyle) {
    if let Some(font) = &style.font {
        xml.push_str(r#"<rFont val=""#);
        xml.push_str(&escape_attr(font));
        xml.push_str(r#""/>"#);
    }

    if let Some(size_100pt) = style.size_100pt {
        xml.push_str(r#"<sz val=""#);
        xml.push_str(&format_size_100pt(size_100pt));
        xml.push_str(r#""/>"#);
    }

    if let Some(color) = style.color.and_then(|c| c.argb()) {
        xml.push_str(r#"<color rgb=""#);
        xml.push_str(&format!("{:08X}", color));
        xml.push_str(r#""/>"#);
    }

    if let Some(bold) = style.bold {
        if bold {
            xml.push_str("<b/>");
        } else {
            xml.push_str(r#"<b val="0"/>"#);
        }
    }

    if let Some(italic) = style.italic {
        if italic {
            xml.push_str("<i/>");
        } else {
            xml.push_str(r#"<i val="0"/>"#);
        }
    }

    if let Some(underline) = style.underline {
        match underline {
            Underline::Single => xml.push_str("<u/>"),
            other => {
                xml.push_str(r#"<u val=""#);
                xml.push_str(other.to_ooxml().unwrap_or("single"));
                xml.push_str(r#""/>"#);
            }
        }
    }
}

fn format_size_100pt(size_100pt: u16) -> String {
    let int = size_100pt / 100;
    let frac = size_100pt % 100;
    if frac == 0 {
        return int.to_string();
    }

    let mut s = format!("{int}.{frac:02}");
    while s.ends_with('0') {
        s.pop();
    }
    s
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

fn write_workbook_xml(
    doc: &XlsxDocument,
    original: Option<&[u8]>,
    sheets: &[SheetMeta],
) -> Result<Vec<u8>, WriteError> {
    if let Some(original) = original {
        return patch_workbook_xml(doc, original, sheets);
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
    for sheet_meta in sheets {
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

fn patch_workbook_xml(
    doc: &XlsxDocument,
    original: &[u8],
    sheets: &[SheetMeta],
) -> Result<Vec<u8>, WriteError> {
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

                for sheet_meta in sheets {
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
                for sheet_meta in sheets {
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
    shared_lookup: &HashMap<SharedStringKey, u32>,
    style_to_xf: &HashMap<u32, u32>,
) -> Result<Vec<u8>, WriteError> {
    if let Some(original) = original {
        return patch_worksheet_xml(doc, sheet_meta, sheet, original, shared_lookup, style_to_xf);
    }

    let dimension = dimension::worksheet_dimension_range(sheet).to_string();
    let sheet_data_xml =
        render_sheet_data(doc, sheet_meta, sheet, shared_lookup, style_to_xf, None);

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
    shared_lookup: &HashMap<SharedStringKey, u32>,
    style_to_xf: &HashMap<u32, u32>,
) -> Result<Vec<u8>, WriteError> {
    let (original_has_dimension, original_used_range) = scan_worksheet_xml(original)?;
    let new_used_range = dimension::worksheet_used_range(sheet);
    let insert_dimension = !original_has_dimension && original_used_range != new_used_range;
    let dimension_range = dimension::worksheet_dimension_range(sheet);
    let dimension_ref = dimension_range.to_string();

    let outline = std::str::from_utf8(original)
        .ok()
        .and_then(|xml| crate::outline::read_outline_from_worksheet_xml(xml).ok());
    let sheet_data_xml = render_sheet_data(
        doc,
        sheet_meta,
        sheet,
        shared_lookup,
        style_to_xf,
        outline.as_ref(),
    );

    let mut reader = Reader::from_reader(original);
    reader.config_mut().trim_text(false);
    let mut buf = Vec::new();
    let mut writer = Writer::new(Vec::with_capacity(original.len() + sheet_data_xml.len()));

    let mut skipping_sheet_data = false;
    let mut inserted_dimension = false;
    let mut saw_sheet_pr = false;
    let mut in_sheet_pr = false;
    loop {
        match reader.read_event_into(&mut buf)? {
            Event::Start(e) if e.name().as_ref() == b"sheetPr" => {
                saw_sheet_pr = true;
                in_sheet_pr = true;
                writer.write_event(Event::Start(e.into_owned()))?;
            }
            Event::Empty(e) if e.name().as_ref() == b"sheetPr" => {
                saw_sheet_pr = true;
                writer.write_event(Event::Empty(e.into_owned()))?;
                if insert_dimension && !inserted_dimension {
                    insert_dimension_element(&mut writer, &dimension_ref);
                    inserted_dimension = true;
                }
            }
            Event::End(e) if e.name().as_ref() == b"sheetPr" => {
                in_sheet_pr = false;
                writer.write_event(Event::End(e.into_owned()))?;
                if insert_dimension && !inserted_dimension {
                    insert_dimension_element(&mut writer, &dimension_ref);
                    inserted_dimension = true;
                }
            }

            Event::Start(e) if e.name().as_ref() == b"dimension" => {
                if dimension_matches(&e, dimension_range)? {
                    writer.write_event(Event::Start(e.into_owned()))?;
                } else {
                    write_dimension_element(&mut writer, &e, &dimension_ref, false)?;
                }
            }
            Event::Empty(e) if e.name().as_ref() == b"dimension" => {
                if dimension_matches(&e, dimension_range)? {
                    writer.write_event(Event::Empty(e.into_owned()))?;
                } else {
                    write_dimension_element(&mut writer, &e, &dimension_ref, true)?;
                }
            }

            Event::Start(e) if e.name().as_ref() == b"sheetData" => {
                if insert_dimension
                    && !inserted_dimension
                    && !saw_sheet_pr
                    && !in_sheet_pr
                {
                    insert_dimension_element(&mut writer, &dimension_ref);
                    inserted_dimension = true;
                }
                skipping_sheet_data = true;
                writer.get_mut().extend_from_slice(sheet_data_xml.as_bytes());
            }
            Event::Empty(e) if e.name().as_ref() == b"sheetData" => {
                if insert_dimension
                    && !inserted_dimension
                    && !saw_sheet_pr
                    && !in_sheet_pr
                {
                    insert_dimension_element(&mut writer, &dimension_ref);
                    inserted_dimension = true;
                }
                writer.get_mut().extend_from_slice(sheet_data_xml.as_bytes());
                drop(e);
            }
            Event::End(e) if e.name().as_ref() == b"sheetData" => {
                skipping_sheet_data = false;
                drop(e);
            }
            Event::Eof => break,
            ev if skipping_sheet_data => drop(ev),
            ev => {
                match &ev {
                    Event::Start(e) | Event::Empty(e)
                        if insert_dimension
                            && !inserted_dimension
                            && !saw_sheet_pr
                            && !in_sheet_pr
                            && e.name().as_ref() != b"worksheet" =>
                    {
                        insert_dimension_element(&mut writer, &dimension_ref);
                        inserted_dimension = true;
                    }
                    Event::End(e)
                        if insert_dimension
                            && !inserted_dimension
                            && e.name().as_ref() == b"worksheet" =>
                    {
                        insert_dimension_element(&mut writer, &dimension_ref);
                        inserted_dimension = true;
                    }
                    _ => {}
                }
                writer.write_event(ev.into_owned())?
            }
        }
        buf.clear();
    }

    Ok(writer.into_inner())
}

fn scan_worksheet_xml(original: &[u8]) -> Result<(bool, Option<Range>), WriteError> {
    let mut reader = Reader::from_reader(original);
    reader.config_mut().trim_text(false);
    let mut buf = Vec::new();

    let mut has_dimension = false;
    let mut in_sheet_data = false;
    let mut min_cell: Option<CellRef> = None;
    let mut max_cell: Option<CellRef> = None;

    loop {
        match reader.read_event_into(&mut buf)? {
            Event::Start(e) | Event::Empty(e) if e.name().as_ref() == b"dimension" => {
                has_dimension = true;
            }
            Event::Start(e) if e.name().as_ref() == b"sheetData" => in_sheet_data = true,
            Event::End(e) if e.name().as_ref() == b"sheetData" => in_sheet_data = false,
            Event::Empty(e) if e.name().as_ref() == b"sheetData" => {
                in_sheet_data = false;
                drop(e);
            }
            Event::Start(e) | Event::Empty(e) if in_sheet_data && e.name().as_ref() == b"c" => {
                for attr in e.attributes() {
                    let attr = attr?;
                    if attr.key.as_ref() != b"r" {
                        continue;
                    }
                    let a1 = attr.unescape_value()?.into_owned();
                    let Ok(cell_ref) = CellRef::from_a1(&a1) else {
                        continue;
                    };
                    min_cell = Some(match min_cell {
                        Some(min) => CellRef::new(min.row.min(cell_ref.row), min.col.min(cell_ref.col)),
                        None => cell_ref,
                    });
                    max_cell = Some(match max_cell {
                        Some(max) => CellRef::new(max.row.max(cell_ref.row), max.col.max(cell_ref.col)),
                        None => cell_ref,
                    });
                    break;
                }
            }
            Event::Eof => break,
            _ => {}
        }
        buf.clear();
    }

    let used_range = match (min_cell, max_cell) {
        (Some(start), Some(end)) => Some(Range::new(start, end)),
        _ => None,
    };
    Ok((has_dimension, used_range))
}

fn insert_dimension_element(writer: &mut Writer<Vec<u8>>, dimension_ref: &str) {
    writer.get_mut().extend_from_slice(b"<dimension ref=\"");
    writer
        .get_mut()
        .extend_from_slice(escape_attr(dimension_ref).as_bytes());
    writer.get_mut().extend_from_slice(b"\"/>");
}

fn dimension_matches(
    e: &quick_xml::events::BytesStart<'_>,
    expected: Range,
) -> Result<bool, WriteError> {
    let mut ref_value = None;
    for attr in e.attributes() {
        let attr = attr?;
        if attr.key.as_ref() == b"ref" {
            ref_value = Some(attr.unescape_value()?.into_owned());
            break;
        }
    }
    let Some(ref_value) = ref_value else {
        return Ok(false);
    };
    Ok(dimension::parse_dimension_ref(&ref_value) == Some(expected))
}

fn write_dimension_element(
    writer: &mut Writer<Vec<u8>>,
    e: &quick_xml::events::BytesStart<'_>,
    dimension_ref: &str,
    is_empty: bool,
) -> Result<(), WriteError> {
    writer.get_mut().extend_from_slice(b"<dimension");
    let mut wrote_ref = false;
    for attr in e.attributes() {
        let attr = attr?;
        writer.get_mut().push(b' ');
        writer.get_mut().extend_from_slice(attr.key.as_ref());
        writer.get_mut().extend_from_slice(b"=\"");
        if attr.key.as_ref() == b"ref" {
            wrote_ref = true;
            writer
                .get_mut()
                .extend_from_slice(escape_attr(dimension_ref).as_bytes());
        } else {
            writer.get_mut().extend_from_slice(
                escape_attr(&attr.unescape_value()?.into_owned()).as_bytes(),
            );
        }
        writer.get_mut().push(b'"');
    }

    if !wrote_ref {
        writer.get_mut().extend_from_slice(b" ref=\"");
        writer
            .get_mut()
            .extend_from_slice(escape_attr(dimension_ref).as_bytes());
        writer.get_mut().push(b'"');
    }

    if is_empty {
        writer.get_mut().extend_from_slice(b"/>");
    } else {
        writer.get_mut().push(b'>');
    }
    Ok(())
}

fn render_sheet_data(
    doc: &XlsxDocument,
    sheet_meta: &SheetMeta,
    sheet: &Worksheet,
    shared_lookup: &HashMap<SharedStringKey, u32>,
    style_to_xf: &HashMap<u32, u32>,
    outline: Option<&Outline>,
) -> String {
    let mut out = String::new();
    out.push_str("<sheetData>");

    let mut cells: Vec<(CellRef, &formula_model::Cell)> = sheet.iter_cells().collect();
    cells.sort_by_key(|(r, _)| (r.row, r.col));

    let mut outline_rows: Vec<u32> = Vec::new();
    if let Some(outline) = outline {
        // Preserve outline-only rows (groups, hidden rows, etc) even if they contain no cells.
        // We don't attempt to preserve all row-level metadata yet—only the outline-related attrs.
        for (row, entry) in outline.rows.iter() {
            if entry.level > 0 || entry.hidden.is_hidden() || entry.collapsed {
                outline_rows.push(row);
            }
        }
    }

    let mut cell_idx = 0usize;
    let mut outline_idx = 0usize;

    while cell_idx < cells.len() || outline_idx < outline_rows.len() {
        let next_cell_row = cells
            .get(cell_idx)
            .map(|(cell_ref, _)| cell_ref.row + 1)
            .unwrap_or(u32::MAX);
        let next_outline_row = outline_rows.get(outline_idx).copied().unwrap_or(u32::MAX);
        let row_1_based = next_cell_row.min(next_outline_row);

        if row_1_based == next_outline_row {
            outline_idx += 1;
        }

        let outline_entry: OutlineEntry = outline
            .map(|outline| outline.rows.entry(row_1_based))
            .unwrap_or_default();

        out.push_str(&format!(r#"<row r="{row_1_based}""#));
        if outline_entry.level > 0 {
            out.push_str(&format!(r#" outlineLevel="{}""#, outline_entry.level));
        }
        if outline_entry.hidden.is_hidden() {
            out.push_str(r#" hidden="1""#);
        }
        if outline_entry.collapsed {
            out.push_str(r#" collapsed="1""#);
        }

        let mut wrote_any_cell = false;

        while let Some((cell_ref, cell)) = cells.get(cell_idx).copied() {
            if cell_ref.row + 1 != row_1_based {
                break;
            }
            if !wrote_any_cell {
                out.push('>');
                wrote_any_cell = true;
            }
            cell_idx += 1;

        out.push_str(r#"<c r=""#);
        out.push_str(&cell_ref.to_a1());
        out.push('"');

        if cell.style_id != 0 {
            if let Some(xf_index) = style_to_xf.get(&cell.style_id) {
                out.push_str(&format!(r#" s="{xf_index}""#));
            }
        }

        let meta = doc.meta.cell_meta.get(&(sheet_meta.worksheet_id, cell_ref));
        let value_kind = effective_value_kind(meta, cell);

        if !matches!(cell.value, CellValue::Empty) {
            match &value_kind {
                CellValueKind::SharedString { .. } => out.push_str(r#" t="s""#),
                CellValueKind::InlineString => out.push_str(r#" t="inlineStr""#),
                CellValueKind::Bool => out.push_str(r#" t="b""#),
                CellValueKind::Error => out.push_str(r#" t="e""#),
                CellValueKind::Str => out.push_str(r#" t="str""#),
                CellValueKind::Number => {}
                CellValueKind::Other { t } => {
                    out.push_str(&format!(r#" t="{}""#, escape_attr(t)));
                }
            }
        }

        out.push('>');

        let model_formula = cell.formula.as_deref();
        let formula_meta = match (model_formula, meta.and_then(|m| m.formula.clone())) {
            (Some(_), Some(meta)) => Some(meta),
            (Some(formula), None) => Some(crate::FormulaMeta {
                file_text: crate::formula_text::add_xlfn_prefixes(strip_leading_equals(formula)),
                ..Default::default()
            }),
            (None, Some(meta)) => {
                // The in-memory model doesn't currently represent shared formulas for follower
                // cells. Preserve those formulas when the stored SpreadsheetML indicates a formula
                // even if the model omits it.
                if meta.file_text.is_empty()
                    && meta.t.is_none()
                    && meta.reference.is_none()
                    && meta.shared_index.is_none()
                    && meta.always_calc.is_none()
                {
                    None
                } else if meta.file_text.is_empty() {
                    Some(meta)
                } else {
                    // Model cleared the formula; don't keep stale formula text from metadata.
                    None
                }
            }
            (None, None) => None,
        };

        if let Some(formula_meta) = formula_meta {
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

            let file_text = formula_file_text(&formula_meta, model_formula);
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
            value @ CellValue::String(s) if matches!(value_kind, CellValueKind::Other { .. }) => {
                out.push_str("<v>");
                out.push_str(&escape_text(&raw_or_other(meta, s)));
                out.push_str("</v>");
            }
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
            value @ CellValue::String(s) => match &value_kind {
                CellValueKind::SharedString { .. } => {
                    let idx = shared_string_index(doc, meta, value, shared_lookup);
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
                    let idx = shared_string_index(doc, meta, value, shared_lookup);
                    out.push_str("<v>");
                    out.push_str(&idx.to_string());
                    out.push_str("</v>");
                }
            },
            value @ CellValue::RichText(rich) => {
                // Rich text is stored in the shared strings table.
                let idx = shared_string_index(doc, meta, value, shared_lookup);
                if idx != 0 || !rich.text.is_empty() {
                    out.push_str("<v>");
                    out.push_str(&idx.to_string());
                    out.push_str("</v>");
                }
            }
            _ => {
                // Array/Spill not yet modeled for writing. Preserve as blank.
            }
        }

        out.push_str("</c>");
    }

        if wrote_any_cell {
            out.push_str("</row>");
        } else {
            out.push_str("/>");
        }
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
        CellValue::RichText(_) => CellValueKind::SharedString { index: 0 },
        CellValue::Empty => CellValueKind::Number,
        _ => CellValueKind::Number,
    }
}

fn effective_value_kind(meta: Option<&crate::CellMeta>, cell: &formula_model::Cell) -> CellValueKind {
    if let Some(meta) = meta {
        if let Some(kind) = meta.value_kind.clone() {
            // Cells with less-common or unknown `t=` attributes require the original `<v>` payload
            // to round-trip safely. If we don't have it, fall back to the inferred kind so we emit
            // a valid SpreadsheetML representation.
            if matches!(kind, CellValueKind::Other { .. }) {
                if meta.raw_value.is_some() && matches!(cell.value, CellValue::String(_)) {
                    return kind;
                }
            } else if value_kind_compatible(&kind, &cell.value) {
                return kind;
            }
        }
    }

    infer_value_kind(cell)
}

fn value_kind_compatible(kind: &CellValueKind, value: &CellValue) -> bool {
    match (kind, value) {
        (_, CellValue::Empty) => true,
        (CellValueKind::Number, CellValue::Number(_)) => true,
        (CellValueKind::Bool, CellValue::Boolean(_)) => true,
        (CellValueKind::Error, CellValue::Error(_)) => true,
        (CellValueKind::SharedString { .. }, CellValue::String(_) | CellValue::RichText(_)) => true,
        (CellValueKind::InlineString, CellValue::String(_)) => true,
        (CellValueKind::Str, CellValue::String(_)) => true,
        _ => false,
    }
}

fn formula_file_text(meta: &crate::FormulaMeta, display: Option<&str>) -> String {
    let Some(display) = display else {
        return strip_leading_equals(&meta.file_text).to_string();
    };

    let display = strip_leading_equals(display);

    // Preserve stored file text if the model's display text matches.
    if !meta.file_text.is_empty() && crate::formula_text::strip_xlfn_prefixes(&meta.file_text) == display {
        return strip_leading_equals(&meta.file_text).to_string();
    }

    crate::formula_text::add_xlfn_prefixes(display)
}

fn strip_leading_equals(s: &str) -> &str {
    s.strip_prefix('=').unwrap_or(s)
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

fn raw_or_other(meta: Option<&crate::CellMeta>, s: &str) -> String {
    // Unknown/less-common `t=` types store their payload as text; preserve the original `<v>`
    // content when it still matches the in-memory value.
    raw_or_str(meta, s)
}

fn shared_string_index(
    doc: &XlsxDocument,
    meta: Option<&crate::CellMeta>,
    value: &CellValue,
    shared_lookup: &HashMap<SharedStringKey, u32>,
) -> u32 {
    match value {
        CellValue::String(text) => {
            if let Some(meta) = meta {
                if let Some(CellValueKind::SharedString { index }) = &meta.value_kind {
                    if doc
                        .shared_strings
                        .get(*index as usize)
                        .map(|rt| rt.text.as_str())
                        == Some(text.as_str())
                    {
                        return *index;
                    }
                }
            }
            shared_lookup
                .get(&SharedStringKey::plain(text))
                .copied()
                .unwrap_or(0)
        }
        CellValue::RichText(rich) => {
            if let Some(meta) = meta {
                if let Some(CellValueKind::SharedString { index }) = &meta.value_kind {
                    if doc
                        .shared_strings
                        .get(*index as usize)
                        .map(|rt| rt == rich)
                        .unwrap_or(false)
                    {
                        return *index;
                    }
                }
            }
            shared_lookup
                .get(&SharedStringKey::from_rich_text(rich))
                .copied()
                .unwrap_or(0)
        }
        _ => 0,
    }
}

fn generate_minimal_package(sheets: &[SheetMeta]) -> Result<BTreeMap<String, Vec<u8>>, WriteError> {
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
        minimal_workbook_rels_xml(sheets).into_bytes(),
    );

    parts.insert(
        "[Content_Types].xml".to_string(),
        minimal_content_types_xml(sheets).into_bytes(),
    );

    Ok(parts)
}

fn minimal_workbook_rels_xml(sheets: &[SheetMeta]) -> String {
    let mut xml = String::new();
    xml.push_str(r#"<?xml version="1.0" encoding="UTF-8" standalone="yes"?>"#);
    xml.push_str(r#"<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">"#);

    for sheet_meta in sheets {
        let target = relationship_target_from_workbook(&sheet_meta.path);
        xml.push_str(r#"<Relationship Id=""#);
        xml.push_str(&escape_attr(&sheet_meta.relationship_id));
        xml.push_str(r#"" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/worksheet" Target=""#);
        xml.push_str(&escape_attr(&target));
        xml.push_str(r#""/>"#);
    }

    let next = next_relationship_id(
        sheets.iter().map(|s| s.relationship_id.as_str()),
    );
    xml.push_str(&format!(
        r#"<Relationship Id="rId{next}" Type="{REL_TYPE_STYLES}" Target="styles.xml"/>"#
    ));
    let next2 = next + 1;
    xml.push_str(&format!(
        r#"<Relationship Id="rId{next2}" Type="{REL_TYPE_SHARED_STRINGS}" Target="sharedStrings.xml"/>"#
    ));
    xml.push_str("</Relationships>");
    xml
}

fn relationship_target_from_workbook(part_name: &str) -> String {
    let base_dir = WORKBOOK_PART
        .rsplit_once('/')
        .map(|(dir, _)| dir)
        .unwrap_or("");
    relative_target(base_dir, part_name)
}

fn relative_target(base_dir: &str, part_name: &str) -> String {
    let base_parts: Vec<&str> = base_dir.split('/').filter(|p| !p.is_empty()).collect();
    let target_parts: Vec<&str> = part_name.split('/').filter(|p| !p.is_empty()).collect();

    let mut common = 0usize;
    while common < base_parts.len()
        && common < target_parts.len()
        && base_parts[common] == target_parts[common]
    {
        common += 1;
    }

    let mut out: Vec<&str> = Vec::new();
    for _ in common..base_parts.len() {
        out.push("..");
    }
    out.extend_from_slice(&target_parts[common..]);

    if out.is_empty() {
        ".".to_string()
    } else {
        out.join("/")
    }
}

fn minimal_content_types_xml(sheets: &[SheetMeta]) -> String {
    let mut xml = String::new();
    xml.push_str(r#"<?xml version="1.0" encoding="UTF-8" standalone="yes"?>"#);
    xml.push_str(r#"<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">"#);
    xml.push_str(r#"<Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/>"#);
    xml.push_str(r#"<Default Extension="xml" ContentType="application/xml"/>"#);
    xml.push_str(r#"<Override PartName="/xl/workbook.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet.main+xml"/>"#);
    for sheet_meta in sheets {
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

fn relationship_target_by_type(rels_xml: &[u8], rel_type: &str) -> Result<Option<String>, WriteError> {
    let mut reader = Reader::from_reader(rels_xml);
    reader.config_mut().trim_text(false);
    let mut buf = Vec::new();
    loop {
        match reader.read_event_into(&mut buf)? {
            Event::Start(e) | Event::Empty(e) if e.name().as_ref() == b"Relationship" => {
                let mut type_ = None;
                let mut target = None;
                for attr in e.attributes() {
                    let attr = attr?;
                    match attr.key.as_ref() {
                        b"Type" => type_ = Some(attr.unescape_value()?.into_owned()),
                        b"Target" => target = Some(attr.unescape_value()?.into_owned()),
                        _ => {}
                    }
                }
                if type_.as_deref() == Some(rel_type) {
                    return Ok(target);
                }
            }
            Event::Eof => break,
            _ => {}
        }
        buf.clear();
    }
    Ok(None)
}

fn ensure_workbook_rels_has_relationship(
    parts: &mut BTreeMap<String, Vec<u8>>,
    rel_type: &str,
    target: &str,
) -> Result<(), WriteError> {
    let rels_name = WORKBOOK_RELS_PART;
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

fn patch_workbook_rels_for_sheet_edits(
    parts: &mut BTreeMap<String, Vec<u8>>,
    removed: &[SheetMeta],
    added: &[SheetMeta],
) -> Result<(), WriteError> {
    let rels_name = "xl/_rels/workbook.xml.rels";
    let Some(existing) = parts.get(rels_name).cloned() else {
        return Ok(());
    };

    let remove_ids: HashSet<&str> = removed.iter().map(|m| m.relationship_id.as_str()).collect();

    let mut reader = Reader::from_reader(existing.as_slice());
    reader.config_mut().trim_text(false);
    let mut writer = Writer::new(Vec::with_capacity(existing.len() + added.len() * 128));
    let mut buf = Vec::new();

    let mut skipping = false;
    loop {
        let event = reader.read_event_into(&mut buf)?;
        match event {
            Event::Eof => break,
            Event::Start(ref e) if e.name().as_ref() == b"Relationship" => {
                let mut id = None;
                for attr in e.attributes() {
                    let attr = attr?;
                    if attr.key.as_ref() == b"Id" {
                        id = Some(attr.unescape_value()?.into_owned());
                    }
                }
                if id.as_deref().is_some_and(|id| remove_ids.contains(id)) {
                    skipping = true;
                } else {
                    writer.write_event(Event::Start(e.to_owned()))?;
                }
            }
            Event::Empty(ref e) if e.name().as_ref() == b"Relationship" => {
                let mut id = None;
                for attr in e.attributes() {
                    let attr = attr?;
                    if attr.key.as_ref() == b"Id" {
                        id = Some(attr.unescape_value()?.into_owned());
                    }
                }
                if !id.as_deref().is_some_and(|id| remove_ids.contains(id)) {
                    writer.write_event(Event::Empty(e.to_owned()))?;
                }
            }
            Event::End(ref e) if skipping && e.name().as_ref() == b"Relationship" => {
                skipping = false;
            }
            Event::End(ref e) if e.name().as_ref() == b"Relationships" => {
                for sheet in added {
                    let target = relationship_target_from_workbook(&sheet.path);
                    let mut rel = quick_xml::events::BytesStart::new("Relationship");
                    rel.push_attribute(("Id", sheet.relationship_id.as_str()));
                    rel.push_attribute(("Type", WORKSHEET_REL_TYPE));
                    rel.push_attribute(("Target", target.as_str()));
                    writer.write_event(Event::Empty(rel))?;
                }
                writer.write_event(Event::End(e.to_owned()))?;
            }
            ev if skipping => drop(ev),
            ev => writer.write_event(ev.into_owned())?,
        }
        buf.clear();
    }

    parts.insert(rels_name.to_string(), writer.into_inner());
    Ok(())
}

fn patch_content_types_for_sheet_edits(
    parts: &mut BTreeMap<String, Vec<u8>>,
    removed: &[SheetMeta],
    added: &[SheetMeta],
) -> Result<(), WriteError> {
    let ct_name = "[Content_Types].xml";
    let Some(existing) = parts.get(ct_name).cloned() else {
        return Ok(());
    };

    let removed_parts: HashSet<String> = removed
        .iter()
        .map(|m| {
            if m.path.starts_with('/') {
                m.path.clone()
            } else {
                format!("/{}", m.path)
            }
        })
        .collect();

    let mut reader = Reader::from_reader(existing.as_slice());
    reader.config_mut().trim_text(false);
    let mut writer = Writer::new(Vec::with_capacity(existing.len() + added.len() * 128));
    let mut buf = Vec::new();

    let mut existing_overrides: HashSet<String> = HashSet::new();
    let mut skipping = false;

    loop {
        let event = reader.read_event_into(&mut buf)?;
        match event {
            Event::Eof => break,
            Event::Start(ref e) if e.name().as_ref() == b"Override" => {
                let mut part_name = None;
                for attr in e.attributes() {
                    let attr = attr?;
                    if attr.key.as_ref() == b"PartName" {
                        part_name = Some(attr.unescape_value()?.into_owned());
                    }
                }
                if let Some(name) = &part_name {
                    existing_overrides.insert(name.clone());
                    if removed_parts.contains(name) {
                        skipping = true;
                        continue;
                    }
                }
                writer.write_event(Event::Start(e.to_owned()))?;
            }
            Event::Empty(ref e) if e.name().as_ref() == b"Override" => {
                let mut part_name = None;
                for attr in e.attributes() {
                    let attr = attr?;
                    if attr.key.as_ref() == b"PartName" {
                        part_name = Some(attr.unescape_value()?.into_owned());
                    }
                }
                if let Some(name) = &part_name {
                    existing_overrides.insert(name.clone());
                    if removed_parts.contains(name) {
                        continue;
                    }
                }
                writer.write_event(Event::Empty(e.to_owned()))?;
            }
            Event::End(ref e) if skipping && e.name().as_ref() == b"Override" => {
                skipping = false;
            }
            Event::End(ref e) if e.name().as_ref() == b"Types" => {
                for sheet in added {
                    let part_name = if sheet.path.starts_with('/') {
                        sheet.path.clone()
                    } else {
                        format!("/{}", sheet.path)
                    };
                    if existing_overrides.contains(&part_name) {
                        continue;
                    }
                    let mut override_el = quick_xml::events::BytesStart::new("Override");
                    override_el.push_attribute(("PartName", part_name.as_str()));
                    override_el.push_attribute(("ContentType", WORKSHEET_CONTENT_TYPE));
                    writer.write_event(Event::Empty(override_el))?;
                }
                writer.write_event(Event::End(e.to_owned()))?;
            }
            ev if skipping => drop(ev),
            ev => writer.write_event(ev.into_owned())?,
        }
        buf.clear();
    }

    parts.insert(ct_name.to_string(), writer.into_inner());
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

#[cfg(test)]
mod tests;
