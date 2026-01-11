use std::collections::{BTreeMap, HashMap};
use std::io::{Cursor, Read, Write};

use quick_xml::events::{BytesStart, Event};
use quick_xml::{Reader as XmlReader, Writer as XmlWriter};
use thiserror::Error;

use crate::patch::{
    apply_cell_patches_to_package, apply_cell_patches_to_package_with_styles, WorkbookCellPatches,
};
use crate::pivots::cache_records::{PivotCacheRecordsReader, PivotCacheValue};
use crate::pivots::XlsxPivots;
use crate::recalc_policy::RecalcPolicyError;
use crate::sheet_metadata::{
    parse_sheet_tab_color, parse_workbook_sheets, write_sheet_tab_color, write_workbook_sheets,
    WorkbookSheetInfo,
};
use crate::{DateSystem, RecalcPolicy};
use crate::theme::{parse_theme_palette, ThemePalette};
use formula_model::{CellRef, CellValue, SheetVisibility, StyleTable, TabColor};

#[derive(Debug, Error)]
pub enum XlsxError {
    #[error("zip error: {0}")]
    Zip(#[from] zip::result::ZipError),
    #[error("io error: {0}")]
    Io(#[from] std::io::Error),
    #[error("xml error: {0}")]
    Xml(#[from] quick_xml::Error),
    #[error("xml error: {0}")]
    RoXml(#[from] roxmltree::Error),
    #[error("utf-8 error: {0}")]
    Utf8(#[from] std::string::FromUtf8Error),
    #[error("xml attribute error: {0}")]
    Attr(#[from] quick_xml::events::attributes::AttrError),
    #[error("missing required attribute: {0}")]
    MissingAttr(&'static str),
    #[error("missing xlsx part: {0}")]
    MissingPart(String),
    #[error("invalid xlsx: {0}")]
    Invalid(String),
    #[error("invalid sheetId value")]
    InvalidSheetId,
    #[error("hyperlink error: {0}")]
    Hyperlink(String),
    #[error(transparent)]
    StreamingPatch(#[from] Box<crate::streaming::StreamingPatchError>),
}

impl From<crate::streaming::StreamingPatchError> for XlsxError {
    fn from(err: crate::streaming::StreamingPatchError) -> Self {
        Self::StreamingPatch(Box::new(err))
    }
}

/// Resolved metadata for a workbook sheet and its corresponding worksheet part.
#[derive(Debug, Clone, PartialEq, Eq)]
pub struct WorksheetPartInfo {
    pub name: String,
    pub sheet_id: u32,
    pub rel_id: String,
    pub visibility: SheetVisibility,
    /// ZIP entry name for the worksheet XML (e.g. `xl/worksheets/sheet1.xml`).
    pub worksheet_part: String,
}

/// Select a target worksheet for a cell patch.
#[derive(Debug, Clone, PartialEq, Eq)]
pub enum CellPatchSheet {
    /// Identify the sheet by workbook sheet name (e.g. `"Sheet1"`).
    SheetName(String),
    /// Identify the sheet by its worksheet XML part name (e.g. `"xl/worksheets/sheet1.xml"`).
    WorksheetPart(String),
}

/// A single cell edit to apply to an [`XlsxPackage`].
#[derive(Debug, Clone, PartialEq)]
pub struct CellPatch {
    pub sheet: CellPatchSheet,
    pub cell: CellRef,
    pub value: CellValue,
    /// Optional formula to write into the `<f>` element. Leading `=` is permitted.
    pub formula: Option<String>,
}

impl CellPatch {
    pub fn new(
        sheet: CellPatchSheet,
        cell: CellRef,
        value: CellValue,
        formula: Option<String>,
    ) -> Self {
        Self {
            sheet,
            cell,
            value,
            formula,
        }
    }

    pub fn for_sheet_name(
        sheet_name: impl Into<String>,
        cell: CellRef,
        value: CellValue,
        formula: Option<String>,
    ) -> Self {
        Self::new(CellPatchSheet::SheetName(sheet_name.into()), cell, value, formula)
    }

    pub fn for_worksheet_part(
        worksheet_part: impl Into<String>,
        cell: CellRef,
        value: CellValue,
        formula: Option<String>,
    ) -> Self {
        Self::new(
            CellPatchSheet::WorksheetPart(worksheet_part.into()),
            cell,
            value,
            formula,
        )
    }
}

impl From<RecalcPolicyError> for XlsxError {
    fn from(err: RecalcPolicyError) -> Self {
        match err {
            RecalcPolicyError::Io(err) => XlsxError::Io(err),
            RecalcPolicyError::Xml(err) => XlsxError::Xml(err),
            RecalcPolicyError::XmlAttr(err) => XlsxError::Attr(err),
        }
    }
}

/// In-memory representation of an XLSX/XLSM package as a map of part name -> bytes.
///
/// We keep the API minimal to support macro preservation; a full model will
/// eventually build on top of this.
#[derive(Debug, Clone)]
pub struct XlsxPackage {
    parts: BTreeMap<String, Vec<u8>>,
}

impl XlsxPackage {
    pub fn from_bytes(bytes: &[u8]) -> Result<Self, XlsxError> {
        let reader = Cursor::new(bytes);
        let mut zip = zip::ZipArchive::new(reader)?;

        let mut parts = BTreeMap::new();
        for i in 0..zip.len() {
            let mut file = zip.by_index(i)?;
            if !file.is_file() {
                continue;
            }

            let mut buf = Vec::new();
            file.read_to_end(&mut buf)?;
            parts.insert(file.name().to_string(), buf);
        }

        Ok(Self { parts })
    }

    pub fn part(&self, name: &str) -> Option<&[u8]> {
        self.parts.get(name).map(|v| v.as_slice())
    }

    pub fn parts(&self) -> impl Iterator<Item = (&str, &[u8])> {
        self.parts
            .iter()
            .map(|(name, bytes)| (name.as_str(), bytes.as_slice()))
    }

    pub fn part_names(&self) -> impl Iterator<Item = &str> {
        self.parts.keys().map(String::as_str)
    }

    /// Borrow the raw part map (useful for higher-fidelity operations).
    pub fn parts_map(&self) -> &BTreeMap<String, Vec<u8>> {
        &self.parts
    }

    pub fn parts_map_mut(&mut self) -> &mut BTreeMap<String, Vec<u8>> {
        &mut self.parts
    }

    pub fn set_part(&mut self, name: impl Into<String>, bytes: Vec<u8>) {
        self.parts.insert(name.into(), bytes);
    }

    pub fn vba_project_bin(&self) -> Option<&[u8]> {
        self.part("xl/vbaProject.bin")
    }

    /// Parse the workbook theme palette from `xl/theme/theme1.xml` (if present).
    pub fn theme_palette(&self) -> Result<Option<ThemePalette>, XlsxError> {
        let Some(theme_xml) = self.part("xl/theme/theme1.xml") else {
            return Ok(None);
        };
        Ok(Some(parse_theme_palette(theme_xml)?))
    }

    pub fn write_to_bytes(&self) -> Result<Vec<u8>, XlsxError> {
        let mut buf = Vec::new();
        self.write_to(&mut buf)?;
        Ok(buf)
    }

    pub fn write_to<W: Write>(&self, mut w: W) -> Result<(), XlsxError> {
        let mut parts = self.parts.clone();
        if parts.contains_key("xl/vbaProject.bin") {
            ensure_xlsm_content_types(&mut parts)?;
            ensure_workbook_rels_has_vba(&mut parts)?;
        }

        let cursor = Cursor::new(Vec::new());
        let mut zip = zip::ZipWriter::new(cursor);
        let options = zip::write::FileOptions::<()>::default()
            .compression_method(zip::CompressionMethod::Deflated);

        for (name, bytes) in parts {
            zip.start_file(name, options)?;
            zip.write_all(&bytes)?;
        }

        let cursor = zip.finish()?;
        w.write_all(&cursor.into_inner())?;
        Ok(())
    }

    /// Return the ordered workbook sheets with their resolved worksheet part paths.
    ///
    /// This reads `xl/workbook.xml` for the `<sheet>` list and `xl/_rels/workbook.xml.rels`
    /// to resolve each sheet's `r:id` relationship to a concrete worksheet XML part name.
    pub fn worksheet_parts(&self) -> Result<Vec<WorksheetPartInfo>, XlsxError> {
        let sheets = self.workbook_sheets()?;

        let rels_bytes = self
            .part("xl/_rels/workbook.xml.rels")
            .ok_or_else(|| XlsxError::MissingPart("xl/_rels/workbook.xml.rels".to_string()))?;
        let relationships = crate::openxml::parse_relationships(rels_bytes)?;
        let rel_by_id: HashMap<String, crate::openxml::Relationship> = relationships
            .into_iter()
            .map(|rel| (rel.id.clone(), rel))
            .collect();

        let mut out = Vec::with_capacity(sheets.len());
        for sheet in sheets {
            let rel = rel_by_id.get(&sheet.rel_id).ok_or_else(|| {
                XlsxError::Invalid(format!("missing relationship for {}", sheet.rel_id))
            })?;
            let worksheet_part = crate::path::resolve_target("xl/workbook.xml", &rel.target);
            out.push(WorksheetPartInfo {
                name: sheet.name,
                sheet_id: sheet.sheet_id,
                rel_id: sheet.rel_id,
                visibility: sheet.visibility,
                worksheet_part,
            });
        }
        Ok(out)
    }

    /// Apply a set of cell edits to the package and return the updated ZIP bytes.
    ///
    /// This uses the streaming patch pipeline, which rewrites the targeted worksheet parts and
    /// updates dependent workbook parts when needed (for example `xl/sharedStrings.xml` and the
    /// calcChain/full-calc settings after formula edits).
    pub fn apply_cell_patches_to_bytes(&self, patches: &[CellPatch]) -> Result<Vec<u8>, XlsxError> {
        let mut sheet_name_to_part: HashMap<String, String> = HashMap::new();
        if patches
            .iter()
            .any(|p| matches!(p.sheet, CellPatchSheet::SheetName(_)))
        {
            for entry in self.worksheet_parts()? {
                sheet_name_to_part.insert(entry.name, entry.worksheet_part);
            }
        }

        let mut patches_by_part: HashMap<String, BTreeMap<(u32, u32), crate::streaming::WorksheetCellPatch>> =
            HashMap::new();

        for patch in patches {
            let worksheet_part = match &patch.sheet {
                CellPatchSheet::WorksheetPart(part) => part.clone(),
                CellPatchSheet::SheetName(name) => sheet_name_to_part.get(name).cloned().ok_or_else(|| {
                    XlsxError::Invalid(format!("unknown sheet name {name}"))
                })?,
            };

            patches_by_part
                .entry(worksheet_part.clone())
                .or_default()
                .insert(
                    (patch.cell.row, patch.cell.col),
                    crate::streaming::WorksheetCellPatch::new(
                        worksheet_part,
                        patch.cell,
                        patch.value.clone(),
                        patch.formula.clone(),
                    ),
                );
        }

        let mut patches_by_part: HashMap<String, Vec<crate::streaming::WorksheetCellPatch>> = patches_by_part
            .into_iter()
            .map(|(part, cells)| (part, cells.into_values().collect()))
            .collect();
        for patches in patches_by_part.values_mut() {
            patches.sort_by_key(|p| (p.cell.row, p.cell.col));
        }

        for part in patches_by_part.keys() {
            if !self.parts.contains_key(part) {
                return Err(crate::streaming::StreamingPatchError::MissingWorksheetPart(
                    part.clone(),
                )
                .into());
            }
        }

        // Route through the full streaming patch pipeline (sharedStrings-aware + recalc-policy
        // aware) rather than directly rewriting worksheet XML parts.
        let input_bytes = self.write_to_bytes()?;
        let mut streaming_patches = Vec::with_capacity(patches.len());
        for patches in patches_by_part.values() {
            streaming_patches.extend_from_slice(patches);
        }

        let mut out = Cursor::new(Vec::new());
        crate::streaming::patch_xlsx_streaming(
            Cursor::new(input_bytes),
            &mut out,
            &streaming_patches,
        )?;
        Ok(out.into_inner())
    }

    /// Parse pivot-related parts (pivot tables + pivot caches) from the package.
    ///
    /// This is a lightweight metadata parser; the raw XML parts remain preserved
    /// verbatim in the package.
    pub fn pivots(&self) -> Result<XlsxPivots, XlsxError> {
        XlsxPivots::parse_from_entries(&self.parts)
    }

    /// Create a streaming reader for a pivot cache records part
    /// (e.g. `xl/pivotCache/pivotCacheRecords1.xml`).
    pub fn pivot_cache_records<'a>(
        &'a self,
        part_name: &str,
    ) -> Result<PivotCacheRecordsReader<'a>, XlsxError> {
        let part_name = part_name.strip_prefix('/').unwrap_or(part_name);
        let bytes = self
            .part(part_name)
            .ok_or_else(|| XlsxError::MissingPart(part_name.to_string()))?;
        Ok(PivotCacheRecordsReader::new(bytes))
    }

    /// Parse all `pivotCacheRecords*.xml` parts in the package into memory.
    ///
    /// Prefer [`Self::pivot_cache_records`] for large caches.
    pub fn pivot_cache_records_all(&self) -> BTreeMap<String, Vec<Vec<PivotCacheValue>>> {
        let mut out = BTreeMap::new();
        for (name, bytes) in &self.parts {
            if name.starts_with("xl/pivotCache/")
                && name.contains("pivotCacheRecords")
                && name.ends_with(".xml")
            {
                let mut reader = PivotCacheRecordsReader::new(bytes);
                out.insert(name.clone(), reader.parse_all_records());
            }
        }
        out
    }

    /// Parse the ordered list of workbook sheets from `xl/workbook.xml`.
    pub fn workbook_sheets(&self) -> Result<Vec<WorkbookSheetInfo>, XlsxError> {
        let workbook_xml = self
            .part("xl/workbook.xml")
            .ok_or_else(|| XlsxError::MissingPart("xl/workbook.xml".to_string()))?;
        let workbook_xml = String::from_utf8(workbook_xml.to_vec())?;
        parse_workbook_sheets(&workbook_xml)
    }

    /// Set the workbook date system (`workbookPr/@date1904`) inside `xl/workbook.xml`.
    ///
    /// This is required for correct serial date interpretation when opening the workbook in Excel
    /// and for aligning formula evaluation semantics (1900 vs 1904) during round-trip edits.
    pub fn set_workbook_date_system(&mut self, date_system: DateSystem) -> Result<(), XlsxError> {
        let workbook_xml = self
            .parts
            .get("xl/workbook.xml")
            .cloned()
            .ok_or_else(|| XlsxError::MissingPart("xl/workbook.xml".to_string()))?;
        let updated = workbook_xml_set_date_system(&workbook_xml, date_system)?;
        self.parts.insert("xl/workbook.xml".to_string(), updated);
        Ok(())
    }

    /// Rewrite the `<sheets>` list in `xl/workbook.xml` to match `sheets`.
    pub fn set_workbook_sheets(&mut self, sheets: &[WorkbookSheetInfo]) -> Result<(), XlsxError> {
        let workbook_xml = self
            .part("xl/workbook.xml")
            .ok_or_else(|| XlsxError::MissingPart("xl/workbook.xml".to_string()))?;
        let workbook_xml = String::from_utf8(workbook_xml.to_vec())?;
        let updated = write_workbook_sheets(&workbook_xml, sheets)?;
        self.set_part("xl/workbook.xml", updated.into_bytes());
        Ok(())
    }

    /// Read a worksheet tab color from a worksheet part (e.g. `xl/worksheets/sheet1.xml`).
    pub fn worksheet_tab_color(&self, worksheet_part: &str) -> Result<Option<TabColor>, XlsxError> {
        let xml = self
            .part(worksheet_part)
            .ok_or_else(|| XlsxError::MissingPart(worksheet_part.to_string()))?;
        let xml = String::from_utf8(xml.to_vec())?;
        parse_sheet_tab_color(&xml)
    }

    /// Update (or remove) a worksheet tab color in a worksheet part.
    pub fn set_worksheet_tab_color(
        &mut self,
        worksheet_part: &str,
        tab_color: Option<&TabColor>,
    ) -> Result<(), XlsxError> {
        let xml = self
            .part(worksheet_part)
            .ok_or_else(|| XlsxError::MissingPart(worksheet_part.to_string()))?;
        let xml = String::from_utf8(xml.to_vec())?;
        let updated = write_sheet_tab_color(&xml, tab_color)?;
        self.set_part(worksheet_part.to_string(), updated.into_bytes());
        Ok(())
    }

    /// Apply a set of cell edits to the existing workbook package.
    ///
    /// This is a high-fidelity edit pipeline intended for "edit existing XLSX/XLSM"
    /// scenarios (e.g. the desktop app save path). Only the affected worksheet parts
    /// (plus `xl/sharedStrings.xml` / `xl/workbook.xml` when required) are rewritten;
    /// every unrelated part is preserved byte-for-byte.
    pub fn apply_cell_patches(&mut self, patches: &WorkbookCellPatches) -> Result<(), XlsxError> {
        self.apply_cell_patches_with_recalc_policy(patches, RecalcPolicy::default())
    }

    /// Apply a set of cell edits to the existing workbook package using the provided
    /// [`RecalcPolicy`].
    pub fn apply_cell_patches_with_recalc_policy(
        &mut self,
        patches: &WorkbookCellPatches,
        recalc_policy: RecalcPolicy,
    ) -> Result<(), XlsxError> {
        apply_cell_patches_to_package(self, patches, recalc_policy)
    }

    /// Apply cell edits that reference `formula_model` style IDs.
    ///
    /// This behaves like [`Self::apply_cell_patches`], but allows patches to specify cell styles
    /// via `style_id` and updates `styles.xml` deterministically when new styles are introduced.
    pub fn apply_cell_patches_with_styles(
        &mut self,
        patches: &WorkbookCellPatches,
        style_table: &StyleTable,
    ) -> Result<(), XlsxError> {
        apply_cell_patches_to_package_with_styles(self, patches, style_table, RecalcPolicy::default())
    }

    /// Remove macro-related parts and relationships from the package.
    ///
    /// This is used when saving a macro-enabled workbook (`.xlsm`) as `.xlsx`.
    pub fn remove_vba_project(&mut self) -> Result<(), XlsxError> {
        crate::macro_strip::strip_macros(&mut self.parts)
    }
}

fn workbook_xml_set_date_system(
    workbook_xml: &[u8],
    date_system: DateSystem,
) -> Result<Vec<u8>, XlsxError> {
    let has_workbook_pr = workbook_xml
        .windows(b"workbookPr".len())
        .any(|w| w == b"workbookPr");

    let mut reader = XmlReader::from_reader(workbook_xml);
    reader.config_mut().trim_text(false);
    let mut writer = XmlWriter::new(Vec::with_capacity(workbook_xml.len() + 64));

    let mut buf = Vec::new();
    let mut skipping_workbook_pr = false;

    loop {
        let event = reader.read_event_into(&mut buf)?;
        match event {
            Event::Start(ref e) if local_name(e.name().as_ref()) == b"workbook" => {
                writer.write_event(Event::Start(e.to_owned()))?;
                if date_system == DateSystem::V1904 && !has_workbook_pr {
                    let mut wb_pr = BytesStart::new("workbookPr");
                    wb_pr.push_attribute(("date1904", "1"));
                    writer.write_event(Event::Empty(wb_pr))?;
                }
            }
            Event::Empty(ref e) if local_name(e.name().as_ref()) == b"workbookPr" => {
                writer.write_event(Event::Empty(patched_workbook_pr(e, date_system)?))?;
            }
            Event::Start(ref e) if local_name(e.name().as_ref()) == b"workbookPr" => {
                skipping_workbook_pr = true;
                writer.write_event(Event::Empty(patched_workbook_pr(e, date_system)?))?;
            }
            Event::End(ref e)
                if skipping_workbook_pr && local_name(e.name().as_ref()) == b"workbookPr" =>
            {
                skipping_workbook_pr = false;
            }
            Event::Eof => break,
            ev if skipping_workbook_pr => drop(ev),
            other => writer.write_event(other.into_owned())?,
        }

        buf.clear();
    }

    Ok(writer.into_inner())
}

fn patched_workbook_pr(e: &BytesStart<'_>, date_system: DateSystem) -> Result<BytesStart<'static>, XlsxError> {
    let name = e.name();
    let mut wb_pr =
        BytesStart::new(std::str::from_utf8(name.as_ref()).unwrap_or("workbookPr"));
    let mut had_date1904 = false;
    for attr in e.attributes() {
        let attr = attr?;
        if attr.key.as_ref() == b"date1904" {
            had_date1904 = true;
            continue;
        }
        wb_pr.push_attribute((attr.key.as_ref(), attr.value.as_ref()));
    }

    match date_system {
        DateSystem::V1900 => {
            if had_date1904 {
                wb_pr.push_attribute(("date1904", "0"));
            }
        }
        DateSystem::V1904 => wb_pr.push_attribute(("date1904", "1")),
    }

    Ok(wb_pr.into_owned())
}

fn local_name(name: &[u8]) -> &[u8] {
    name.rsplit(|&b| b == b':').next().unwrap_or(name)
}

fn ensure_xlsm_content_types(parts: &mut BTreeMap<String, Vec<u8>>) -> Result<(), XlsxError> {
    let ct_name = "[Content_Types].xml";
    let Some(existing) = parts.get(ct_name).cloned() else {
        // We don't attempt to synthesize a full content types file; macro
        // preservation in this minimal crate assumes an existing workbook.
        return Ok(());
    };

    let mut xml = String::from_utf8(existing)?;

    if !xml.contains("vbaProject.bin") {
        // Insert before closing </Types>
        if let Some(idx) = xml.rfind("</Types>") {
            let insert = r#"<Override PartName="/xl/vbaProject.bin" ContentType="application/vnd.ms-office.vbaProject"/>"#;
            xml.insert_str(idx, insert);
        }
    }

    // Ensure workbook content type reflects macro-enabled if we can find it.
    if xml.contains(r#"PartName="/xl/workbook.xml""#)
        && !xml.contains("application/vnd.ms-excel.sheet.macroEnabled.main+xml")
    {
        xml = xml.replace(
            r#"ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet.main+xml""#,
            r#"ContentType="application/vnd.ms-excel.sheet.macroEnabled.main+xml""#,
        );
    }

    parts.insert(ct_name.to_string(), xml.into_bytes());
    Ok(())
}

fn ensure_workbook_rels_has_vba(parts: &mut BTreeMap<String, Vec<u8>>) -> Result<(), XlsxError> {
    let rels_name = "xl/_rels/workbook.xml.rels";
    let Some(existing) = parts.get(rels_name).cloned() else {
        return Ok(());
    };

    let mut xml = String::from_utf8(existing)?;
    let rel_type = "http://schemas.microsoft.com/office/2006/relationships/vbaProject";
    if xml.contains(rel_type) {
        parts.insert(rels_name.to_string(), xml.into_bytes());
        return Ok(());
    }

    let next_rid = next_relationship_id(&xml);
    let rel =
        format!(r#"<Relationship Id="rId{next_rid}" Type="{rel_type}" Target="vbaProject.bin"/>"#);

    if let Some(idx) = xml.rfind("</Relationships>") {
        xml.insert_str(idx, &rel);
    }
    parts.insert(rels_name.to_string(), xml.into_bytes());
    Ok(())
}

fn next_relationship_id(xml: &str) -> u32 {
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
mod tests {
    use super::*;
    use pretty_assertions::assert_eq;

    fn build_package(files: &[(&str, &[u8])]) -> Vec<u8> {
        let cursor = Cursor::new(Vec::new());
        let mut zip = zip::ZipWriter::new(cursor);
        let options = zip::write::FileOptions::<()>::default()
            .compression_method(zip::CompressionMethod::Deflated);

        for (name, bytes) in files {
            zip.start_file(*name, options).unwrap();
            zip.write_all(bytes).unwrap();
        }

        zip.finish().unwrap().into_inner()
    }

    fn build_minimal_package() -> Vec<u8> {
        let workbook_xml = r#"<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<workbook xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main"
  xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">
  <sheets>
    <sheet name="Sheet1" sheetId="1" r:id="rId1"/>
  </sheets>
</workbook>"#;

        let worksheet_xml = r#"<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main"></worksheet>"#;

        let cursor = Cursor::new(Vec::new());
        let mut zip = zip::ZipWriter::new(cursor);
        let options = zip::write::FileOptions::<()>::default()
            .compression_method(zip::CompressionMethod::Deflated);

        zip.start_file("xl/workbook.xml", options).unwrap();
        zip.write_all(workbook_xml.as_bytes()).unwrap();

        zip.start_file("xl/worksheets/sheet1.xml", options).unwrap();
        zip.write_all(worksheet_xml.as_bytes()).unwrap();

        zip.finish().unwrap().into_inner()
    }

    fn load_fixture() -> Vec<u8> {
        std::fs::read(concat!(
            env!("CARGO_MANIFEST_DIR"),
            "/../../fixtures/xlsx/macros/basic.xlsm"
        ))
        .expect("fixture exists")
    }

    #[test]
    fn round_trip_preserves_vba_project_bin_bytes() {
        let fixture = load_fixture();
        let pkg = XlsxPackage::from_bytes(&fixture).expect("read pkg");

        let original_bin = pkg
            .vba_project_bin()
            .expect("vbaProject.bin present")
            .to_vec();

        let written = pkg.write_to_bytes().expect("write pkg");
        let pkg2 = XlsxPackage::from_bytes(&written).expect("read pkg2");
        let roundtrip_bin = pkg2
            .vba_project_bin()
            .expect("vbaProject.bin present in roundtrip");

        assert_eq!(original_bin, roundtrip_bin);
    }

    #[test]
    fn ensures_content_types_and_relationships_for_xlsm() {
        let fixture = load_fixture();
        let pkg = XlsxPackage::from_bytes(&fixture).expect("read pkg");
        let written = pkg.write_to_bytes().expect("write");
        let pkg2 = XlsxPackage::from_bytes(&written).expect("read");

        let ct = std::str::from_utf8(pkg2.part("[Content_Types].xml").unwrap()).unwrap();
        assert!(ct.contains("application/vnd.ms-office.vbaProject"));

        let rels = std::str::from_utf8(pkg2.part("xl/_rels/workbook.xml.rels").unwrap()).unwrap();
        assert!(rels.contains("http://schemas.microsoft.com/office/2006/relationships/vbaProject"));
    }

    #[test]
    fn parses_vba_project_from_package() {
        let fixture = load_fixture();
        let pkg = XlsxPackage::from_bytes(&fixture).expect("read pkg");
        let project = pkg
            .vba_project()
            .expect("parse vba project")
            .expect("vba project present");

        assert_eq!(project.name.as_deref(), Some("VBAProject"));
        let module = project
            .modules
            .iter()
            .find(|m| m.name == "Module1")
            .expect("Module1 present");
        assert!(module.code.contains("Sub Hello"));
        assert_eq!(
            module.attributes.get("VB_Name").map(String::as_str),
            Some("Module1")
        );
    }

    #[test]
    fn package_exposes_sheet_list_and_tab_color_helpers() {
        let bytes = build_minimal_package();
        let mut pkg = XlsxPackage::from_bytes(&bytes).expect("read pkg");

        let sheets = pkg.workbook_sheets().expect("parse sheets");
        assert_eq!(sheets.len(), 1);
        assert_eq!(sheets[0].name, "Sheet1");
        assert_eq!(sheets[0].sheet_id, 1);
        assert_eq!(sheets[0].rel_id, "rId1");
        assert_eq!(
            sheets[0].visibility,
            formula_model::SheetVisibility::Visible
        );

        let mut updated = sheets.clone();
        updated[0].name = "Renamed".to_string();
        pkg.set_workbook_sheets(&updated).expect("write sheets");
        let renamed = pkg.workbook_sheets().expect("parse renamed sheets");
        assert_eq!(renamed[0].name, "Renamed");

        let color = TabColor::rgb("FFFF0000");
        pkg.set_worksheet_tab_color("xl/worksheets/sheet1.xml", Some(&color))
            .expect("set tab color");
        let parsed = pkg
            .worksheet_tab_color("xl/worksheets/sheet1.xml")
            .expect("parse tab color");
        assert_eq!(parsed.unwrap().rgb.as_deref(), Some("FFFF0000"));

        pkg.set_worksheet_tab_color("xl/worksheets/sheet1.xml", None)
            .expect("remove tab color");
        assert_eq!(
            pkg.worksheet_tab_color("xl/worksheets/sheet1.xml")
                .expect("parse tab color"),
            None
        );
    }

    #[test]
    fn remove_vba_project_strips_vba_parts() {
        let fixture = load_fixture();
        let mut pkg = XlsxPackage::from_bytes(&fixture).expect("read pkg");

        assert!(pkg.vba_project_bin().is_some());
        pkg.remove_vba_project().expect("remove vba project");

        let written = pkg.write_to_bytes().expect("write pkg");
        let pkg2 = XlsxPackage::from_bytes(&written).expect("read pkg2");

        assert!(pkg2.vba_project_bin().is_none());

        let ct = std::str::from_utf8(pkg2.part("[Content_Types].xml").unwrap()).unwrap();
        assert!(!ct.contains("vbaProject.bin"));
        assert!(!ct.contains("macroEnabled.main+xml"));

        let rels = std::str::from_utf8(pkg2.part("xl/_rels/workbook.xml.rels").unwrap()).unwrap();
        assert!(!rels.contains("relationships/vbaProject"));
    }

    fn build_synthetic_macro_package() -> Vec<u8> {
        let content_types = r#"<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">
  <Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/>
  <Default Extension="xml" ContentType="application/xml"/>
  <Default Extension="bin" ContentType="application/vnd.openxmlformats-officedocument.oleObject"/>
  <Default Extension="png" ContentType="image/png"/>
  <Override PartName="/xl/workbook.xml" ContentType="application/vnd.ms-excel.sheet.macroEnabled.main+xml"/>
  <Override PartName="/xl/worksheets/sheet1.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.worksheet+xml"/>
  <Override PartName="/xl/vbaProject.bin" ContentType="application/vnd.ms-office.vbaProject"/>
  <Override PartName="/xl/vbaProjectSignature.bin" ContentType="application/vnd.ms-office.vbaProjectSignature"/>
  <Override PartName="/xl/vbaData.xml" ContentType="application/vnd.ms-office.vbaData+xml"/>
  <Override PartName="/customUI/customUI.xml" ContentType="application/xml"/>
  <Override PartName="/customUI/customUI14.xml" ContentType="application/xml"/>
  <Override PartName="/xl/activeX/activeX1.xml" ContentType="application/vnd.ms-office.activeX+xml"/>
  <Override PartName="/xl/ctrlProps/ctrlProp1.xml" ContentType="application/vnd.ms-office.activeX+xml"/>
  <Override PartName="/xl/embeddings/oleObject1.bin" ContentType="application/vnd.openxmlformats-officedocument.oleObject"/>
</Types>"#;

        let root_rels = r#"<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
  <Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="xl/workbook.xml"/>
  <Relationship Id="rId2" Type="http://schemas.microsoft.com/office/2006/relationships/ui/extensibility" Target="customUI/customUI.xml"/>
  <Relationship Id="rId3" Type="http://schemas.microsoft.com/office/2007/relationships/ui/extensibility" Target="customUI/customUI14.xml"/>
</Relationships>"#;

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
  <Relationship Id="rId2" Type="http://schemas.microsoft.com/office/2006/relationships/vbaProject" Target="vbaProject.bin"/>
</Relationships>"#;

        let worksheet_xml = r#"<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main"></worksheet>"#;

        let sheet_rels = r#"<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
  <Relationship Id="rId1" Type="http://schemas.microsoft.com/office/2006/relationships/activeXControl" Target="../activeX/activeX1.xml#_x0000_s1025"/>
  <Relationship Id="rId2" Type="http://schemas.microsoft.com/office/2006/relationships/ctrlProp" Target="../ctrlProps/ctrlProp1.xml"/>
</Relationships>"#;

        let vba_rels = r#"<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
  <Relationship Id="rId1" Type="http://schemas.microsoft.com/office/2006/relationships/vbaProjectSignature" Target="vbaProjectSignature.bin"/>
</Relationships>"#;

        let custom_ui_xml = r#"<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<customUI xmlns="http://schemas.microsoft.com/office/2006/01/customui"></customUI>"#;

        let custom_ui_rels = r#"<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
  <Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/image" Target="image1.png"/>
</Relationships>"#;

        let activex_xml = r#"<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<ax:ocx xmlns:ax="http://schemas.microsoft.com/office/2006/activeX"></ax:ocx>"#;

        let activex_rels = r#"<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
  <Relationship Id="rId1" Type="http://schemas.microsoft.com/office/2006/relationships/activeXControlBinary" Target="activeX1.bin"/>
  <Relationship Id="rId2" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/oleObject" Target="../embeddings/oleObject1.bin"/>
</Relationships>"#;

        let ctrl_props_xml = r#"<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<ctrlProp xmlns="http://schemas.microsoft.com/office/2006/activeX"></ctrlProp>"#;

        let cursor = Cursor::new(Vec::new());
        let mut zip = zip::ZipWriter::new(cursor);
        let options = zip::write::FileOptions::<()>::default()
            .compression_method(zip::CompressionMethod::Deflated);

        zip.start_file("[Content_Types].xml", options).unwrap();
        zip.write_all(content_types.as_bytes()).unwrap();

        zip.start_file("_rels/.rels", options).unwrap();
        zip.write_all(root_rels.as_bytes()).unwrap();

        zip.start_file("xl/workbook.xml", options).unwrap();
        zip.write_all(workbook_xml.as_bytes()).unwrap();

        zip.start_file("xl/_rels/workbook.xml.rels", options).unwrap();
        zip.write_all(workbook_rels.as_bytes()).unwrap();

        zip.start_file("xl/worksheets/sheet1.xml", options).unwrap();
        zip.write_all(worksheet_xml.as_bytes()).unwrap();

        zip.start_file("xl/worksheets/_rels/sheet1.xml.rels", options)
            .unwrap();
        zip.write_all(sheet_rels.as_bytes()).unwrap();

        zip.start_file("customUI/customUI.xml", options).unwrap();
        zip.write_all(custom_ui_xml.as_bytes()).unwrap();

        zip.start_file("customUI/customUI14.xml", options).unwrap();
        zip.write_all(custom_ui_xml.as_bytes()).unwrap();

        zip.start_file("customUI/_rels/customUI.xml.rels", options)
            .unwrap();
        zip.write_all(custom_ui_rels.as_bytes()).unwrap();

        zip.start_file("customUI/image1.png", options).unwrap();
        zip.write_all(b"not-a-real-png").unwrap();

        zip.start_file("xl/vbaProject.bin", options).unwrap();
        zip.write_all(b"fake-vba-project").unwrap();

        zip.start_file("xl/_rels/vbaProject.bin.rels", options).unwrap();
        zip.write_all(vba_rels.as_bytes()).unwrap();

        zip.start_file("xl/vbaProjectSignature.bin", options).unwrap();
        zip.write_all(b"fake-signature").unwrap();

        zip.start_file("xl/vbaData.xml", options).unwrap();
        zip.write_all(b"<vbaData/>").unwrap();

        zip.start_file("xl/activeX/activeX1.xml", options).unwrap();
        zip.write_all(activex_xml.as_bytes()).unwrap();

        zip.start_file("xl/activeX/_rels/activeX1.xml.rels", options)
            .unwrap();
        zip.write_all(activex_rels.as_bytes()).unwrap();

        zip.start_file("xl/activeX/activeX1.bin", options).unwrap();
        zip.write_all(b"activex-binary").unwrap();

        zip.start_file("xl/embeddings/oleObject1.bin", options).unwrap();
        zip.write_all(b"ole-embedding").unwrap();

        zip.start_file("xl/ctrlProps/ctrlProp1.xml", options).unwrap();
        zip.write_all(ctrl_props_xml.as_bytes()).unwrap();

        zip.finish().unwrap().into_inner()
    }

    #[test]
    fn remove_vba_project_strips_macro_part_graph_and_repairs_relationships() {
        let bytes = build_synthetic_macro_package();
        let mut pkg = XlsxPackage::from_bytes(&bytes).expect("read pkg");
        pkg.remove_vba_project().expect("strip macros");

        // Round-trip through ZIP writing to ensure we didn't leave any dangling references.
        let written = pkg.write_to_bytes().expect("write stripped pkg");
        let pkg2 = XlsxPackage::from_bytes(&written).expect("read stripped pkg");

        for removed in [
            "xl/vbaProject.bin",
            "xl/_rels/vbaProject.bin.rels",
            "xl/vbaProjectSignature.bin",
            "xl/vbaData.xml",
            "customUI/customUI.xml",
            "customUI/customUI14.xml",
            "customUI/_rels/customUI.xml.rels",
            "customUI/image1.png",
            "xl/activeX/activeX1.xml",
            "xl/activeX/_rels/activeX1.xml.rels",
            "xl/activeX/activeX1.bin",
            "xl/ctrlProps/ctrlProp1.xml",
            // Child part referenced only by the removed ActiveX graph.
            "xl/embeddings/oleObject1.bin",
        ] {
            assert!(
                pkg2.part(removed).is_none(),
                "expected {removed} to be removed"
            );
        }

        let ct = std::str::from_utf8(pkg2.part("[Content_Types].xml").unwrap()).unwrap();
        assert!(!ct.contains("macroEnabled.main+xml"));
        assert!(!ct.contains("vbaProject.bin"));
        assert!(!ct.contains("customUI/customUI.xml"));
        assert!(!ct.contains("customUI/customUI14.xml"));
        assert!(!ct.contains("activeX1.xml"));
        assert!(!ct.contains("ctrlProp1.xml"));

        // Relationship parts should no longer mention the stripped macro graph.
        for (name, bytes) in pkg2.parts() {
            if !name.ends_with(".rels") {
                continue;
            }
            let xml = std::str::from_utf8(bytes).unwrap();
            assert!(!xml.contains("vbaProject"));
            assert!(!xml.contains("customUI"));
            assert!(!xml.contains("activeX"));
            assert!(!xml.contains("ctrlProps"));
        }

        crate::macro_strip::validate_opc_relationships(pkg2.parts_map())
            .expect("stripped package relationships are consistent");
    }

    #[test]
    fn remove_vba_project_strips_worksheet_rid_references() {
        let worksheet_xml = r#"<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main"
 xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">
  <controls>
    <control r:id="rIdA"/>
  </controls>
</worksheet>"#;

        let rels_xml = r#"<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
  <Relationship Id="rIdA" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/control" Target="../activeX/activeX1.xml"/>
</Relationships>"#;

        let bytes = build_package(&[
            ("xl/worksheets/sheet1.xml", worksheet_xml.as_bytes()),
            ("xl/worksheets/_rels/sheet1.xml.rels", rels_xml.as_bytes()),
            ("xl/activeX/activeX1.xml", b"<activeX/>"),
        ]);

        let mut pkg = XlsxPackage::from_bytes(&bytes).expect("read pkg");
        pkg.remove_vba_project().expect("strip macros");

        let updated_rels =
            std::str::from_utf8(pkg.part("xl/worksheets/_rels/sheet1.xml.rels").unwrap()).unwrap();
        assert!(!updated_rels.contains("rIdA"));

        let updated_sheet =
            std::str::from_utf8(pkg.part("xl/worksheets/sheet1.xml").unwrap()).unwrap();
        assert!(!updated_sheet.contains("rIdA"));
        assert!(!updated_sheet.contains("<control r:id"));

        assert!(pkg.part("xl/activeX/activeX1.xml").is_none());
    }

    #[test]
    fn remove_vba_project_strips_vml_rid_references() {
        let vml_xml = r##"<?xml version="1.0" encoding="UTF-8"?>
<xml xmlns:v="urn:schemas-microsoft-com:vml"
 xmlns:o="urn:schemas-microsoft-com:office:office"
 xmlns:x="urn:schemas-microsoft-com:office:excel"
 xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">
  <v:shape id="_x0000_s1025" type="#_x0000_t75">
    <o:OLEObject r:id="rIdOle"/>
  </v:shape>
  <v:shape id="_x0000_s1026" type="#_x0000_t75">
    <x:ClientData ObjectType="Note"></x:ClientData>
  </v:shape>
</xml>"##;

        let rels_xml = r#"<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
  <Relationship Id="rIdOle" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/activeXControlBinary" Target="../activeX/activeX1.bin"/>
</Relationships>"#;

        let bytes = build_package(&[
            ("xl/drawings/vmlDrawing1.vml", vml_xml.as_bytes()),
            ("xl/drawings/_rels/vmlDrawing1.vml.rels", rels_xml.as_bytes()),
            ("xl/activeX/activeX1.bin", b"dummy-bin"),
        ]);

        let mut pkg = XlsxPackage::from_bytes(&bytes).expect("read pkg");
        pkg.remove_vba_project().expect("strip macros");

        let updated_rels =
            std::str::from_utf8(pkg.part("xl/drawings/_rels/vmlDrawing1.vml.rels").unwrap()).unwrap();
        assert!(!updated_rels.contains("rIdOle"));

        let updated_vml =
            std::str::from_utf8(pkg.part("xl/drawings/vmlDrawing1.vml").unwrap()).unwrap();
        assert!(!updated_vml.contains("rIdOle"));
        assert!(!updated_vml.contains("OLEObject"));
        assert!(updated_vml.contains("ObjectType=\"Note\""));

        assert!(pkg.part("xl/activeX/activeX1.bin").is_none());
    }

    #[test]
    fn remove_vba_project_strips_drawing_embed_references() {
        let drawing_xml = r#"<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<xdr:wsDr xmlns:xdr="http://schemas.openxmlformats.org/drawingml/2006/spreadsheetDrawing"
 xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main"
 xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">
  <xdr:twoCellAnchor>
    <xdr:pic>
      <xdr:blipFill>
        <a:blip r:embed="rIdImg"/>
      </xdr:blipFill>
    </xdr:pic>
  </xdr:twoCellAnchor>
  <xdr:twoCellAnchor>
    <xdr:pic>
      <xdr:blipFill>
        <a:blip r:embed="rIdKeep"/>
      </xdr:blipFill>
    </xdr:pic>
  </xdr:twoCellAnchor>
</xdr:wsDr>"#;

        let rels_xml = r#"<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
  <Relationship Id="rIdImg" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/image" Target="../ctrlProps/image1.png"/>
  <Relationship Id="rIdKeep" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/image" Target="../media/image2.png"/>
</Relationships>"#;

        let bytes = build_package(&[
            ("xl/drawings/drawing1.xml", drawing_xml.as_bytes()),
            ("xl/drawings/_rels/drawing1.xml.rels", rels_xml.as_bytes()),
            ("xl/ctrlProps/image1.png", b"macro-image"),
            ("xl/media/image2.png", b"keep-image"),
        ]);

        let mut pkg = XlsxPackage::from_bytes(&bytes).expect("read pkg");
        pkg.remove_vba_project().expect("strip macros");

        let updated_rels =
            std::str::from_utf8(pkg.part("xl/drawings/_rels/drawing1.xml.rels").unwrap()).unwrap();
        assert!(!updated_rels.contains("rIdImg"));
        assert!(updated_rels.contains("rIdKeep"));

        let updated_drawing = std::str::from_utf8(pkg.part("xl/drawings/drawing1.xml").unwrap()).unwrap();
        assert!(!updated_drawing.contains("rIdImg"));
        assert!(updated_drawing.contains("rIdKeep"));

        assert!(pkg.part("xl/ctrlProps/image1.png").is_none());
        assert!(pkg.part("xl/media/image2.png").is_some());
    }

    #[test]
    fn remove_vba_project_strips_vml_relid_references() {
        let vml_xml = r##"<?xml version="1.0" encoding="UTF-8"?>
<xml xmlns:v="urn:schemas-microsoft-com:vml"
 xmlns:o="urn:schemas-microsoft-com:office:office"
 xmlns:x="urn:schemas-microsoft-com:office:excel">
  <v:shape id="_x0000_s1025" type="#_x0000_t75">
    <v:imagedata o:relid="rIdImg"/>
  </v:shape>
  <v:shape id="_x0000_s1026" type="#_x0000_t75">
    <v:imagedata o:relid="rIdKeep"/>
  </v:shape>
</xml>"##;

        let rels_xml = r#"<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
  <Relationship Id="rIdImg" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/image" Target="../ctrlProps/image1.png"/>
  <Relationship Id="rIdKeep" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/image" Target="../media/image2.png"/>
</Relationships>"#;

        let bytes = build_package(&[
            ("xl/drawings/vmlDrawing1.vml", vml_xml.as_bytes()),
            ("xl/drawings/_rels/vmlDrawing1.vml.rels", rels_xml.as_bytes()),
            ("xl/ctrlProps/image1.png", b"macro-image"),
            ("xl/media/image2.png", b"keep-image"),
        ]);

        let mut pkg = XlsxPackage::from_bytes(&bytes).expect("read pkg");
        pkg.remove_vba_project().expect("strip macros");

        let updated_rels =
            std::str::from_utf8(pkg.part("xl/drawings/_rels/vmlDrawing1.vml.rels").unwrap()).unwrap();
        assert!(!updated_rels.contains("rIdImg"));
        assert!(updated_rels.contains("rIdKeep"));

        let updated_vml =
            std::str::from_utf8(pkg.part("xl/drawings/vmlDrawing1.vml").unwrap()).unwrap();
        assert!(!updated_vml.contains("rIdImg"));
        assert!(updated_vml.contains("rIdKeep"));

        assert!(pkg.part("xl/ctrlProps/image1.png").is_none());
        assert!(pkg.part("xl/media/image2.png").is_some());
    }
}
