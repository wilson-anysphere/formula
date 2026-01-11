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
use crate::RecalcPolicy;
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

    /// Apply a set of cell edits to the package, rewriting only the targeted worksheet XML parts.
    ///
    /// All non-targeted parts are copied byte-for-byte from the original package.
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

        let cursor = Cursor::new(Vec::new());
        let mut zip = zip::ZipWriter::new(cursor);
        let options = zip::write::FileOptions::<()>::default()
            .compression_method(zip::CompressionMethod::Deflated);

        for (name, bytes) in &self.parts {
            zip.start_file(name, options)?;
            if let Some(sheet_patches) = patches_by_part.get(name) {
                crate::streaming::patch_worksheet_xml_streaming(
                    Cursor::new(bytes.as_slice()),
                    &mut zip,
                    name,
                    sheet_patches,
                )?;
            } else {
                zip.write_all(bytes)?;
            }
        }

        let cursor = zip.finish()?;
        Ok(cursor.into_inner())
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
        self.parts.remove("xl/vbaProject.bin");
        self.parts.remove("xl/_rels/vbaProject.bin.rels");

        // Drop the VBA relationship from workbook.xml.rels (if present).
        if let Some(rels_bytes) = self.parts.get("xl/_rels/workbook.xml.rels").cloned() {
            let updated = remove_relationship_type(
                &rels_bytes,
                "http://schemas.microsoft.com/office/2006/relationships/vbaProject",
            )?;
            self.set_part("xl/_rels/workbook.xml.rels", updated);
        }

        // Remove the content type override for vbaProject.bin and convert the workbook content
        // type back to a standard `.xlsx`.
        if let Some(ct_bytes) = self.parts.get("[Content_Types].xml").cloned() {
            let updated = remove_vba_content_types(&ct_bytes)?;
            self.set_part("[Content_Types].xml", updated);
        }

        Ok(())
    }
}

fn remove_relationship_type(xml: &[u8], rel_type: &str) -> Result<Vec<u8>, XlsxError> {
    let mut reader = XmlReader::from_reader(xml);
    reader.config_mut().trim_text(false);
    let mut writer = XmlWriter::new(Vec::with_capacity(xml.len()));
    let mut buf = Vec::new();

    loop {
        match reader.read_event_into(&mut buf)? {
            Event::Eof => break,
            Event::Empty(e) if e.name().as_ref() == b"Relationship" => {
                let mut type_match = false;
                for attr in e.attributes().with_checks(false) {
                    let attr = attr?;
                    if attr.key.as_ref() == b"Type" && attr.unescape_value()?.as_ref() == rel_type {
                        type_match = true;
                        break;
                    }
                }
                if !type_match {
                    writer.write_event(Event::Empty(e.to_owned()))?;
                }
            }
            ev => writer.write_event(ev.into_owned())?,
        }
        buf.clear();
    }

    Ok(writer.into_inner())
}

fn remove_vba_content_types(xml: &[u8]) -> Result<Vec<u8>, XlsxError> {
    let mut reader = XmlReader::from_reader(xml);
    reader.config_mut().trim_text(false);
    let mut writer = XmlWriter::new(Vec::with_capacity(xml.len()));
    let mut buf = Vec::new();

    loop {
        match reader.read_event_into(&mut buf)? {
            Event::Eof => break,
            Event::Empty(e) if e.name().as_ref() == b"Override" => {
                let mut part_name: Option<String> = None;
                let mut content_type: Option<String> = None;
                for attr in e.attributes().with_checks(false) {
                    let attr = attr?;
                    match attr.key.as_ref() {
                        b"PartName" => part_name = Some(attr.unescape_value()?.into_owned()),
                        b"ContentType" => content_type = Some(attr.unescape_value()?.into_owned()),
                        _ => {}
                    }
                }

                if part_name.as_deref() == Some("/xl/vbaProject.bin") {
                    continue;
                }

                // Convert macro-enabled workbook content type back to `.xlsx`.
                if part_name.as_deref() == Some("/xl/workbook.xml")
                    && content_type
                        .as_deref()
                        .is_some_and(|ct| ct.contains("macroEnabled.main+xml"))
                {
                    let mut updated = BytesStart::new("Override");
                    updated.push_attribute(("PartName", "/xl/workbook.xml"));
                    updated.push_attribute((
                        "ContentType",
                        "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet.main+xml",
                    ));
                    writer.write_event(Event::Empty(updated))?;
                    continue;
                }

                writer.write_event(Event::Empty(e.to_owned()))?;
            }
            ev => writer.write_event(ev.into_owned())?,
        }
        buf.clear();
    }

    Ok(writer.into_inner())
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
}
