use std::collections::BTreeMap;
use std::io::{Cursor, Read, Write};

use quick_xml::events::{BytesStart, Event};
use quick_xml::{Reader as XmlReader, Writer as XmlWriter};
use thiserror::Error;

use crate::patch::{apply_cell_patches_to_package, WorkbookCellPatches};
use crate::pivots::cache_records::{PivotCacheRecordsReader, PivotCacheValue};
use crate::pivots::XlsxPivots;
use crate::recalc_policy::RecalcPolicyError;
use crate::sheet_metadata::{
    parse_sheet_tab_color, parse_workbook_sheets, write_sheet_tab_color, write_workbook_sheets,
    WorkbookSheetInfo,
};
use crate::RecalcPolicy;
use formula_model::TabColor;

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
