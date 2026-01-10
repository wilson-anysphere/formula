use std::collections::BTreeMap;
use std::io::{Cursor, Read, Write};

use thiserror::Error;

use crate::pivots::XlsxPivots;

#[derive(Debug, Error)]
pub enum XlsxError {
    #[error("zip error: {0}")]
    Zip(#[from] zip::result::ZipError),
    #[error("io error: {0}")]
    Io(#[from] std::io::Error),
    #[error("xml error: {0}")]
    Xml(#[from] quick_xml::Error),
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
}
