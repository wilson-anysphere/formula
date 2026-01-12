#![allow(dead_code)]

use std::collections::{BTreeMap, BTreeSet};
use std::io::{Cursor, Write};
use std::path::Path;

/// Test helper for constructing synthetic XLSX ZIPs that include the richData / metadata parts
/// Excel uses for "rich values" (linked data types, images, etc.).
///
/// The goal is to keep richData edge-case tests readable by avoiding repeated boilerplate around
/// workbook.xml / workbook.xml.rels / ZIP creation.
///
/// This builder intentionally does *not* attempt to validate the XML payloads; tests are free to
/// provide malformed XML to exercise error handling in parsers.
#[derive(Debug, Clone)]
pub struct RichDataXlsxBuilder {
    sheets: Vec<SheetSpec>,

    // Optional richData/metadata parts.
    metadata_xml: Option<String>,
    rich_value_parts: BTreeMap<u32, String>,
    rich_value_rel_xml: Option<String>,
    rich_value_rel_rels_xml: Option<String>,

    // Binary payloads referenced from richValueRel.xml.rels (typically under `xl/media/`).
    media_parts: BTreeMap<String, Vec<u8>>,

    // Optional additional parts for tests that need them.
    extra_parts: BTreeMap<String, Vec<u8>>,

    // Optional [Content_Types].xml configuration.
    //
    // If `content_types_xml` is set we write it verbatim and ignore `content_type_overrides`.
    content_types_xml: Option<String>,
    content_type_overrides: Vec<(String, String)>,

    // Optional overrides for workbook.xml / workbook.xml.rels. When unset, we generate minimal
    // versions based on the configured sheets.
    workbook_xml: Option<String>,
    workbook_rels_xml: Option<String>,

    // Extra relationships to inject into workbook.xml.rels when we generate it.
    workbook_relationships: Vec<RelationshipSpec>,

    // Relationship type URIs used when auto-wiring metadata/richData parts.
    workbook_metadata_rel_type: String,
    workbook_rich_value_rel_type: String,
}

#[derive(Debug, Clone)]
struct SheetSpec {
    name: String,
    xml: String,
    part_name: String,
    rel_id: String,
    sheet_id: u32,
}

#[derive(Debug, Clone)]
struct RelationshipSpec {
    id: String,
    type_uri: String,
    target: String,
    target_mode: Option<String>,
}

impl Default for SheetSpec {
    fn default() -> Self {
        Self {
            name: "Sheet1".to_string(),
            xml: minimal_worksheet_xml(),
            part_name: "xl/worksheets/sheet1.xml".to_string(),
            rel_id: "rId1".to_string(),
            sheet_id: 1,
        }
    }
}

impl Default for RelationshipSpec {
    fn default() -> Self {
        Self {
            id: "rId1".to_string(),
            type_uri: REL_TYPE_WORKSHEET.to_string(),
            target: "worksheets/sheet1.xml".to_string(),
            target_mode: None,
        }
    }
}

impl Default for RichDataXlsxBuilder {
    fn default() -> Self {
        Self {
            sheets: Vec::new(),
            metadata_xml: None,
            rich_value_parts: BTreeMap::new(),
            rich_value_rel_xml: None,
            rich_value_rel_rels_xml: None,
            media_parts: BTreeMap::new(),
            extra_parts: BTreeMap::new(),
            content_types_xml: None,
            content_type_overrides: Vec::new(),
            workbook_xml: None,
            workbook_rels_xml: None,
            workbook_relationships: Vec::new(),
            workbook_metadata_rel_type: REL_TYPE_METADATA.to_string(),
            workbook_rich_value_rel_type: REL_TYPE_RICH_VALUE.to_string(),
        }
    }
}

impl RichDataXlsxBuilder {
    /// Relationship type used for workbook -> worksheet parts.
    pub const REL_TYPE_WORKSHEET: &'static str = REL_TYPE_WORKSHEET;

    /// Relationship type typically used for workbook -> `xl/metadata.xml`.
    ///
    /// This is part of the OpenXML spec and should be stable.
    pub const REL_TYPE_METADATA: &'static str = REL_TYPE_METADATA;

    /// Relationship type Excel emits for workbook -> `xl/richData/richValueRel.xml`.
    ///
    /// This is Microsoft-specific; if a test needs a different URI it can call
    /// [`Self::workbook_rich_value_rel_type`].
    pub const REL_TYPE_RICH_VALUE: &'static str = REL_TYPE_RICH_VALUE;

    pub fn new() -> Self {
        Self::default()
    }

    /// Remove any configured sheets.
    pub fn clear_sheets(mut self) -> Self {
        self.sheets.clear();
        self
    }

    /// Add a workbook sheet with a specific worksheet XML string.
    ///
    /// The worksheet part will be created at `xl/worksheets/sheet{n}.xml` and wired via
    /// `xl/workbook.xml` + `xl/_rels/workbook.xml.rels`.
    pub fn add_sheet(mut self, name: impl Into<String>, worksheet_xml: impl Into<String>) -> Self {
        let idx = self.sheets.len() + 1;
        let name = name.into();
        let sheet_id = idx as u32;
        let rel_id = format!("rId{idx}");
        let part_name = format!("xl/worksheets/sheet{idx}.xml");
        self.sheets.push(SheetSpec {
            name,
            xml: worksheet_xml.into(),
            part_name,
            rel_id,
            sheet_id,
        });
        self
    }

    /// Add a sheet with an explicit worksheet part name.
    ///
    /// This is useful for edge-cases where the worksheet isn't stored at the conventional
    /// `xl/worksheets/sheetN.xml` path.
    pub fn add_sheet_with_part_name(
        mut self,
        name: impl Into<String>,
        worksheet_part_name: impl Into<String>,
        worksheet_xml: impl Into<String>,
    ) -> Self {
        let idx = self.sheets.len() + 1;
        let name = name.into();
        let sheet_id = idx as u32;
        let rel_id = format!("rId{idx}");
        let part_name = worksheet_part_name.into();
        self.sheets.push(SheetSpec {
            name,
            xml: worksheet_xml.into(),
            part_name,
            rel_id,
            sheet_id,
        });
        self
    }

    /// Provide an explicit `xl/workbook.xml` payload.
    ///
    /// If unset, the builder generates a minimal workbook XML listing the configured sheets.
    pub fn workbook_xml(mut self, workbook_xml: impl Into<String>) -> Self {
        self.workbook_xml = Some(workbook_xml.into());
        self
    }

    /// Provide an explicit `xl/_rels/workbook.xml.rels` payload.
    ///
    /// If unset, the builder generates a minimal relationship part containing worksheet
    /// relationships (and, if provided, metadata + richValueRel relationships).
    pub fn workbook_rels_xml(mut self, workbook_rels_xml: impl Into<String>) -> Self {
        self.workbook_rels_xml = Some(workbook_rels_xml.into());
        self
    }

    /// Add an extra workbook relationship entry when generating `xl/_rels/workbook.xml.rels`.
    ///
    /// Relationship IDs must be unique within the `.rels` part. The builder does not enforce this.
    pub fn add_workbook_relationship(
        mut self,
        id: impl Into<String>,
        type_uri: impl Into<String>,
        target: impl Into<String>,
    ) -> Self {
        self.workbook_relationships.push(RelationshipSpec {
            id: id.into(),
            type_uri: type_uri.into(),
            target: target.into(),
            target_mode: None,
        });
        self
    }

    /// Override the relationship type URI used for workbook -> `xl/metadata.xml`.
    pub fn workbook_metadata_rel_type(mut self, rel_type: impl Into<String>) -> Self {
        self.workbook_metadata_rel_type = rel_type.into();
        self
    }

    /// Override the relationship type URI used for workbook -> `xl/richData/richValueRel.xml`.
    pub fn workbook_rich_value_rel_type(mut self, rel_type: impl Into<String>) -> Self {
        self.workbook_rich_value_rel_type = rel_type.into();
        self
    }

    /// Set `xl/metadata.xml`.
    pub fn metadata_xml(mut self, xml: impl Into<String>) -> Self {
        self.metadata_xml = Some(xml.into());
        self
    }

    /// Add a `xl/richData/richValue{n}.xml` part.
    pub fn rich_value_xml(mut self, index: u32, xml: impl Into<String>) -> Self {
        self.rich_value_parts.insert(index, xml.into());
        self
    }

    /// Set `xl/richData/richValueRel.xml`.
    pub fn rich_value_rel_xml(mut self, xml: impl Into<String>) -> Self {
        self.rich_value_rel_xml = Some(xml.into());
        self
    }

    /// Set `xl/richData/_rels/richValueRel.xml.rels`.
    pub fn rich_value_rel_rels_xml(mut self, xml: impl Into<String>) -> Self {
        self.rich_value_rel_rels_xml = Some(xml.into());
        self
    }

    /// Add a binary media part (e.g. `xl/media/image1.png`).
    pub fn media_part(mut self, part_name: impl Into<String>, bytes: impl Into<Vec<u8>>) -> Self {
        self.media_parts.insert(part_name.into(), bytes.into());
        self
    }

    /// Add an arbitrary extra part to the ZIP.
    pub fn part(mut self, part_name: impl Into<String>, bytes: impl Into<Vec<u8>>) -> Self {
        self.extra_parts.insert(part_name.into(), bytes.into());
        self
    }

    /// Provide an explicit `[Content_Types].xml` payload.
    ///
    /// If unset, the builder generates a minimal content-types file for workbook + worksheets.
    pub fn content_types_xml(mut self, xml: impl Into<String>) -> Self {
        self.content_types_xml = Some(xml.into());
        self
    }

    /// Add an `<Override>` entry to `[Content_Types].xml` when generating it.
    ///
    /// `part_name` may be passed with or without a leading `/`.
    pub fn content_type_override(
        mut self,
        part_name: impl Into<String>,
        content_type: impl Into<String>,
    ) -> Self {
        self.content_type_overrides
            .push((part_name.into(), content_type.into()));
        self
    }

    /// Build the final XLSX bytes.
    pub fn build_bytes(self) -> Vec<u8> {
        let parts = self.build_parts_map();
        write_zip(parts)
    }

    fn build_parts_map(self) -> BTreeMap<String, Vec<u8>> {
        let mut parts: BTreeMap<String, Vec<u8>> = BTreeMap::new();

        parts.insert(
            "[Content_Types].xml".to_string(),
            self.build_content_types_xml().into_bytes(),
        );

        parts.insert("_rels/.rels".to_string(), root_rels_xml().into_bytes());

        let workbook_xml = self
            .workbook_xml
            .unwrap_or_else(|| workbook_xml_for_sheets(&self.sheets));
        parts.insert("xl/workbook.xml".to_string(), workbook_xml.into_bytes());

        let workbook_rels_xml = self.workbook_rels_xml.unwrap_or_else(|| {
            workbook_rels_xml_for_builder(
                &self.sheets,
                self.metadata_xml.is_some(),
                self.rich_value_rel_xml.is_some(),
                &self.workbook_relationships,
                &self.workbook_metadata_rel_type,
                &self.workbook_rich_value_rel_type,
            )
        });
        parts.insert(
            "xl/_rels/workbook.xml.rels".to_string(),
            workbook_rels_xml.into_bytes(),
        );

        for sheet in &self.sheets {
            parts.insert(sheet.part_name.clone(), sheet.xml.as_bytes().to_vec());
        }

        if let Some(xml) = self.metadata_xml {
            parts.insert("xl/metadata.xml".to_string(), xml.into_bytes());
        }

        for (idx, xml) in self.rich_value_parts {
            parts.insert(
                format!("xl/richData/richValue{idx}.xml"),
                xml.into_bytes(),
            );
        }

        if let Some(xml) = self.rich_value_rel_xml {
            parts.insert("xl/richData/richValueRel.xml".to_string(), xml.into_bytes());
        }

        if let Some(xml) = self.rich_value_rel_rels_xml {
            parts.insert(
                "xl/richData/_rels/richValueRel.xml.rels".to_string(),
                xml.into_bytes(),
            );
        }

        parts.extend(self.media_parts);
        parts.extend(self.extra_parts);
        parts
    }

    fn build_content_types_xml(&self) -> String {
        if let Some(xml) = &self.content_types_xml {
            return xml.clone();
        }
        content_types_xml_for_builder(&self.sheets, &self.media_parts, &self.content_type_overrides)
    }
}

const REL_TYPE_WORKSHEET: &str =
    "http://schemas.openxmlformats.org/officeDocument/2006/relationships/worksheet";
const REL_TYPE_METADATA: &str =
    "http://schemas.openxmlformats.org/officeDocument/2006/relationships/metadata";

// Excel rich data types (linked data types / images) use a Microsoft-specific relationship.
// This URI is widely observed in the wild; tests can override if needed.
const REL_TYPE_RICH_VALUE: &str = "http://schemas.microsoft.com/office/2017/10/relationships/richValue";

fn minimal_worksheet_xml() -> String {
    r#"<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main">
  <dimension ref="A1"/>
  <sheetData/>
</worksheet>"#
        .to_string()
}

fn root_rels_xml() -> String {
    r#"<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
  <Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="xl/workbook.xml"/>
</Relationships>"#
        .to_string()
}

fn workbook_xml_for_sheets(sheets: &[SheetSpec]) -> String {
    let mut out = String::new();
    out.push_str(r#"<?xml version="1.0" encoding="UTF-8" standalone="yes"?>"#);
    out.push('\n');
    out.push_str(
        r#"<workbook xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main""#,
    );
    out.push('\n');
    out.push_str(r#" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">"#);
    out.push('\n');
    out.push_str("  <sheets>\n");
    for sheet in sheets {
        out.push_str(&format!(
            r#"    <sheet name="{}" sheetId="{}" r:id="{}"/>"#,
            xml_escape(&sheet.name),
            sheet.sheet_id,
            xml_escape(&sheet.rel_id),
        ));
        out.push('\n');
    }
    out.push_str("  </sheets>\n");
    out.push_str("</workbook>");
    out
}

fn workbook_rels_xml_for_builder(
    sheets: &[SheetSpec],
    include_metadata: bool,
    include_rich_value_rel: bool,
    extra_relationships: &[RelationshipSpec],
    metadata_rel_type: &str,
    rich_value_rel_type: &str,
) -> String {
    let mut rels: Vec<RelationshipSpec> = Vec::new();

    for sheet in sheets {
        rels.push(RelationshipSpec {
            id: sheet.rel_id.clone(),
            type_uri: REL_TYPE_WORKSHEET.to_string(),
            target: target_relative_to_workbook(&sheet.part_name),
            target_mode: None,
        });
    }

    // Auto-wire the common richData parts so tests don't have to repeat this boilerplate.
    //
    // We intentionally only add these relationships when the corresponding parts are present.
    if include_metadata {
        rels.push(RelationshipSpec {
            id: next_r_id(&rels),
            type_uri: metadata_rel_type.to_string(),
            target: "metadata.xml".to_string(),
            target_mode: None,
        });
    }

    if include_rich_value_rel {
        rels.push(RelationshipSpec {
            id: next_r_id(&rels),
            type_uri: rich_value_rel_type.to_string(),
            target: "richData/richValueRel.xml".to_string(),
            target_mode: None,
        });
    }

    rels.extend_from_slice(extra_relationships);

    relationships_xml(&rels)
}

fn relationships_xml(rels: &[RelationshipSpec]) -> String {
    let mut out = String::new();
    out.push_str(r#"<?xml version="1.0" encoding="UTF-8" standalone="yes"?>"#);
    out.push('\n');
    out.push_str(r#"<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">"#);
    out.push('\n');
    for rel in rels {
        out.push_str("  ");
        out.push_str(&format!(
            r#"<Relationship Id="{}" Type="{}" Target="{}""#,
            xml_escape(&rel.id),
            xml_escape(&rel.type_uri),
            xml_escape(&rel.target)
        ));
        if let Some(mode) = &rel.target_mode {
            out.push_str(&format!(r#" TargetMode="{}""#, xml_escape(mode)));
        }
        out.push_str("/>\n");
    }
    out.push_str("</Relationships>");
    out
}

fn next_r_id(existing: &[RelationshipSpec]) -> String {
    let mut max = 0u32;
    for rel in existing {
        if let Some(num) = rel.id.strip_prefix("rId").and_then(|s| s.parse::<u32>().ok()) {
            max = max.max(num);
        }
    }
    format!("rId{}", max + 1)
}

fn target_relative_to_workbook(part_name: &str) -> String {
    // The standard case: `xl/worksheets/sheet1.xml` -> `worksheets/sheet1.xml`.
    if let Some(without_xl) = part_name.strip_prefix("xl/") {
        return without_xl.to_string();
    }
    part_name.to_string()
}

fn content_types_xml_for_builder(
    sheets: &[SheetSpec],
    media_parts: &BTreeMap<String, Vec<u8>>,
    overrides: &[(String, String)],
) -> String {
    let mut default_exts: BTreeMap<String, String> = BTreeMap::new();
    default_exts.insert(
        "rels".to_string(),
        "application/vnd.openxmlformats-package.relationships+xml".to_string(),
    );
    default_exts.insert("xml".to_string(), "application/xml".to_string());

    // Add defaults for common media extensions we see in XLSX (png/jpeg/gif).
    let mut seen_media_exts: BTreeSet<String> = BTreeSet::new();
    for name in media_parts.keys() {
        if let Some(ext) = Path::new(name)
            .extension()
            .and_then(|s| s.to_str())
            .map(|s| s.to_ascii_lowercase())
        {
            seen_media_exts.insert(ext);
        }
    }
    for ext in seen_media_exts {
        match ext.as_str() {
            "png" => {
                default_exts
                    .entry("png".to_string())
                    .or_insert_with(|| "image/png".to_string());
            }
            "jpg" | "jpeg" => {
                default_exts
                    .entry(ext.clone())
                    .or_insert_with(|| "image/jpeg".to_string());
            }
            "gif" => {
                default_exts
                    .entry("gif".to_string())
                    .or_insert_with(|| "image/gif".to_string());
            }
            _ => {}
        }
    }

    let mut out = String::new();
    out.push_str(r#"<?xml version="1.0" encoding="UTF-8" standalone="yes"?>"#);
    out.push('\n');
    out.push_str(r#"<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">"#);
    out.push('\n');

    for (ext, ty) in default_exts {
        out.push_str(&format!(
            r#"  <Default Extension="{}" ContentType="{}"/>"#,
            xml_escape(&ext),
            xml_escape(&ty)
        ));
        out.push('\n');
    }

    // Minimal overrides: workbook + worksheets.
    out.push_str(r#"  <Override PartName="/xl/workbook.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet.main+xml"/>"#);
    out.push('\n');
    for sheet in sheets {
        out.push_str(&format!(
            r#"  <Override PartName="/{}" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.worksheet+xml"/>"#,
            xml_escape(&sheet.part_name)
        ));
        out.push('\n');
    }

    for (part_name, content_type) in overrides {
        let part_name = if part_name.starts_with('/') {
            part_name.clone()
        } else {
            format!("/{part_name}")
        };
        out.push_str(&format!(
            r#"  <Override PartName="{}" ContentType="{}"/>"#,
            xml_escape(&part_name),
            xml_escape(content_type),
        ));
        out.push('\n');
    }

    out.push_str("</Types>");
    out
}

fn write_zip(parts: BTreeMap<String, Vec<u8>>) -> Vec<u8> {
    let cursor = Cursor::new(Vec::new());
    let mut zip = zip::ZipWriter::new(cursor);
    let options = zip::write::FileOptions::<()>::default()
        .compression_method(zip::CompressionMethod::Deflated);

    for (name, bytes) in parts {
        zip.start_file(name, options).unwrap();
        zip.write_all(&bytes).unwrap();
    }

    zip.finish().unwrap().into_inner()
}

fn xml_escape(input: &str) -> String {
    input
        .replace('&', "&amp;")
        .replace('<', "&lt;")
        .replace('>', "&gt;")
        .replace('"', "&quot;")
        .replace('\'', "&apos;")
}
