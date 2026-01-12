//! Best-effort extraction for Excel "linked data types" (Stocks/Geography/etc.).
//!
//! Excel 365 stores linked data types via the RichData chain:
//! - worksheet cells (`xl/worksheets/sheet*.xml`) reference a value-metadata record via `c/@vm`
//! - `xl/metadata.xml` resolves that value-metadata index to a rich value index (`xlrd:rvb/@i`)
//! - rich value instances live in `xl/richData/richValue*.xml`
//! - `xl/richData/richValueTypes.xml` maps an instance `rv/@type` numeric ID to a type name +
//!   structure ID
//! - `xl/richData/richValueStructure.xml` defines the ordered member list for each structure
//!
//! This module implements a "best effort" extractor focused on surfacing:
//! - type name (e.g. `com.microsoft.excel.stocks`)
//! - display string (when the structure contains a `display` member)
//! - raw `<v>` scalar payloads

use std::collections::HashMap;

use formula_model::CellRef;

use crate::XlsxPackage;

use super::RichDataError;

const DEFAULT_RICH_VALUE_TYPES_PART: &str = "xl/richData/richValueTypes.xml";
const DEFAULT_RICH_VALUE_STRUCTURE_PART: &str = "xl/richData/richValueStructure.xml";

/// A linked data type instance extracted for a single cell.
#[derive(Debug, Clone, PartialEq, Eq)]
pub struct ExtractedLinkedDataType {
    /// The rich value `rv/@type` numeric ID (when present).
    pub type_id: Option<u32>,
    /// Type name from `xl/richData/richValueTypes.xml` (e.g. `com.microsoft.excel.stocks`).
    pub type_name: Option<String>,
    /// Structure ID from `xl/richData/richValueTypes.xml` (e.g. `s_stock`).
    pub structure_id: Option<String>,
    /// The `display` member (when the structure contains a `display` member and the payload has a
    /// value at that index).
    pub display: Option<String>,
    /// Raw scalar payloads from `<rv><v>...</v>*</rv>` in positional order.
    pub raw_values: Vec<String>,
}

#[derive(Debug, Clone, PartialEq, Eq)]
struct RichValueScalarRecord {
    type_id: Option<u32>,
    raw_values: Vec<String>,
}

/// Best-effort extraction of linked data types (Stocks/Geography/etc.) from a package.
///
/// Missing parts return `Ok(empty)`. Malformed XML returns an error.
pub fn extract_linked_data_types(
    pkg: &XlsxPackage,
) -> Result<HashMap<(String, CellRef), ExtractedLinkedDataType>, RichDataError> {
    // If we can't resolve sheet names to parts, we can't provide stable (sheet name, cell) keys.
    // Treat missing workbook parts as "no richData".
    if pkg.part("xl/workbook.xml").is_none() || pkg.part("xl/_rels/workbook.xml.rels").is_none() {
        return Ok(HashMap::new());
    }

    // The workbook parsing stack can error for malformed workbook.xml; bubble that up.
    let worksheet_parts = pkg.worksheet_parts()?;

    let metadata_part = super::resolve_workbook_metadata_part_name(pkg)?;
    let Some(metadata_bytes) = pkg.part(&metadata_part) else {
        return Ok(HashMap::new());
    };
    let vm_to_rich_value = super::parse_vm_to_rich_value_index_map(metadata_bytes, &metadata_part)?;
    if vm_to_rich_value.is_empty() {
        return Ok(HashMap::new());
    }

    let mut cells_with_rich_value: Vec<(String, CellRef, u32)> = Vec::new();
    for sheet in worksheet_parts {
        let Some(sheet_bytes) = pkg.part(&sheet.worksheet_part) else {
            continue;
        };
        let cells = super::parse_worksheet_vm_cells(sheet_bytes)?;
        for (cell, vm) in cells {
            let Some(&rich_value_idx) = vm_to_rich_value.get(&vm) else {
                continue;
            };
            cells_with_rich_value.push((sheet.name.clone(), cell, rich_value_idx));
        }
    }
    if cells_with_rich_value.is_empty() {
        return Ok(HashMap::new());
    }

    // `richValueTypes.xml` + `richValueStructure.xml` can be referenced either via workbook-level
    // relationships or via `xl/_rels/metadata.xml.rels`, and some producers vary the exact file
    // naming/casing. Locate them best-effort.
    let Some((_rich_value_types_part, types)) = find_rich_value_types_table(pkg)? else {
        return Ok(HashMap::new());
    };
    let types_by_id: HashMap<u32, (Option<String>, Option<String>)> = types
        .into_iter()
        .map(|t| (t.id, (t.name, t.structure_id)))
        .collect();

    let Some((_rich_value_structure_part, structures)) = find_rich_value_structure_table(pkg)? else {
        return Ok(HashMap::new());
    };
    let display_member_index_by_structure: HashMap<&str, usize> = structures
        .iter()
        .filter_map(|(id, structure)| {
            structure
                .members
                .iter()
                .position(|m| m.name.eq_ignore_ascii_case("display"))
                .map(|idx| (id.as_str(), idx))
        })
        .collect();

    let rich_values = parse_rich_value_store(pkg)?;
    if rich_values.is_empty() {
        return Ok(HashMap::new());
    }

    let mut out: HashMap<(String, CellRef), ExtractedLinkedDataType> = HashMap::new();
    for (sheet_name, cell, rich_value_idx) in cells_with_rich_value {
        let Some(record) = rich_values.get(rich_value_idx as usize) else {
            continue;
        };

        let (type_name, structure_id) = record
            .type_id
            .and_then(|id| types_by_id.get(&id))
            .cloned()
            .unwrap_or((None, None));

        let display = structure_id
            .as_deref()
            .and_then(|id| display_member_index_by_structure.get(id).copied())
            .and_then(|idx| record.raw_values.get(idx).cloned());

        out.insert(
            (sheet_name, cell),
            ExtractedLinkedDataType {
                type_id: record.type_id,
                type_name,
                structure_id,
                display,
                raw_values: record.raw_values.clone(),
            },
        );
    }

    Ok(out)
}

fn parse_rich_value_store(pkg: &XlsxPackage) -> Result<Vec<RichValueScalarRecord>, RichDataError> {
    let mut rich_value_parts: Vec<&str> = pkg
        .part_names()
        .filter(|name| super::is_rich_value_part(name))
        .collect();
    rich_value_parts.sort_by(|a, b| super::cmp_rich_value_parts_by_numeric_suffix(a, b));
    if rich_value_parts.is_empty() {
        return Ok(Vec::new());
    }

    let mut out: Vec<RichValueScalarRecord> = Vec::new();
    for part_name in rich_value_parts {
        let Some(bytes) = pkg.part(part_name) else {
            continue;
        };

        let xml = std::str::from_utf8(bytes).map_err(|source| RichDataError::XmlNonUtf8 {
            part: part_name.to_string(),
            source,
        })?;
        let doc = roxmltree::Document::parse(xml).map_err(|source| RichDataError::XmlParse {
            part: part_name.to_string(),
            source,
        })?;

        for rv in doc
            .descendants()
            .filter(|n| n.is_element() && n.tag_name().name() == "rv")
        {
            let type_id = rv
                .attribute("type")
                .or_else(|| rv.attribute("t"))
                .and_then(|v| v.trim().parse::<u32>().ok());

            // Rich values store scalar payloads in `<v>` nodes, typically matching the ordering of
            // members declared in `richValueStructure.xml`. Some producers wrap values in
            // additional container nodes, so scan descendants but avoid crossing into nested `<rv>`
            // records.
            let mut raw_values: Vec<String> = Vec::new();
            for v in rv
                .descendants()
                .filter(|n| n.is_element() && n.tag_name().name() == "v")
            {
                if v.ancestors()
                    .filter(|n| n.is_element())
                    .find(|n| n.tag_name().name() == "rv")
                    .is_some_and(|closest_rv| closest_rv != rv)
                {
                    continue;
                }

                raw_values.push(v.text().unwrap_or("").trim().to_string());
            }

            out.push(RichValueScalarRecord { type_id, raw_values });
        }
    }

    Ok(out)
}

fn find_part_case_insensitive<'a>(pkg: &'a XlsxPackage, desired: &str) -> Option<&'a str> {
    let desired = desired.strip_prefix('/').unwrap_or(desired);
    let desired_lower = desired.to_ascii_lowercase();
    pkg.part_names().find(|name| {
        let normalized = name.strip_prefix('/').unwrap_or(name);
        normalized.to_ascii_lowercase() == desired_lower
    })
}

fn find_rich_value_types_table(
    pkg: &XlsxPackage,
) -> Result<Option<(String, super::rich_value_types::RichValueTypes)>, RichDataError> {
    if let Some(part_name) = find_part_case_insensitive(pkg, DEFAULT_RICH_VALUE_TYPES_PART) {
        if let Some(bytes) = pkg.part(part_name) {
            // If the canonical part is present but malformed, surface the error. This avoids silently
            // masking a corrupted `richValueTypes.xml` in well-formed workbooks.
            let parsed = parse_rich_value_types(bytes, part_name)?;
            if !parsed.is_empty() {
                return Ok(Some((part_name.to_string(), parsed)));
            }
        }
    }

    // If metadata has a `.rels`, prefer it as a discovery mechanism for non-canonical part names.
    // This avoids hard-coding relationship Type URIs.
    if let Ok(discovered) = super::discover_rich_data_part_names(pkg) {
        for part in discovered {
            let Some(bytes) = pkg.part(&part) else {
                continue;
            };
            let parsed = match parse_rich_value_types(bytes, &part) {
                Ok(v) => v,
                Err(_) => continue,
            };
            if !parsed.is_empty() {
                return Ok(Some((part, parsed)));
            }
        }
    }

    // Last resort: scan all `xl/richData/*.xml` parts and pick the first that parses as a non-empty
    // type table.
    for part_name in pkg.part_names().filter(|name| {
        name.starts_with("xl/richData/")
            && !name.contains("/_rels/")
            && name.to_ascii_lowercase().ends_with(".xml")
    }) {
        let Some(bytes) = pkg.part(part_name) else {
            continue;
        };
        let parsed = match parse_rich_value_types(bytes, part_name) {
            Ok(v) => v,
            Err(_) => continue,
        };
        if !parsed.is_empty() {
            return Ok(Some((part_name.to_string(), parsed)));
        }
    }

    Ok(None)
}

fn find_rich_value_structure_table(
    pkg: &XlsxPackage,
) -> Result<Option<(String, super::rich_value_structure::RichValueStructures)>, RichDataError> {
    if let Some(part_name) = find_part_case_insensitive(pkg, DEFAULT_RICH_VALUE_STRUCTURE_PART) {
        if let Some(bytes) = pkg.part(part_name) {
            // If the canonical part is present but malformed, surface the error. This avoids silently
            // masking a corrupted `richValueStructure.xml` in well-formed workbooks.
            let parsed = parse_rich_value_structure(bytes, part_name)?;
            if !parsed.is_empty() {
                return Ok(Some((part_name.to_string(), parsed)));
            }
        }
    }

    if let Ok(discovered) = super::discover_rich_data_part_names(pkg) {
        for part in discovered {
            let Some(bytes) = pkg.part(&part) else {
                continue;
            };
            let parsed = match parse_rich_value_structure(bytes, &part) {
                Ok(v) => v,
                Err(_) => continue,
            };
            if !parsed.is_empty() {
                return Ok(Some((part, parsed)));
            }
        }
    }

    for part_name in pkg.part_names().filter(|name| {
        name.starts_with("xl/richData/")
            && !name.contains("/_rels/")
            && name.to_ascii_lowercase().ends_with(".xml")
    }) {
        let Some(bytes) = pkg.part(part_name) else {
            continue;
        };
        let parsed = match parse_rich_value_structure(bytes, part_name) {
            Ok(v) => v,
            Err(_) => continue,
        };
        if !parsed.is_empty() {
            return Ok(Some((part_name.to_string(), parsed)));
        }
    }

    Ok(None)
}

#[cfg(test)]
mod tests {
    use std::io::{Cursor, Write};

    use zip::write::FileOptions;

    use super::parse_rich_value_store;
    use super::{extract_linked_data_types, DEFAULT_RICH_VALUE_STRUCTURE_PART, DEFAULT_RICH_VALUE_TYPES_PART};
    use formula_model::CellRef;

    #[test]
    fn rich_value_store_collects_nested_v_elements_in_document_order() {
        let rich_value_xml = br#"<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<rvData xmlns="http://schemas.microsoft.com/office/spreadsheetml/2017/richdata">
  <values>
    <rv type="0">
      <wrapper><v> one </v></wrapper>
      <v> two </v>
    </rv>
    <rv type="1"><v> three </v></rv>
  </values>
</rvData>"#;

        // Also include a second part to ensure multi-part concatenation is honored.
        let rich_value10_xml = br#"<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<rvData xmlns="http://schemas.microsoft.com/office/spreadsheetml/2017/richdata">
  <values>
    <rv type="2"><v> four </v></rv>
  </values>
</rvData>"#;

        let cursor = Cursor::new(Vec::new());
        let mut zip = zip::ZipWriter::new(cursor);
        let options =
            FileOptions::<()>::default().compression_method(zip::CompressionMethod::Deflated);

        zip.start_file("xl/richData/richValue.xml", options).unwrap();
        zip.write_all(rich_value_xml).unwrap();

        // Use `richValue10.xml` to validate numeric sorting (10 should come after 0).
        zip.start_file("xl/richData/richValue10.xml", options).unwrap();
        zip.write_all(rich_value10_xml).unwrap();

        let bytes = zip.finish().unwrap().into_inner();
        let pkg = crate::XlsxPackage::from_bytes(&bytes).unwrap();

        let parsed = parse_rich_value_store(&pkg).unwrap();
        let got: Vec<(Option<u32>, Vec<String>)> = parsed
            .into_iter()
            .map(|rv| (rv.type_id, rv.raw_values))
            .collect();
        assert_eq!(
            got,
            vec![
                (Some(0), vec!["one".to_string(), "two".to_string()]),
                (Some(1), vec!["three".to_string()]),
                (Some(2), vec!["four".to_string()]),
            ]
        );
    }

    #[test]
    fn discovers_noncanonical_types_and_structure_part_names() {
        // Minimal workbook with one sheet and one linked data type cell.
        let workbook_xml = br#"<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<workbook xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main"
  xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">
  <sheets>
    <sheet name="Sheet1" sheetId="1" r:id="rId1"/>
  </sheets>
</workbook>"#;

        let workbook_rels = br#"<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
  <Relationship Id="rId1"
    Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/worksheet"
    Target="worksheets/sheet1.xml"/>
  <Relationship Id="rId2"
    Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/metadata"
    Target="metadata.xml"/>
</Relationships>"#;

        let sheet1_xml = br#"<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main">
  <sheetData>
    <row r="1">
      <c r="A1" t="inlineStr" vm="1"><is><t>MSFT</t></is></c>
    </row>
  </sheetData>
</worksheet>"#;

        let metadata_xml = br#"<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<metadata xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main"
  xmlns:xlrd="http://schemas.microsoft.com/office/spreadsheetml/2017/richdata">
  <metadataTypes count="1">
    <metadataType name="XLRICHVALUE"/>
  </metadataTypes>
  <futureMetadata name="XLRICHVALUE" count="1">
    <bk>
      <extLst>
        <ext uri="{00000000-0000-0000-0000-000000000000}">
          <xlrd:rvb i="0"/>
        </ext>
      </extLst>
    </bk>
  </futureMetadata>
  <valueMetadata count="1">
    <bk><rc t="1" v="0"/></bk>
  </valueMetadata>
</metadata>"#;

        // Note: these targets are *not* the canonical richValueTypes/richValueStructure names.
        let metadata_rels = br#"<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
  <Relationship Id="rId1" Type="urn:example:types" Target="richData/customTypes.xml"/>
  <Relationship Id="rId2" Type="urn:example:struct" Target="richData/customStructure.xml"/>
</Relationships>"#;

        let rich_value_xml = br#"<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<rvData xmlns="http://schemas.microsoft.com/office/spreadsheetml/2017/richdata">
  <values>
    <rv type="0"><v>MSFT</v></rv>
  </values>
</rvData>"#;

        let custom_types_xml = br#"<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<rvTypes xmlns="http://schemas.microsoft.com/office/spreadsheetml/2017/richdata">
  <types>
    <type id="0" name="com.microsoft.excel.stocks" structure="s_stock"/>
  </types>
</rvTypes>"#;

        let custom_structure_xml = br#"<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<rvStruct xmlns="http://schemas.microsoft.com/office/spreadsheetml/2017/richdata">
  <structures>
    <structure id="s_stock">
      <member name="display" kind="string"/>
    </structure>
  </structures>
</rvStruct>"#;

        let cursor = Cursor::new(Vec::new());
        let mut zip = zip::ZipWriter::new(cursor);
        let options =
            FileOptions::<()>::default().compression_method(zip::CompressionMethod::Deflated);

        zip.start_file("xl/workbook.xml", options).unwrap();
        zip.write_all(workbook_xml).unwrap();
        zip.start_file("xl/_rels/workbook.xml.rels", options).unwrap();
        zip.write_all(workbook_rels).unwrap();
        zip.start_file("xl/worksheets/sheet1.xml", options).unwrap();
        zip.write_all(sheet1_xml).unwrap();
        zip.start_file("xl/metadata.xml", options).unwrap();
        zip.write_all(metadata_xml).unwrap();
        zip.start_file("xl/_rels/metadata.xml.rels", options).unwrap();
        zip.write_all(metadata_rels).unwrap();

        // Put canonical filenames *absent* so discovery has to pick up the custom targets.
        assert_ne!(DEFAULT_RICH_VALUE_TYPES_PART, "xl/richData/customTypes.xml");
        assert_ne!(
            DEFAULT_RICH_VALUE_STRUCTURE_PART,
            "xl/richData/customStructure.xml"
        );
        zip.start_file("xl/richData/customTypes.xml", options).unwrap();
        zip.write_all(custom_types_xml).unwrap();
        zip.start_file("xl/richData/customStructure.xml", options).unwrap();
        zip.write_all(custom_structure_xml).unwrap();
        zip.start_file("xl/richData/richValue.xml", options).unwrap();
        zip.write_all(rich_value_xml).unwrap();

        let bytes = zip.finish().unwrap().into_inner();
        let pkg = crate::XlsxPackage::from_bytes(&bytes).unwrap();
        let extracted = extract_linked_data_types(&pkg).unwrap();

        let key = ("Sheet1".to_string(), CellRef::from_a1("A1").unwrap());
        let entry = extracted.get(&key).expect("missing extracted entry");
        assert_eq!(entry.type_name.as_deref(), Some("com.microsoft.excel.stocks"));
        assert_eq!(entry.display.as_deref(), Some("MSFT"));
    }
}

fn parse_rich_value_types(
    bytes: &[u8],
    part_name: &str,
) -> Result<super::rich_value_types::RichValueTypes, RichDataError> {
    // Preflight UTF-8 so we can produce a part-scoped error message. `parse_rich_value_types_xml`
    // converts utf-8 errors into `XlsxError::Utf8`, which loses the part context.
    std::str::from_utf8(bytes).map_err(|source| RichDataError::XmlNonUtf8 {
        part: part_name.to_string(),
        source,
    })?;
    super::rich_value_types::parse_rich_value_types_xml(bytes).map_err(|err| match err {
        crate::XlsxError::RoXml(source) => RichDataError::XmlParse {
            part: part_name.to_string(),
            source,
        },
        other => RichDataError::Xlsx(other),
    })
}

fn parse_rich_value_structure(
    bytes: &[u8],
    part_name: &str,
) -> Result<super::rich_value_structure::RichValueStructures, RichDataError> {
    std::str::from_utf8(bytes).map_err(|source| RichDataError::XmlNonUtf8 {
        part: part_name.to_string(),
        source,
    })?;
    super::rich_value_structure::parse_rich_value_structure_xml(bytes).map_err(|err| match err {
        crate::XlsxError::RoXml(source) => RichDataError::XmlParse {
            part: part_name.to_string(),
            source,
        },
        other => RichDataError::Xlsx(other),
    })
}
