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

    // The linked data type lookup tables are fixed part names in `xl/richData/*`. Treat missing
    // parts as "no linked data types".
    let rich_value_types_part = "xl/richData/richValueTypes.xml";
    let Some(rich_value_types_bytes) = pkg.part(rich_value_types_part) else {
        return Ok(HashMap::new());
    };
    let types = parse_rich_value_types(rich_value_types_bytes, rich_value_types_part)?;
    let types_by_id: HashMap<u32, (Option<String>, Option<String>)> = types
        .into_iter()
        .map(|t| (t.id, (t.name, t.structure_id)))
        .collect();

    let rich_value_structure_part = "xl/richData/richValueStructure.xml";
    let Some(rich_value_structure_bytes) = pkg.part(rich_value_structure_part) else {
        return Ok(HashMap::new());
    };
    let structures =
        parse_rich_value_structure(rich_value_structure_bytes, rich_value_structure_part)?;
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

            let raw_values = rv
                .children()
                .filter(|n| n.is_element() && n.tag_name().name() == "v")
                .map(|v| v.text().unwrap_or("").trim().to_string())
                .collect();

            out.push(RichValueScalarRecord { type_id, raw_values });
        }
    }

    Ok(out)
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
