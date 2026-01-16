use std::collections::{BTreeMap, HashMap};

use formula_model::{CellRef, CellValue, EntityValue, ErrorValue, RecordValue, Workbook, WorksheetId};

use crate::rich_data::rich_value_structure::parse_rich_value_structure_xml;
use crate::rich_data::rich_value_types::parse_rich_value_types_xml;

/// Best-effort application of Excel rich values (linked data types / records) to the workbook
/// model.
///
/// Excel stores rich values via a combination of:
/// - worksheet `c/@vm` attributes
/// - `xl/metadata.xml` value metadata
/// - `xl/richData/*` parts describing the value payload
///
/// `load_from_bytes` already resolves `c/@vm` → rich-value index and stores the mapping in
/// `XlsxMeta.rich_value_cells`. This helper uses that mapping plus the richData tables to populate
/// cells as [`CellValue::Entity`] / [`CellValue::Record`].
///
/// This function is intentionally best-effort:
/// - If required parts are missing or malformed, it leaves the workbook untouched.
/// - Rich value kinds we don't recognize (e.g. images-in-cell) are ignored.
pub(crate) fn apply_rich_values_to_workbook(
    workbook: &mut Workbook,
    rich_value_cells: &HashMap<(WorksheetId, CellRef), u32>,
    parts: &BTreeMap<String, Vec<u8>>,
) {
    if rich_value_cells.is_empty() {
        return;
    }

    // We currently only attempt to decode entity/record rich values when richValueTypes +
    // richValueStructure exist. These parts are absent in many richData uses (like images-in-cell),
    // and this should not prevent the workbook from loading.
    let Some(types_bytes) = parts.get("xl/richData/richValueTypes.xml") else {
        return;
    };
    let Some(struct_bytes) = parts.get("xl/richData/richValueStructure.xml") else {
        return;
    };

    let types = match parse_rich_value_types_xml(types_bytes) {
        Ok(v) => v,
        Err(_) => return,
    };
    let structures = match parse_rich_value_structure_xml(struct_bytes) {
        Ok(v) => v,
        Err(_) => return,
    };

    let mut types_by_id = HashMap::with_capacity(types.len());
    for ty in types {
        types_by_id.insert(ty.id, ty);
    }

    let rich_values = match parse_all_rich_value_parts(parts) {
        Ok(v) => v,
        Err(_) => return,
    };
    if rich_values.is_empty() {
        return;
    }

    // Update workbook cells in-place.
    for (&(worksheet_id, cell_ref), &rich_value_index) in rich_value_cells {
        let Some(sheet) = workbook.sheet_mut(worksheet_id) else {
            continue;
        };
        let Some(existing_cell) = sheet.cell(cell_ref).cloned() else {
            continue;
        };

        let Some(rv) = rich_values.get(rich_value_index as usize) else {
            continue;
        };
        let Some(ty) = types_by_id.get(&rv.type_id) else {
            continue;
        };

        let Some(type_name) = ty.name.as_deref() else {
            continue;
        };

        // Skip non-entity rich values (e.g. images-in-cell) for now.
        // Images typically use a type name like `com.microsoft.excel.image`.
        if crate::ascii::contains_ignore_case(type_name, "image") {
            continue;
        }

        let Some(structure_id) = ty.structure_id.as_deref() else {
            continue;
        };
        let Some(structure) = structures.get(structure_id) else {
            continue;
        };

        let display_value = cell_value_display_string(&existing_cell.value);

        let mut fields: BTreeMap<String, CellValue> = BTreeMap::new();
        for (idx, member) in structure.members.iter().enumerate() {
            let atom = rv.values.get(idx);
            let raw_text = atom.map(|a| a.text.as_str()).unwrap_or_default();
            let kind = atom
                .and_then(|a| a.kind.as_deref())
                .or(member.kind.as_deref());
            let value = rich_value_atom_to_cell_value(kind, raw_text);
            fields.insert(member.name.clone(), value);
        }

        let new_value = if crate::ascii::contains_ignore_case(type_name, "record") {
            let display_field = fields.contains_key("display").then(|| "display".to_string());
            CellValue::Record(RecordValue {
                fields,
                display_field,
                display_value: display_value.clone(),
            })
        } else {
            // Default: treat as an entity.
            let mut entity = EntityValue {
                entity_type: type_name.to_string(),
                display_value: display_value.clone(),
                ..Default::default()
            };

            // Best-effort: attempt to promote an `id` field into `entity_id`.
            if let Some(CellValue::String(id)) = fields.get("id") {
                entity.entity_id = id.clone();
            }

            // Properties: keep all non-display fields.
            fields.remove("display");
            fields.remove("id");
            entity.properties = fields;

            CellValue::Entity(entity)
        };

        if new_value != existing_cell.value {
            let mut updated = existing_cell;
            updated.value = new_value;
            sheet.set_cell(cell_ref, updated);
        }
    }
}

#[derive(Debug, Clone)]
struct RichValueRecord {
    type_id: u32,
    values: Vec<RichValueAtom>,
}

#[derive(Debug, Clone)]
struct RichValueAtom {
    kind: Option<String>,
    text: String,
}

fn parse_all_rich_value_parts(
    parts: &BTreeMap<String, Vec<u8>>,
) -> Result<Vec<RichValueRecord>, crate::XlsxError> {
    let mut part_names: Vec<&str> = parts
        .keys()
        .map(|s| s.as_str())
        .filter(|p| rich_value_part_sort_key(p).is_some())
        .collect();
    part_names.sort_by_key(|p| rich_value_part_sort_key(p).unwrap_or((u8::MAX, u32::MAX, *p)));

    let mut out = Vec::new();
    for name in part_names {
        let Some(bytes) = parts.get(name) else {
            continue;
        };
        let mut records = parse_rich_value_part(bytes, name)?;
        out.append(&mut records);
    }

    Ok(out)
}

fn rich_value_part_sort_key(part_name: &str) -> Option<(u8, u32, &str)> {
    if !part_name.starts_with("xl/richData/") {
        return None;
    }

    let file_name = part_name.rsplit('/').next()?;
    let stem = crate::ascii::strip_suffix_ignore_case(file_name, ".xml")?;

    // Check the plural prefix first: `richvalues` starts with `richvalue`.
    let (family, suffix) = if let Some(rest) = crate::ascii::strip_prefix_ignore_case(stem, "richvalues") {
        (0u8, rest)
    } else if let Some(rest) = crate::ascii::strip_prefix_ignore_case(stem, "richvalue") {
        (0u8, rest)
    } else if let Some(rest) = crate::ascii::strip_prefix_ignore_case(stem, "rdrichvalue") {
        (1u8, rest)
    } else {
        return None;
    };

    let idx = if suffix.is_empty() {
        0
    } else if suffix.as_bytes().iter().all(u8::is_ascii_digit) {
        suffix.parse::<u32>().ok()?
    } else {
        return None;
    };

    Some((family, idx, part_name))
}

fn parse_rich_value_part(
    xml_bytes: &[u8],
    part_name: &str,
) -> Result<Vec<RichValueRecord>, crate::XlsxError> {
    let xml = std::str::from_utf8(xml_bytes).map_err(|e| {
        crate::XlsxError::Invalid(format!("{part_name} is not valid UTF-8: {e}"))
    })?;
    let doc = roxmltree::Document::parse(xml)?;

    let mut out = Vec::new();
    for rv in doc
        .descendants()
        .filter(|n| n.is_element() && n.tag_name().name().eq_ignore_ascii_case("rv"))
    {
        let type_attr = rv.attribute("type").or_else(|| rv.attribute("t"));
        let Some(type_attr) = type_attr else {
            continue;
        };
        let Ok(type_id) = type_attr.trim().parse::<u32>() else {
            continue;
        };

        let mut values = Vec::new();
        for v in rv
            .children()
            .filter(|n| n.is_element() && n.tag_name().name().eq_ignore_ascii_case("v"))
        {
            let kind = v
                .attribute("kind")
                .or_else(|| v.attribute("k"))
                .or_else(|| v.attribute("t"))
                .or_else(|| v.attribute("type"))
                .map(|s| s.to_string());
            let text = v.text().unwrap_or_default().to_string();
            values.push(RichValueAtom { kind, text });
        }

        out.push(RichValueRecord { type_id, values });
    }

    Ok(out)
}

fn rich_value_atom_to_cell_value(kind: Option<&str>, text: &str) -> CellValue {
    let kind = kind.unwrap_or("string");
    if kind.eq_ignore_ascii_case("string") {
        return CellValue::String(text.to_string());
    }
    if kind.eq_ignore_ascii_case("number") || kind.eq_ignore_ascii_case("n") {
        return text
            .trim()
            .parse::<f64>()
            .map(CellValue::Number)
            .unwrap_or_else(|_| CellValue::String(text.to_string()));
    }
    if kind.eq_ignore_ascii_case("bool")
        || kind.eq_ignore_ascii_case("boolean")
        || kind.eq_ignore_ascii_case("b")
    {
        let trimmed = text.trim();
        let b = trimmed == "1" || trimmed.eq_ignore_ascii_case("true");
        return CellValue::Boolean(b);
    }
    if kind.eq_ignore_ascii_case("error") || kind.eq_ignore_ascii_case("e") {
        let err = text.trim().parse::<ErrorValue>().unwrap_or(ErrorValue::Unknown);
        return CellValue::Error(err);
    }
    // Relationship indices (used by images) are stored as integers.
    if kind.eq_ignore_ascii_case("rel") || kind.eq_ignore_ascii_case("r") {
        return text
            .trim()
            .parse::<f64>()
            .map(CellValue::Number)
            .unwrap_or_else(|_| CellValue::String(text.to_string()));
    }
    CellValue::String(text.to_string())
}

fn cell_value_display_string(value: &CellValue) -> String {
    match value {
        CellValue::Empty => String::new(),
        CellValue::Number(n) => n.to_string(),
        CellValue::String(s) => s.clone(),
        CellValue::Boolean(b) => {
            if *b {
                "TRUE".to_string()
            } else {
                "FALSE".to_string()
            }
        }
        CellValue::Error(e) => e.as_str().to_string(),
        CellValue::RichText(rt) => rt.text.clone(),
        CellValue::Entity(e) => e.display_value.clone(),
        CellValue::Record(r) => r.to_string(),
        _ => String::new(),
    }
}
