use std::collections::HashMap;
use std::io::Cursor;

use formula_model::CellRef;
use quick_xml::events::Event;
use quick_xml::Reader;

use crate::openxml;
use crate::package::{XlsxError, XlsxPackage};
use crate::path;
use crate::rich_data::{scan_cells_with_metadata_indices, RichDataError};
use crate::rich_data::rich_value::parse_rich_value_relationship_indices;

const REL_TYPE_SHEET_METADATA: &str =
    "http://schemas.openxmlformats.org/officeDocument/2006/relationships/sheetMetadata";
const REL_TYPE_METADATA: &str =
    "http://schemas.openxmlformats.org/officeDocument/2006/relationships/metadata";
const REL_TYPE_RD_RICH_VALUE: &str =
    "http://schemas.microsoft.com/office/2017/06/relationships/rdRichValue";
const REL_TYPE_RD_RICH_VALUE_STRUCTURE: &str =
    "http://schemas.microsoft.com/office/2017/06/relationships/rdRichValueStructure";
const REL_TYPE_RICH_VALUE: &str = "http://schemas.microsoft.com/office/2017/06/relationships/richValue";
const REL_TYPE_RICH_VALUE_REL: &str =
    "http://schemas.microsoft.com/office/2022/10/relationships/richValueRel";
const REL_TYPE_RICH_VALUE_REL_2017: &str =
    "http://schemas.microsoft.com/office/2017/06/relationships/richValueRel";
const REL_TYPE_IMAGE: &str =
    "http://schemas.openxmlformats.org/officeDocument/2006/relationships/image";

/// Embedded ("Place in Cell") image mapping extracted from an XLSX package.
#[derive(Debug, Clone, PartialEq, Eq)]
pub struct EmbeddedImageCell {
    /// ZIP entry name for the worksheet XML (e.g. `xl/worksheets/sheet1.xml`).
    pub sheet_part: String,
    pub cell: CellRef,
    /// Normalized OPC part path for the image payload (e.g. `xl/media/image1.png`).
    pub image_target: String,
    pub bytes: Vec<u8>,
    pub alt_text: Option<String>,
    /// Whether the image is marked as decorative.
    ///
    /// Derived from the `CalcOrigin` field in rich value data (`5` for decorative, `6` for
    /// informative images with alt text).
    pub decorative: bool,
}

#[derive(Debug, Clone)]
struct RichValueImage {
    local_image_identifier: u32,
    calc_origin: u32,
    alt_text: Option<String>,
}

#[derive(Debug, Clone, Default)]
struct LocalImageStructurePositions {
    local_image_identifier: Option<usize>,
    calc_origin: Option<usize>,
    text: Option<usize>,
}

/// Extract embedded images-in-cells ("Place in Cell") using the `vm` + `metadata.xml` + `xl/richData/*`
/// schema.
///
/// The extractor is intentionally resilient: missing parts or broken mappings result in skipped cells.
/// Only invalid ZIP/XML/UTF-8 errors are returned as [`XlsxError`].
pub fn extract_embedded_images(pkg: &XlsxPackage) -> Result<Vec<EmbeddedImageCell>, XlsxError> {
    // Discover the relevant rich data parts via `xl/_rels/workbook.xml.rels`.
    let workbook_rels_xml = match pkg.part("xl/_rels/workbook.xml.rels") {
        Some(bytes) => bytes,
        None => return Ok(Vec::new()),
    };

    let relationships = openxml::parse_relationships(workbook_rels_xml)?;
    let mut metadata_part: Option<String> = None;
    let mut rich_value_part: Option<String> = None;
    let mut rdrichvalue_part: Option<String> = None;
    let mut rdrichvaluestructure_part: Option<String> = None;
    let mut rich_value_rel_part: Option<String> = None;

    for rel in relationships {
        if metadata_part.is_none()
            && (rel.type_uri.eq_ignore_ascii_case(REL_TYPE_SHEET_METADATA)
                || rel.type_uri.eq_ignore_ascii_case(REL_TYPE_METADATA))
        {
            let target = path::resolve_target("xl/workbook.xml", &rel.target);
            metadata_part = Some(target);
        } else if rich_value_part.is_none()
            && rel.type_uri.eq_ignore_ascii_case(REL_TYPE_RICH_VALUE)
        {
            let target = path::resolve_target("xl/workbook.xml", &rel.target);
            rich_value_part = Some(target);
        } else if rdrichvalue_part.is_none() && rel.type_uri.eq_ignore_ascii_case(REL_TYPE_RD_RICH_VALUE)
        {
            let target = path::resolve_target("xl/workbook.xml", &rel.target);
            rdrichvalue_part = Some(target);
        } else if rdrichvaluestructure_part.is_none()
            && rel
                .type_uri
                .eq_ignore_ascii_case(REL_TYPE_RD_RICH_VALUE_STRUCTURE)
        {
            let target = path::resolve_target("xl/workbook.xml", &rel.target);
            rdrichvaluestructure_part = Some(target);
        } else if rich_value_rel_part.is_none()
            && (rel.type_uri.eq_ignore_ascii_case(REL_TYPE_RICH_VALUE_REL)
                || rel.type_uri.eq_ignore_ascii_case(REL_TYPE_RICH_VALUE_REL_2017))
        {
            let target = path::resolve_target("xl/workbook.xml", &rel.target);
            rich_value_rel_part = Some(target);
        }
    }

    // Fallback to canonical part names when the workbook relationship graph doesn't include the
    // expected rich-data relationships (best-effort).
    if metadata_part.is_none() && pkg.part("xl/metadata.xml").is_some() {
        metadata_part = Some("xl/metadata.xml".to_string());
    }
    if rich_value_part.is_none() && pkg.part("xl/richData/richValue.xml").is_some() {
        rich_value_part = Some("xl/richData/richValue.xml".to_string());
    }
    if rdrichvalue_part.is_none() && pkg.part("xl/richData/rdrichvalue.xml").is_some() {
        rdrichvalue_part = Some("xl/richData/rdrichvalue.xml".to_string());
    }
    if rdrichvaluestructure_part.is_none() && pkg.part("xl/richData/rdrichvaluestructure.xml").is_some()
    {
        rdrichvaluestructure_part = Some("xl/richData/rdrichvaluestructure.xml".to_string());
    }
    if rich_value_rel_part.is_none() && pkg.part("xl/richData/richValueRel.xml").is_some() {
        rich_value_rel_part = Some("xl/richData/richValueRel.xml".to_string());
    }

    let vm_to_rich_value_index = match metadata_part
        .as_deref()
        .and_then(|part| pkg.part(part).map(|bytes| (part, bytes)))
    {
        Some((_part, bytes)) => parse_vm_to_rich_value_index(bytes)?,
        None => HashMap::new(),
    };

    let local_image_structure_positions = match rdrichvaluestructure_part
        .as_deref()
        .and_then(|part| pkg.part(part).map(|bytes| (part, bytes)))
    {
        Some((_part, bytes)) => Some(parse_local_image_structure_positions(bytes)?),
        None => None,
    };

    let rich_values = match rdrichvalue_part
        .as_deref()
        .and_then(|part| pkg.part(part).map(|bytes| (part, bytes)))
    {
        Some((_part, bytes)) => parse_rdrichvalue(bytes, local_image_structure_positions.as_deref())?,
        None => HashMap::new(),
    };

    let rich_value_rel_indices = match rich_value_part
        .as_deref()
        .and_then(|part| pkg.part(part).map(|bytes| (part, bytes)))
    {
        Some((_part, bytes)) => parse_rich_value_relationship_indices(bytes)?,
        None => Vec::new(),
    };

    let local_image_identifier_to_rid = match rich_value_rel_part
        .as_deref()
        .and_then(|part| pkg.part(part).map(|bytes| (part, bytes)))
    {
        Some((_part, bytes)) => parse_rich_value_rel_ids(bytes)?,
        None => Vec::new(),
    };

    let rid_to_target: HashMap<String, String> = match rich_value_rel_part.as_deref() {
        Some(rich_value_rel_part) => {
            let rels_part = path::rels_for_part(rich_value_rel_part);
            let rels_bytes = match pkg.part(&rels_part) {
                Some(bytes) => bytes,
                None => &[],
            };

            if rels_bytes.is_empty() {
                HashMap::new()
            } else {
                let rels = openxml::parse_relationships(rels_bytes)?;
                rels.into_iter()
                    .filter(|rel| {
                        rel.type_uri.eq_ignore_ascii_case(REL_TYPE_IMAGE)
                            && !rel
                                .target_mode
                                .as_deref()
                                .is_some_and(|mode| mode.trim().eq_ignore_ascii_case("External"))
                    })
                    .map(|rel| {
                        let target = path::resolve_target(rich_value_rel_part, &rel.target);
                        (rel.id, target)
                    })
                    .collect()
            }
        }
        None => HashMap::new(),
    };

    let worksheet_parts: Vec<String> = match pkg.worksheet_parts() {
        Ok(infos) => infos.into_iter().map(|info| info.worksheet_part).collect(),
        Err(XlsxError::MissingPart(_))
        | Err(XlsxError::Invalid(_))
        | Err(XlsxError::InvalidSheetId) => pkg
            .part_names()
            .filter(|name| name.starts_with("xl/worksheets/") && name.ends_with(".xml"))
            .map(str::to_string)
            .collect(),
        Err(err) => return Err(err),
    };

    let mut out = Vec::new();
    for worksheet_part in worksheet_parts {
        let sheet_bytes = match pkg.part(&worksheet_part) {
            Some(bytes) => bytes,
            None => continue,
        };

        let cells_with_metadata =
            scan_cells_with_metadata_indices(sheet_bytes).map_err(|err| match err {
                RichDataError::Xlsx(err) => err,
                RichDataError::XmlNonUtf8 { source, .. } => XlsxError::Invalid(format!(
                    "worksheet xml contains invalid UTF-8: {source}"
                )),
                RichDataError::XmlParse { source, .. } => XlsxError::RoXml(source),
            })?;

        // Most worksheets contain very few `vm`-annotated cells (only those that actually reference
        // rich values). By streaming-scan filtering first, we avoid parsing `CellRef` for every
        // plain cell in large sheets.
        for (cell, vm, _cm) in cells_with_metadata {
            let Some(vm) = vm else { continue };

            let Some(rich_value_index) = vm_to_rich_value_index.get(&vm).copied() else {
                continue;
            };
            let (local_image_identifier, alt_text, decorative) =
                if let Some(rich_value) = rich_values.get(&rich_value_index) {
                    (
                        Some(rich_value.local_image_identifier),
                        rich_value.alt_text.clone(),
                        rich_value.calc_origin == 5,
                    )
                } else {
                    let idx = rich_value_rel_indices
                        .get(rich_value_index as usize)
                        .copied()
                        .flatten()
                        .and_then(|idx| u32::try_from(idx).ok());
                    (idx, None, false)
                };
            let Some(local_image_identifier) = local_image_identifier else {
                continue;
            };

            let rid = match local_image_identifier_to_rid.get(local_image_identifier as usize) {
                Some(rid) => rid,
                None => continue,
            };

            let target = match rid_to_target.get(rid) {
                Some(target) => target,
                None => continue,
            };

            let bytes = match pkg.part(target) {
                Some(bytes) => bytes.to_vec(),
                None => continue,
            };

            out.push(EmbeddedImageCell {
                sheet_part: worksheet_part.clone(),
                cell,
                image_target: target.clone(),
                bytes,
                alt_text,
                decorative,
            });
        }
    }

    Ok(out)
}

fn parse_vm_to_rich_value_index(xml: &[u8]) -> Result<HashMap<u32, u32>, XlsxError> {
    #[derive(Debug, Clone, Copy)]
    struct BkRun<T> {
        count: u32,
        value: T,
    }

    fn parse_bk_count(e: &quick_xml::events::BytesStart<'_>) -> Result<u32, XlsxError> {
        for attr in e.attributes() {
            let attr = attr?;
            if openxml::local_name(attr.key.as_ref()).eq_ignore_ascii_case(b"count") {
                return Ok(attr
                    .unescape_value()?
                    .trim()
                    .parse::<u32>()
                    .ok()
                    .filter(|v| *v >= 1)
                    .unwrap_or(1));
            }
        }
        Ok(1)
    }

    fn resolve_bk_run<T: Copy>(runs: &[BkRun<T>], idx: u32) -> Option<T> {
        let mut cursor: u32 = 0;
        for run in runs {
            let count = run.count.max(1);
            let end = cursor.saturating_add(count);
            if idx < end {
                return Some(run.value);
            }
            cursor = end;
        }
        None
    }

    let mut reader = Reader::from_reader(Cursor::new(xml));
    reader.config_mut().trim_text(true);

    let mut buf = Vec::new();
    let mut xlr_type_index: Option<u32> = None;
    // Excel has been observed to emit `<rc t="...">` as either 0-based or 1-based indices into
    // the `<metadataTypes>` list. Track the 0-based index while parsing and accept both later.
    let mut next_metadata_type_index = 0u32;

    // `futureMetadata` bk entries in order: each entry contains `xlrd:rvb i="..."`
    let mut future_bk_rich_value_index: Vec<BkRun<Option<u32>>> = Vec::new();
    let mut in_future_xlrichvalue = false;
    let mut current_future_bk: Option<usize> = None;

    // `valueMetadata` bk entries in order: each contains `rc t="..." v="..."`
    let mut value_bk_rc_records: Vec<BkRun<Vec<(u32, u32)>>> = Vec::new();
    let mut in_value_metadata = false;
    let mut current_value_bk: Option<usize> = None;

    loop {
        match reader.read_event_into(&mut buf)? {
            Event::Eof => break,
            Event::Start(e) | Event::Empty(e) => {
                let name = e.local_name();
                let name = name.as_ref();

                if name.eq_ignore_ascii_case(b"metadataType") {
                    let mut type_name: Option<String> = None;
                    for attr in e.attributes() {
                        let attr = attr?;
                        if openxml::local_name(attr.key.as_ref()).eq_ignore_ascii_case(b"name") {
                            type_name = Some(attr.unescape_value()?.into_owned());
                            break;
                        }
                    }

                    if type_name
                        .as_deref()
                        .is_some_and(|v| v.trim().eq_ignore_ascii_case("XLRICHVALUE"))
                    {
                        xlr_type_index = Some(next_metadata_type_index);
                    }
                    next_metadata_type_index = next_metadata_type_index.saturating_add(1);
                } else if name.eq_ignore_ascii_case(b"futureMetadata") {
                    let mut future_name: Option<String> = None;
                    for attr in e.attributes() {
                        let attr = attr?;
                        if openxml::local_name(attr.key.as_ref()).eq_ignore_ascii_case(b"name") {
                            future_name = Some(attr.unescape_value()?.into_owned());
                            break;
                        }
                    }
                    in_future_xlrichvalue = future_name
                        .as_deref()
                        .is_some_and(|v| v.trim().eq_ignore_ascii_case("XLRICHVALUE"));
                    current_future_bk = None;
                } else if in_future_xlrichvalue && name.eq_ignore_ascii_case(b"bk") {
                    let count = parse_bk_count(&e)?;
                    future_bk_rich_value_index.push(BkRun { count, value: None });
                    current_future_bk = Some(future_bk_rich_value_index.len() - 1);
                } else if in_future_xlrichvalue && name.eq_ignore_ascii_case(b"rvb") {
                    let mut i_value: Option<u32> = None;
                    for attr in e.attributes() {
                        let attr = attr?;
                        if openxml::local_name(attr.key.as_ref()).eq_ignore_ascii_case(b"i") {
                            i_value = attr
                                .unescape_value()?
                                .parse::<u32>()
                                .ok();
                            break;
                        }
                    }
                    if let (Some(idx), Some(i_value)) = (current_future_bk, i_value) {
                        if future_bk_rich_value_index
                            .get(idx)
                            .is_some_and(|run| run.value.is_none())
                        {
                            future_bk_rich_value_index[idx].value = Some(i_value);
                        }
                    }
                } else if name.eq_ignore_ascii_case(b"valueMetadata") {
                    in_value_metadata = true;
                    current_value_bk = None;
                } else if in_value_metadata && name.eq_ignore_ascii_case(b"bk") {
                    let count = parse_bk_count(&e)?;
                    value_bk_rc_records.push(BkRun {
                        count,
                        value: Vec::new(),
                    });
                    current_value_bk = Some(value_bk_rc_records.len() - 1);
                } else if in_value_metadata && name.eq_ignore_ascii_case(b"rc") {
                    let mut t: Option<u32> = None;
                    let mut v: Option<u32> = None;
                    for attr in e.attributes() {
                        let attr = attr?;
                        let key = openxml::local_name(attr.key.as_ref());
                        if key.eq_ignore_ascii_case(b"t") {
                            t = attr.unescape_value()?.parse::<u32>().ok();
                        } else if key.eq_ignore_ascii_case(b"v") {
                            v = attr.unescape_value()?.parse::<u32>().ok();
                        }
                    }

                    if let (Some(current), Some(t), Some(v)) = (current_value_bk, t, v) {
                        if let Some(list) = value_bk_rc_records.get_mut(current) {
                            list.value.push((t, v));
                        }
                    }
                }
            }
            Event::End(e) => {
                let name = e.local_name();
                let name = name.as_ref();

                if name.eq_ignore_ascii_case(b"futureMetadata") {
                    in_future_xlrichvalue = false;
                    current_future_bk = None;
                } else if in_future_xlrichvalue && name.eq_ignore_ascii_case(b"bk") {
                    current_future_bk = None;
                } else if name.eq_ignore_ascii_case(b"valueMetadata") {
                    in_value_metadata = false;
                    current_value_bk = None;
                } else if in_value_metadata && name.eq_ignore_ascii_case(b"bk") {
                    current_value_bk = None;
                }
            }
            _ => {}
        }
        buf.clear();
    }

    let Some(xlr_type_index) = xlr_type_index else {
        return Ok(HashMap::new());
    };

    let mut vm_to_rich_value_index: HashMap<u32, u32> = HashMap::new();
    let mut vm_idx_1_based: u32 = 1;
    for bk in value_bk_rc_records {
        let count = bk.count.max(1);
        // Excel emits `rc/@t` (metadata type index) as either 0-based *or* 1-based depending on
        // version/producer. We derive a 0-based index from the `<metadataTypes>` list, so accept
        // both the exact value and its 1-based equivalent.
        let xlr_type_index_0 = xlr_type_index;
        let xlr_type_index_1 = xlr_type_index.saturating_add(1);
        let Some(future_idx) = bk.value.iter().find_map(|(t, v)| {
            (*t == xlr_type_index_0 || *t == xlr_type_index_1).then_some(*v)
        }) else {
            vm_idx_1_based = vm_idx_1_based.saturating_add(count);
            continue;
        };

        // In most workbooks, `rc/@v` is an index into the `<futureMetadata name="XLRICHVALUE">`
        // `<bk>` list and we must dereference `rvb/@i` to get the rich value record index.
        //
        // Some producers omit `<futureMetadata name="XLRICHVALUE">` entirely and store the rich
        // value record index directly in `rc/@v`. If we didn't find any `futureMetadata` rich-value
        // blocks, treat `rc/@v` as the record index.
        let rich_value_index = if future_bk_rich_value_index.is_empty() {
            Some(future_idx)
        } else {
            resolve_bk_run(&future_bk_rich_value_index, future_idx)
                .or_else(|| {
                    future_idx
                        .checked_sub(1)
                        .and_then(|idx| resolve_bk_run(&future_bk_rich_value_index, idx))
                })
                .flatten()
        };
        let Some(rich_value_index) = rich_value_index else {
            vm_idx_1_based = vm_idx_1_based.saturating_add(count);
            continue;
        };

        for offset in 0..count {
            let vm = vm_idx_1_based.saturating_add(offset);
            vm_to_rich_value_index.insert(vm, rich_value_index);
            if vm > 0 {
                vm_to_rich_value_index.entry(vm - 1).or_insert(rich_value_index);
            }
        }
        vm_idx_1_based = vm_idx_1_based.saturating_add(count);
    }

    Ok(vm_to_rich_value_index)
}

#[cfg(test)]
mod tests {
    use super::parse_vm_to_rich_value_index;

    #[test]
    fn metadata_bk_count_is_respected_for_embedded_images_vm_mapping() {
        // Ensure `<bk count="N">` acts as run-length encoding for both:
        // - valueMetadata `vm` indices
        // - futureMetadata `rc/@v` indices
        let xml = br#"<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<metadata xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main"
          xmlns:xlrd="http://schemas.microsoft.com/office/spreadsheetml/2017/richdata">
  <metadataTypes count="1">
    <metadataType name="XLRICHVALUE"/>
  </metadataTypes>
  <futureMetadata name="XLRICHVALUE">
    <bk count="2">
      <extLst>
        <ext uri="{00000000-0000-0000-0000-000000000000}">
          <xlrd:rvb i="5"/>
        </ext>
      </extLst>
    </bk>
  </futureMetadata>
  <valueMetadata count="3">
    <bk count="3"><rc t="1" v="1"/></bk>
  </valueMetadata>
</metadata>
"#;

        let map = parse_vm_to_rich_value_index(xml).expect("parse vm->rich value");
        assert_eq!(map.get(&1), Some(&5));
        assert_eq!(map.get(&2), Some(&5));
        assert_eq!(map.get(&3), Some(&5));
    }

    #[test]
    fn metadata_rc_v_can_store_rich_value_index_directly_when_futuremetadata_missing() {
        // Some producers omit `<futureMetadata name="XLRICHVALUE">` and store the rich value index
        // directly in `rc/@v`. Ensure we still build a vm->rich-value mapping in that case.
        let xml = br#"<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<metadata xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main">
  <metadataTypes count="1">
    <metadataType name="XLRICHVALUE"/>
  </metadataTypes>
  <valueMetadata count="1">
    <bk>
      <rc t="1" v="7"/>
    </bk>
  </valueMetadata>
</metadata>
"#;

        let map = parse_vm_to_rich_value_index(xml).expect("parse vm->rich value");
        assert_eq!(map.get(&1), Some(&7));
    }
}

fn parse_local_image_structure_positions(
    xml: &[u8],
) -> Result<Vec<Option<LocalImageStructurePositions>>, XlsxError> {
    let mut reader = Reader::from_reader(Cursor::new(xml));
    reader.config_mut().trim_text(true);

    let mut buf = Vec::new();
    let mut out = Vec::new();

    let mut in_s = false;
    let mut current_is_local_image = false;
    let mut current_key_idx = 0usize;
    let mut current_positions = LocalImageStructurePositions::default();

    loop {
        match reader.read_event_into(&mut buf)? {
            Event::Eof => break,
            Event::Start(e) if e.local_name().as_ref().eq_ignore_ascii_case(b"s") => {
                in_s = true;
                current_is_local_image = false;
                current_key_idx = 0;
                current_positions = LocalImageStructurePositions::default();

                for attr in e.attributes() {
                    let attr = attr?;
                    if openxml::local_name(attr.key.as_ref()).eq_ignore_ascii_case(b"t") {
                        let v = attr.unescape_value()?.into_owned();
                        let local = v.rsplit(':').next().unwrap_or(v.as_str());
                        current_is_local_image = local.trim().eq_ignore_ascii_case("_localImage");
                        break;
                    }
                }
            }
            Event::Empty(e) if e.local_name().as_ref().eq_ignore_ascii_case(b"s") => {
                // Empty `<s/>` structure with no keys.
                let mut is_local_image = false;
                for attr in e.attributes() {
                    let attr = attr?;
                    if openxml::local_name(attr.key.as_ref()).eq_ignore_ascii_case(b"t") {
                        let v = attr.unescape_value()?.into_owned();
                        let local = v.rsplit(':').next().unwrap_or(v.as_str());
                        is_local_image = local.trim().eq_ignore_ascii_case("_localImage");
                        break;
                    }
                }
                out.push(is_local_image.then_some(LocalImageStructurePositions::default()));
                in_s = false;
            }
            Event::End(e) if e.local_name().as_ref().eq_ignore_ascii_case(b"s") => {
                if in_s {
                    out.push(current_is_local_image.then_some(std::mem::take(
                        &mut current_positions,
                    )));
                }
                in_s = false;
                current_is_local_image = false;
                current_key_idx = 0;
            }
            Event::Start(e) | Event::Empty(e)
                if in_s && e.local_name().as_ref().eq_ignore_ascii_case(b"k") =>
            {
                if current_is_local_image {
                    let mut key_name: Option<String> = None;
                    for attr in e.attributes() {
                        let attr = attr?;
                        if openxml::local_name(attr.key.as_ref()).eq_ignore_ascii_case(b"n") {
                            key_name = Some(attr.unescape_value()?.into_owned());
                            break;
                        }
                    }

                    if let Some(key_name) = key_name {
                        let local = key_name.rsplit(':').next().unwrap_or(key_name.as_str());
                        if local.eq_ignore_ascii_case("LocalImageIdentifier") {
                            if current_positions.local_image_identifier.is_none() {
                                current_positions.local_image_identifier = Some(current_key_idx);
                            }
                        } else if local.eq_ignore_ascii_case("CalcOrigin") {
                            if current_positions.calc_origin.is_none() {
                                current_positions.calc_origin = Some(current_key_idx);
                            }
                        } else if local.eq_ignore_ascii_case("Text") {
                            if current_positions.text.is_none() {
                                current_positions.text = Some(current_key_idx);
                            }
                        }
                    }
                }
                current_key_idx = current_key_idx.saturating_add(1);
            }
            _ => {}
        }
        buf.clear();
    }

    Ok(out)
}

fn parse_rdrichvalue(
    xml: &[u8],
    structure_positions: Option<&[Option<LocalImageStructurePositions>]>,
) -> Result<HashMap<u32, RichValueImage>, XlsxError> {
    let mut reader = Reader::from_reader(Cursor::new(xml));
    reader.config_mut().trim_text(true);

    let mut buf = Vec::new();
    let mut in_rv = false;
    let mut in_v = false;
    let mut current_v_text = String::new();
    let mut current_values: Vec<String> = Vec::new();
    let mut current_structure_index: Option<usize> = None;
    let mut rich_value_index: u32 = 0;
    let mut out: HashMap<u32, RichValueImage> = HashMap::new();

    loop {
        match reader.read_event_into(&mut buf)? {
            Event::Eof => break,
            Event::Start(e) if e.local_name().as_ref().eq_ignore_ascii_case(b"rv") => {
                in_rv = true;
                in_v = false;
                current_v_text.clear();
                current_values.clear();
                current_structure_index = None;
                for attr in e.attributes() {
                    let attr = attr?;
                    if openxml::local_name(attr.key.as_ref()).eq_ignore_ascii_case(b"s") {
                        current_structure_index = attr
                            .unescape_value()?
                            .parse::<usize>()
                            .ok();
                    }
                }
            }
            Event::Empty(e) if e.local_name().as_ref().eq_ignore_ascii_case(b"rv") => {
                // Empty rv elements aren't expected for local images.
                in_rv = false;
                in_v = false;
                current_v_text.clear();
                current_structure_index = None;
                rich_value_index = rich_value_index.saturating_add(1);
            }
            Event::End(e) if e.local_name().as_ref().eq_ignore_ascii_case(b"rv") => {
                in_rv = false;
                in_v = false;
                current_v_text.clear();

                let mut structure: Option<&LocalImageStructurePositions> = None;
                let mut known_non_local = false;
                if let (Some(structs), Some(idx)) = (structure_positions, current_structure_index) {
                    match structs.get(idx) {
                        Some(Some(pos)) => structure = Some(pos),
                        Some(None) => known_non_local = true,
                        None => {}
                    }
                }

                if known_non_local {
                    // This rich value record doesn't correspond to a `_localImage` structure.
                    // Skip it rather than attempting to interpret positional values.
                    current_structure_index = None;
                    rich_value_index = rich_value_index.saturating_add(1);
                    continue;
                }

                let mut local_image_identifier = structure
                    .and_then(|s| s.local_image_identifier)
                    .and_then(|idx| current_values.get(idx))
                    .and_then(|v| v.trim().parse::<u32>().ok());
                let mut calc_origin = structure
                    .and_then(|s| s.calc_origin)
                    .and_then(|idx| current_values.get(idx))
                    .and_then(|v| v.trim().parse::<u32>().ok());

                let mut alt_text = structure
                    .and_then(|s| s.text)
                    .and_then(|idx| current_values.get(idx))
                    .map(|v| v.trim().to_string())
                    .filter(|v| !v.is_empty());

                // Fallback parsing if structure metadata is missing or incomplete.
                if local_image_identifier.is_none() {
                    local_image_identifier = current_values
                        .get(0)
                        .and_then(|v| v.trim().parse::<u32>().ok());
                }
                if calc_origin.is_none() {
                    calc_origin = current_values
                        .get(1)
                        .and_then(|v| v.trim().parse::<u32>().ok());
                }
                if alt_text.is_none() && structure.is_none() {
                    alt_text = current_values
                        .get(2)
                        .map(|v| v.trim().to_string())
                        .filter(|v| !v.is_empty());
                }

                if let (Some(local_image_identifier), Some(calc_origin)) =
                    (local_image_identifier, calc_origin)
                {
                    out.insert(
                        rich_value_index,
                        RichValueImage {
                            local_image_identifier,
                            calc_origin,
                            alt_text,
                        },
                    );
                }

                current_structure_index = None;
                rich_value_index = rich_value_index.saturating_add(1);
            }
            Event::Start(e) if in_rv && e.local_name().as_ref().eq_ignore_ascii_case(b"v") => {
                in_v = true;
                current_v_text.clear();
            }
            Event::Empty(e) if in_rv && e.local_name().as_ref().eq_ignore_ascii_case(b"v") => {
                current_values.push(String::new());
                in_v = false;
                current_v_text.clear();
            }
            Event::End(e) if in_rv && e.local_name().as_ref().eq_ignore_ascii_case(b"v") => {
                current_values.push(std::mem::take(&mut current_v_text));
                in_v = false;
            }
            Event::Text(e) if in_rv && in_v => {
                current_v_text.push_str(&e.unescape()?.into_owned());
            }
            Event::CData(e) if in_rv && in_v => {
                current_v_text.push_str(
                    &e.decode()
                        .map_err(quick_xml::Error::from)?
                        .into_owned(),
                );
            }
            _ => {}
        }
        buf.clear();
    }

    Ok(out)
}

fn parse_rich_value_rel_ids(xml: &[u8]) -> Result<Vec<String>, XlsxError> {
    let mut reader = Reader::from_reader(Cursor::new(xml));
    reader.config_mut().trim_text(true);

    let mut buf = Vec::new();
    let mut out: Vec<String> = Vec::new();

    loop {
        match reader.read_event_into(&mut buf)? {
            Event::Eof => break,
            Event::Start(e) | Event::Empty(e) if e.local_name().as_ref() == b"rel" => {
                let mut rid: Option<String> = None;
                for attr in e.attributes() {
                    let attr = attr?;
                    if openxml::local_name(attr.key.as_ref()).eq_ignore_ascii_case(b"id") {
                        rid = Some(attr.unescape_value()?.into_owned());
                        break;
                    }
                }
                if let Some(rid) = rid {
                    out.push(rid);
                }
            }
            _ => {}
        }
        buf.clear();
    }

    Ok(out)
}
