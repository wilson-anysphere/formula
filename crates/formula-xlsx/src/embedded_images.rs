use std::collections::HashMap;
use std::io::Cursor;

use formula_model::CellRef;
use quick_xml::events::Event;
use quick_xml::Reader;

use crate::openxml;
use crate::package::{XlsxError, XlsxPackage};
use crate::path;
use crate::rich_data::metadata::parse_value_metadata_vm_to_rich_value_index_map;
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
const REL_TYPE_RICH_VALUE_2017: &str = "http://schemas.microsoft.com/office/2017/relationships/richValue";
const REL_TYPE_RICH_VALUE_REL: &str =
    "http://schemas.microsoft.com/office/2022/10/relationships/richValueRel";
const REL_TYPE_RICH_VALUE_REL_2017_06: &str =
    "http://schemas.microsoft.com/office/2017/06/relationships/richValueRel";
const REL_TYPE_RICH_VALUE_REL_2017: &str =
    "http://schemas.microsoft.com/office/2017/relationships/richValueRel";
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
            && (rel.type_uri.eq_ignore_ascii_case(REL_TYPE_RICH_VALUE)
                || rel.type_uri.eq_ignore_ascii_case(REL_TYPE_RICH_VALUE_2017))
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
                || rel.type_uri.eq_ignore_ascii_case(REL_TYPE_RICH_VALUE_REL_2017_06)
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

    // Some workbooks relate the richData parts from `xl/metadata.xml` via `xl/_rels/metadata.xml.rels`
    // (rather than directly from `xl/workbook.xml`). If we still haven't found the rich-value parts,
    // attempt to discover them via metadata relationships.
    if rich_value_part.is_none()
        || rdrichvalue_part.is_none()
        || rdrichvaluestructure_part.is_none()
        || rich_value_rel_part.is_none()
    {
        if let Some(metadata_part_name) = metadata_part.as_deref() {
            let metadata_rels_part = path::rels_for_part(metadata_part_name);
            if let Some(rels_bytes) = pkg.part(&metadata_rels_part) {
                let relationships = openxml::parse_relationships(rels_bytes)?;
                for rel in relationships {
                    let target = path::resolve_target(metadata_part_name, &rel.target);

                    if rich_value_part.is_none()
                        && (rel.type_uri.eq_ignore_ascii_case(REL_TYPE_RICH_VALUE)
                            || rel.type_uri.eq_ignore_ascii_case(REL_TYPE_RICH_VALUE_2017))
                    {
                        rich_value_part = Some(target);
                        continue;
                    }

                    if rich_value_rel_part.is_none()
                        && (rel.type_uri.eq_ignore_ascii_case(REL_TYPE_RICH_VALUE_REL)
                            || rel.type_uri.eq_ignore_ascii_case(REL_TYPE_RICH_VALUE_REL_2017_06)
                            || rel.type_uri.eq_ignore_ascii_case(REL_TYPE_RICH_VALUE_REL_2017))
                    {
                        rich_value_rel_part = Some(target);
                        continue;
                    }

                    if rdrichvalue_part.is_none()
                        && rel.type_uri.eq_ignore_ascii_case(REL_TYPE_RD_RICH_VALUE)
                    {
                        rdrichvalue_part = Some(target);
                        continue;
                    }

                    if rdrichvaluestructure_part.is_none()
                        && rel
                            .type_uri
                            .eq_ignore_ascii_case(REL_TYPE_RD_RICH_VALUE_STRUCTURE)
                    {
                        rdrichvaluestructure_part = Some(target);
                        continue;
                    }
                }
            }
        }
    }

    let vm_to_rich_value_index = match metadata_part
        .as_deref()
        .and_then(|part| pkg.part(part).map(|bytes| (part, bytes)))
    {
        Some((part, bytes)) => {
            parse_value_metadata_vm_to_rich_value_index_map(bytes).map_err(|err| match err {
                crate::xml::XmlDomError::Utf8(source) => {
                    XlsxError::Invalid(format!("xml part {part} is not valid UTF-8: {source}"))
                }
                crate::xml::XmlDomError::Parse(source) => {
                    XlsxError::Invalid(format!("xml parse error in {part}: {source}"))
                }
            })?
        }
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

    // Prefer parsing all `xl/richData/richValue*.xml` parts (concatenated in numeric-suffix order),
    // since some workbooks split the rich value table across multiple parts (e.g. `richValue1.xml`,
    // `richValue2.xml`, ...).
    //
    // Only do this when:
    // - we have not explicitly discovered a custom/nonstandard rich value part name, AND
    // - there is at least one standard `richValue*.xml` part present in the package.
    //
    // This preserves the ability to read synthetic/custom packages while also matching the rich
    // data extraction behavior in `crate::rich_data`.
    let rich_value_rel_indices = {
        let mut rich_value_parts: Vec<String> = pkg
            .part_names()
            .filter(|name| rich_value_part_suffix_index(name).is_some())
            .map(str::to_string)
            .collect();

        let should_use_multi_part = !rich_value_parts.is_empty()
            && (rich_value_part.is_none()
                || rich_value_part
                    .as_deref()
                    .is_some_and(|p| rich_value_part_suffix_index(p).is_some()));

        if should_use_multi_part {
            rich_value_parts.sort_by(|a, b| {
                let a_idx = rich_value_part_suffix_index(a).unwrap_or(u32::MAX);
                let b_idx = rich_value_part_suffix_index(b).unwrap_or(u32::MAX);
                a_idx.cmp(&b_idx).then_with(|| a.cmp(b))
            });

            let mut out: Vec<Option<usize>> = Vec::new();
            for part in rich_value_parts {
                let Some(bytes) = pkg.part(&part) else { continue };
                out.extend(parse_rich_value_relationship_indices(bytes)?);
            }
            out
        } else if let Some(bytes) = rich_value_part.as_deref().and_then(|part| pkg.part(part)) {
            parse_rich_value_relationship_indices(bytes)?
        } else {
            Vec::new()
        }
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
        //
        // Excel has been observed to encode worksheet `c/@vm` as both:
        // - 1-based indices into `xl/metadata.xml` `<valueMetadata>` `<bk>` records (typical), and
        // - 0-based indices (seen in some workbooks/fixtures).
        //
        // We keep the metadata mapping in its canonical 1-based form and infer whether a sheet is
        // 0-based by checking for any `vm="0"` cells. This avoids ambiguous key collisions that
        // occur if we try to store both 0- and 1-based mappings in the same map when multiple
        // images/records exist.
        let vm_offset: u32 =
            if !vm_to_rich_value_index.is_empty()
                && cells_with_metadata.iter().any(|(_, vm, _)| *vm == Some(0))
            {
                1
            } else {
                0
            };

        for (cell, vm, _cm) in cells_with_metadata {
            let Some(vm) = vm else { continue };

            let vm = vm.saturating_add(vm_offset);
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

fn rich_value_part_suffix_index(part_path: &str) -> Option<u32> {
    if !part_path.starts_with("xl/richData/") {
        return None;
    }
    let file_name = part_path.rsplit('/').next()?;
    let file_name_lower = file_name.to_ascii_lowercase();
    if !file_name_lower.ends_with(".xml") {
        return None;
    }

    let stem = &file_name_lower[..file_name_lower.len() - ".xml".len()];
    let suffix = stem.strip_prefix("richvalue")?;
    if suffix.is_empty() {
        return Some(0);
    }
    if !suffix.chars().all(|c| c.is_ascii_digit()) {
        return None;
    }
    suffix.parse::<u32>().ok()
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
                // Preserve missing `r:id` entries as placeholders to avoid shifting relationship
                // indices (Excel treats the `<rel>` list as a dense table).
                out.push(rid.unwrap_or_default());
            }
            _ => {}
        }
        buf.clear();
    }

    Ok(out)
}
