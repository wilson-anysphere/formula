use std::collections::HashMap;
use std::io::Cursor;

use formula_model::{CellRef, HyperlinkTarget};
use quick_xml::events::Event;
use quick_xml::Reader;

use crate::openxml;
use crate::package::{XlsxError, XlsxPackage};
use crate::path;
use crate::rich_data::metadata::parse_value_metadata_vm_to_rich_value_index_map;
use crate::rich_data::{scan_cells_with_metadata_indices, RichDataError};
use crate::rich_data::rich_value::parse_rich_value_relationship_indices;
use crate::drawings::REL_TYPE_IMAGE;

const REL_TYPE_SHEET_METADATA: &str =
    "http://schemas.openxmlformats.org/officeDocument/2006/relationships/sheetMetadata";
const REL_TYPE_METADATA: &str =
    "http://schemas.openxmlformats.org/officeDocument/2006/relationships/metadata";
const REL_TYPE_RD_RICH_VALUE: &str =
    "http://schemas.microsoft.com/office/2017/06/relationships/rdRichValue";
const REL_TYPE_RD_RICH_VALUE_2017: &str =
    "http://schemas.microsoft.com/office/2017/relationships/rdRichValue";
const REL_TYPE_RD_RICH_VALUE_STRUCTURE: &str =
    "http://schemas.microsoft.com/office/2017/06/relationships/rdRichValueStructure";
const REL_TYPE_RD_RICH_VALUE_STRUCTURE_2017: &str =
    "http://schemas.microsoft.com/office/2017/relationships/rdRichValueStructure";
const REL_TYPE_RICH_VALUE: &str = "http://schemas.microsoft.com/office/2017/06/relationships/richValue";
const REL_TYPE_RICH_VALUE_2017: &str = "http://schemas.microsoft.com/office/2017/relationships/richValue";
const REL_TYPE_RICH_VALUE_REL: &str =
    "http://schemas.microsoft.com/office/2022/10/relationships/richValueRel";
const REL_TYPE_RICH_VALUE_REL_2017_06: &str =
    "http://schemas.microsoft.com/office/2017/06/relationships/richValueRel";
const REL_TYPE_RICH_VALUE_REL_2017: &str =
    "http://schemas.microsoft.com/office/2017/relationships/richValueRel";

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
    /// Optional hyperlink target attached to the same worksheet cell.
    pub hyperlink_target: Option<HyperlinkTarget>,
    /// Whether the image appears to be marked as decorative.
    ///
    /// This is derived from the rich value `CalcOrigin` field using observed workbook behavior
    /// (`CalcOrigin == 5` is treated as decorative in this codebase's fixtures). The full enum is
    /// not publicly documented; treat this as best-effort.
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

#[derive(Debug, Clone, Copy, PartialEq, Eq)]
enum RelationshipIndexBase {
    ZeroBased,
    OneBased,
    Unknown,
}

fn infer_relationship_index_base(
    rich_values: &HashMap<u32, RichValueImage>,
    rich_value_rel_indices: &[Option<usize>],
    rel_ids: &[String],
    rid_to_target: &HashMap<String, String>,
) -> RelationshipIndexBase {
    let rel_count = rel_ids.len();
    if rel_count == 0 {
        return RelationshipIndexBase::Unknown;
    }
    let Ok(rel_count_u32) = u32::try_from(rel_count) else {
        return RelationshipIndexBase::Unknown;
    };

    let mut values: Vec<u32> = Vec::new();
    values.extend(rich_values.values().map(|rv| rv.local_image_identifier));
    values.extend(
        rich_value_rel_indices
            .iter()
            .filter_map(|v| (*v).and_then(|v| u32::try_from(v).ok())),
    );
    if values.is_empty() {
        return RelationshipIndexBase::Unknown;
    }

    let zero_possible = values.iter().all(|v| *v < rel_count_u32);
    let one_possible = values
        .iter()
        .all(|v| *v >= 1 && *v <= rel_count_u32);

    match (zero_possible, one_possible) {
        (true, false) => RelationshipIndexBase::ZeroBased,
        (false, true) => RelationshipIndexBase::OneBased,
        (false, false) => RelationshipIndexBase::Unknown,
        (true, true) => {
            // Ambiguous: prefer the interpretation that yields more resolved targets.
            let score = |base: RelationshipIndexBase| -> usize {
                values
                    .iter()
                    .filter(|v| {
                        resolve_image_target_for_local_image_identifier(
                            **v,
                            base,
                            rel_ids,
                            rid_to_target,
                        )
                        .is_some()
                    })
                    .count()
            };

            let score_zero = score(RelationshipIndexBase::ZeroBased);
            let score_one = score(RelationshipIndexBase::OneBased);

            match score_zero.cmp(&score_one) {
                std::cmp::Ordering::Greater => RelationshipIndexBase::ZeroBased,
                std::cmp::Ordering::Less => RelationshipIndexBase::OneBased,
                std::cmp::Ordering::Equal => {
                    if values.iter().any(|v| *v == 0) {
                        RelationshipIndexBase::ZeroBased
                    } else if values.iter().any(|v| *v == rel_count_u32) {
                        RelationshipIndexBase::OneBased
                    } else {
                        RelationshipIndexBase::Unknown
                    }
                }
            }
        }
    }
}

fn resolve_image_target_for_local_image_identifier<'a>(
    local_image_identifier: u32,
    base: RelationshipIndexBase,
    rel_ids: &'a [String],
    rid_to_target: &'a HashMap<String, String>,
) -> Option<&'a String> {
    let a = match base {
        RelationshipIndexBase::ZeroBased => Some(local_image_identifier),
        RelationshipIndexBase::OneBased => local_image_identifier.checked_sub(1),
        RelationshipIndexBase::Unknown => Some(local_image_identifier),
    };
    let b = match base {
        RelationshipIndexBase::Unknown => local_image_identifier.checked_sub(1),
        _ => None,
    };

    for idx in [a, b].into_iter().filter_map(|v| v) {
        let Ok(idx) = usize::try_from(idx) else {
            continue;
        };
        let Some(rid) = rel_ids.get(idx) else {
            continue;
        };
        if rid.is_empty() {
            continue;
        }
        if let Some(target) = rid_to_target.get(rid) {
            return Some(target);
        }
    }

    None
}

/// Extract embedded images-in-cells ("Place in Cell") using the `vm` + `metadata.xml` + `xl/richData/*`
/// schema.
///
/// The extractor is intentionally resilient: missing parts or broken mappings result in skipped cells.
/// Only invalid ZIP/XML/UTF-8 errors are returned as [`XlsxError`].
pub fn extract_embedded_images(pkg: &XlsxPackage) -> Result<Vec<EmbeddedImageCell>, XlsxError> {
    fn resolve_existing_part(pkg: &XlsxPackage, base_part: &str, target: &str) -> Option<String> {
        let candidates = crate::path::resolve_target_candidates(base_part, target);
        // Prefer a candidate that matches an *actual* stored ZIP part name. `XlsxPackage::part`
        // normalizes percent-encoded names, so it would return bytes for both the raw and decoded
        // candidates and prevent us from returning canonical part-name strings.
        for candidate in &candidates {
            if candidate.is_empty() {
                continue;
            }
            if pkg.parts_map().contains_key(candidate)
                || pkg.parts_map().contains_key(format!("/{candidate}").as_str())
            {
                return Some(candidate.clone());
            }
        }
        candidates
            .into_iter()
            .find(|candidate| pkg.part(candidate).is_some())
    }

    // Discover the relevant rich data parts via `xl/_rels/workbook.xml.rels`.
    let relationships = match pkg.part("xl/_rels/workbook.xml.rels") {
        Some(bytes) => openxml::parse_relationships(bytes)?,
        None => Vec::new(),
    };
    let mut metadata_part: Option<String> = None;
    let mut rich_value_part: Option<String> = None;
    let mut rdrichvalue_part: Option<String> = None;
    let mut rdrichvaluestructure_part: Option<String> = None;
    let mut rich_value_rel_part: Option<String> = None;

    for rel in relationships {
        if rel
            .target_mode
            .as_deref()
            .is_some_and(|mode| mode.trim().eq_ignore_ascii_case("External"))
        {
            continue;
        }

        if metadata_part.is_none()
            && (rel.type_uri.eq_ignore_ascii_case(REL_TYPE_SHEET_METADATA)
                || rel.type_uri.eq_ignore_ascii_case(REL_TYPE_METADATA))
        {
            metadata_part = resolve_existing_part(pkg, "xl/workbook.xml", &rel.target);
        } else if rich_value_part.is_none()
            && (rel.type_uri.eq_ignore_ascii_case(REL_TYPE_RICH_VALUE)
                || rel.type_uri.eq_ignore_ascii_case(REL_TYPE_RICH_VALUE_2017))
        {
            rich_value_part = resolve_existing_part(pkg, "xl/workbook.xml", &rel.target);
        } else if rdrichvalue_part.is_none()
            && (rel.type_uri.eq_ignore_ascii_case(REL_TYPE_RD_RICH_VALUE)
                || rel.type_uri.eq_ignore_ascii_case(REL_TYPE_RD_RICH_VALUE_2017))
        {
            rdrichvalue_part = resolve_existing_part(pkg, "xl/workbook.xml", &rel.target);
        } else if rdrichvaluestructure_part.is_none()
            && (rel
                .type_uri
                .eq_ignore_ascii_case(REL_TYPE_RD_RICH_VALUE_STRUCTURE)
                || rel
                    .type_uri
                    .eq_ignore_ascii_case(REL_TYPE_RD_RICH_VALUE_STRUCTURE_2017))
        {
            rdrichvaluestructure_part = resolve_existing_part(pkg, "xl/workbook.xml", &rel.target);
        } else if rich_value_rel_part.is_none()
            && (rel.type_uri.eq_ignore_ascii_case(REL_TYPE_RICH_VALUE_REL)
                || rel.type_uri.eq_ignore_ascii_case(REL_TYPE_RICH_VALUE_REL_2017_06)
                || rel.type_uri.eq_ignore_ascii_case(REL_TYPE_RICH_VALUE_REL_2017))
        {
            rich_value_rel_part = resolve_existing_part(pkg, "xl/workbook.xml", &rel.target);
        }
    }

    // Fallback to canonical part names when the workbook relationship graph doesn't include the
    // expected rich-data relationships (best-effort).
    if metadata_part.is_none() && pkg.part("xl/metadata.xml").is_some() {
        metadata_part = Some("xl/metadata.xml".to_string());
    }
    if rich_value_part.is_none() {
        if pkg.part("xl/richData/richValue.xml").is_some() {
            rich_value_part = Some("xl/richData/richValue.xml".to_string());
        } else if pkg.part("xl/richData/richValues.xml").is_some() {
            // Some producers use the pluralized `richValues*.xml` naming pattern.
            rich_value_part = Some("xl/richData/richValues.xml".to_string());
        }
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
                    let Some(target) = resolve_existing_part(pkg, metadata_part_name, &rel.target)
                    else {
                        continue;
                    };

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
                        && (rel.type_uri.eq_ignore_ascii_case(REL_TYPE_RD_RICH_VALUE)
                            || rel.type_uri.eq_ignore_ascii_case(REL_TYPE_RD_RICH_VALUE_2017))
                    {
                        rdrichvalue_part = Some(target);
                        continue;
                    }

                    if rdrichvaluestructure_part.is_none()
                        && (rel
                            .type_uri
                            .eq_ignore_ascii_case(REL_TYPE_RD_RICH_VALUE_STRUCTURE)
                            || rel
                                .type_uri
                                .eq_ignore_ascii_case(REL_TYPE_RD_RICH_VALUE_STRUCTURE_2017))
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
                        let target = rel.target.split('#').next().unwrap_or(&rel.target);
                        // Be resilient to invalid/unescaped Windows-style path separators.
                        let target: std::borrow::Cow<'_, str> = if target.contains('\\') {
                            std::borrow::Cow::Owned(target.replace('\\', "/"))
                        } else {
                            std::borrow::Cow::Borrowed(target)
                        };
                        let target = target.as_ref();
                        let target = target.strip_prefix("./").unwrap_or(target);
                        // Some producers emit `Target="media/image1.png"` (relative to `xl/`)
                        // rather than `Target="../media/image1.png"` (relative to `xl/richData/`).
                        let resolved = if target.starts_with("media/") {
                            let absolute = format!("xl/{target}");
                            resolve_existing_part(pkg, "", &absolute).unwrap_or(absolute)
                        } else if target.starts_with("xl/") {
                            resolve_existing_part(pkg, "", target)
                                .unwrap_or_else(|| target.to_string())
                        } else {
                            resolve_existing_part(pkg, rich_value_rel_part, target)
                                .unwrap_or_else(|| path::resolve_target(rich_value_rel_part, target))
                        };
                        (rel.id, resolved)
                     })
                     .collect()
             }
         }
        None => HashMap::new(),
    };

    let rel_index_base = infer_relationship_index_base(
        &rich_values,
        &rich_value_rel_indices,
        &local_image_identifier_to_rid,
        &rid_to_target,
    );

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
        let sheet_xml = std::str::from_utf8(sheet_bytes)
            .map_err(|e| XlsxError::Invalid(format!("{worksheet_part} not utf-8: {e}")))?;

        let sheet_rels_part = path::rels_for_part(&worksheet_part);
        let sheet_rels_xml = pkg
            .part(&sheet_rels_part)
            .and_then(|bytes| std::str::from_utf8(bytes).ok());

        // Best-effort: if hyperlink parsing fails (malformed file), still extract images.
        let hyperlinks =
            crate::parse_worksheet_hyperlinks(sheet_xml, sheet_rels_xml).unwrap_or_default();

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
        // Excel typically stores worksheet `c/@vm` as a 1-based index into `xl/metadata.xml`
        // `<valueMetadata>` `<bk>` records, but some producers have been observed to use 0-based
        // indices instead.
        //
        // The metadata parser returns a canonical 1-based `vm -> richValue` mapping. To tolerate
        // 0-based worksheets we choose a per-worksheet offset (0 or 1) that yields the most
        // successfully-resolved images. If the result is ambiguous, we still treat `vm="0"` as
        // strong evidence of 0-based indexing.
        let vm_offset: u32 =
            if vm_to_rich_value_index.is_empty() || cells_with_metadata.is_empty() {
                0
            } else {
                let has_vm_zero = cells_with_metadata.iter().any(|(_, vm, _)| *vm == Some(0));
                let mut resolved_offset_0 = 0usize;
                let mut resolved_offset_1 = 0usize;

                let resolves_image = |vm_raw: u32, offset: u32| -> bool {
                    let vm = vm_raw.saturating_add(offset);
                    let Some(rich_value_index) = vm_to_rich_value_index.get(&vm).copied() else {
                        return false;
                    };

                    let local_image_identifier =
                        if let Some(rich_value) = rich_values.get(&rich_value_index) {
                            Some(rich_value.local_image_identifier)
                        } else {
                            rich_value_rel_indices
                                .get(rich_value_index as usize)
                                .copied()
                                .flatten()
                                .and_then(|idx| u32::try_from(idx).ok())
                        };
                    let Some(local_image_identifier) = local_image_identifier else {
                        return false;
                    };

                    let Some(rid) =
                        local_image_identifier_to_rid.get(local_image_identifier as usize)
                    else {
                        return false;
                    };
                    let Some(target) = rid_to_target.get(rid) else {
                        return false;
                    };

                    pkg.part(target).is_some()
                };

                for (_cell, vm, _cm) in &cells_with_metadata {
                    let Some(vm) = *vm else {
                        continue;
                    };
                    if resolves_image(vm, 0) {
                        resolved_offset_0 += 1;
                    }
                    if resolves_image(vm, 1) {
                        resolved_offset_1 += 1;
                    }
                }

                if resolved_offset_1 > resolved_offset_0 {
                    1
                } else if resolved_offset_0 > resolved_offset_1 {
                    0
                } else if has_vm_zero {
                    1
                } else {
                    0
                }
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

            let target = match resolve_image_target_for_local_image_identifier(
                local_image_identifier,
                rel_index_base,
                &local_image_identifier_to_rid,
                &rid_to_target,
            ) {
                Some(target) => target,
                None => continue,
            };

            let bytes = match pkg.part(target) {
                Some(bytes) => bytes.to_vec(),
                None => continue,
            };

            let hyperlink_target = hyperlinks
                .iter()
                .find(|link| link.range.contains(cell))
                .map(|link| link.target.clone());

            out.push(EmbeddedImageCell {
                sheet_part: worksheet_part.clone(),
                cell,
                image_target: target.clone(),
                bytes,
                alt_text,
                hyperlink_target,
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
    let stem = crate::ascii::strip_suffix_ignore_case(file_name, ".xml")?;
    // Check the plural prefix first: `richvalues` starts with `richvalue`.
    let suffix = if let Some(rest) = crate::ascii::strip_prefix_ignore_case(stem, "richvalues") {
        rest
    } else if let Some(rest) = crate::ascii::strip_prefix_ignore_case(stem, "richvalue") {
        rest
    } else {
        return None;
    };
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
                let mut rid_prefixed: Option<String> = None;
                let mut rid_unprefixed: Option<String> = None;
                for attr in e.attributes() {
                    let attr = attr?;
                    let key = attr.key.as_ref();
                    if openxml::local_name(key).eq_ignore_ascii_case(b"id") {
                        let value = attr.unescape_value()?.into_owned();
                        let trimmed = value.trim();
                        if trimmed.is_empty() {
                            continue;
                        }

                        // Prefer namespaced `r:id`, but be tolerant of unqualified `id`.
                        if key.iter().any(|b| *b == b':') {
                            rid_prefixed = Some(trimmed.to_string());
                        } else {
                            rid_unprefixed = Some(trimmed.to_string());
                        }
                    }
                }
                // Preserve missing `r:id` entries as placeholders to avoid shifting relationship
                // indices (Excel treats the `<rel>` list as a dense table).
                out.push(rid_prefixed.or(rid_unprefixed).unwrap_or_default());
            }
            _ => {}
        }
        buf.clear();
    }

    Ok(out)
}

#[cfg(test)]
mod tests {
    use super::*;

    use std::io::{Cursor, Write};

    use zip::write::FileOptions;
    use zip::ZipWriter;

    fn build_package(entries: &[(&str, &[u8])]) -> XlsxPackage {
        let cursor = Cursor::new(Vec::new());
        let mut zip = ZipWriter::new(cursor);
        let options =
            FileOptions::<()>::default().compression_method(zip::CompressionMethod::Deflated);

        for (name, bytes) in entries {
            zip.start_file(*name, options).unwrap();
            zip.write_all(bytes).unwrap();
        }

        let bytes = zip.finish().unwrap().into_inner();
        XlsxPackage::from_bytes(&bytes).expect("read test pkg")
    }

    #[test]
    fn ignores_external_workbook_relationships_for_rich_data_parts() {
        // External workbook relationships should not prevent discovery of internal rich-data parts.
        // If we accidentally select an external target, we won't find `xl/metadata.xml` or the
        // `xl/richData/*` tables and extraction will incorrectly return an empty result.
        let workbook_rels = br#"<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
  <Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/metadata" Target="https://example.com/metadata.xml" TargetMode="External"/>
  <Relationship Id="rId2" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/metadata" Target="metadata.xml"/>
  <Relationship Id="rId3" Type="http://schemas.microsoft.com/office/2017/06/relationships/richValue" Target="https://example.com/richData/richValue.xml" TargetMode="External"/>
  <Relationship Id="rId4" Type="http://schemas.microsoft.com/office/2017/06/relationships/richValue" Target="richData/richValue.xml"/>
  <Relationship Id="rId5" Type="http://schemas.microsoft.com/office/2022/10/relationships/richValueRel" Target="https://example.com/richData/richValueRel.xml" TargetMode="External"/>
  <Relationship Id="rId6" Type="http://schemas.microsoft.com/office/2022/10/relationships/richValueRel" Target="richData/richValueRel.xml"/>
</Relationships>"#;

        let metadata = br#"<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<metadata xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main">
  <metadataTypes count="1">
    <metadataType name="XLRICHVALUE"/>
  </metadataTypes>
  <valueMetadata count="1">
    <bk><rc t="1" v="0"/></bk>
  </valueMetadata>
</metadata>"#;

        let worksheet = br#"<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main">
  <sheetData>
    <row r="1">
      <c r="A1" vm="1"/>
    </row>
  </sheetData>
</worksheet>"#;

        let rich_value = br#"<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<richValue>
  <rv><v kind="rel">0</v></rv>
</richValue>"#;

        let rich_value_rel = br#"<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<richValueRel xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">
  <rel r:id="rId1"/>
</richValueRel>"#;

        let rich_value_rel_rels = br#"<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
  <Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/image" Target="../media/image1.png"/>
</Relationships>"#;

        let pkg = build_package(&[
            ("xl/_rels/workbook.xml.rels", workbook_rels),
            ("xl/metadata.xml", metadata),
            ("xl/worksheets/sheet1.xml", worksheet),
            ("xl/richData/richValue.xml", rich_value),
            ("xl/richData/richValueRel.xml", rich_value_rel),
            ("xl/richData/_rels/richValueRel.xml.rels", rich_value_rel_rels),
            ("xl/media/image1.png", b"png-bytes"),
        ]);

        let extracted = extract_embedded_images(&pkg).expect("extract embedded images");
        assert_eq!(extracted.len(), 1);
        assert_eq!(extracted[0].sheet_part, "xl/worksheets/sheet1.xml");
        assert_eq!(extracted[0].cell, CellRef::from_a1("A1").unwrap());
        assert_eq!(extracted[0].image_target, "xl/media/image1.png");
        assert_eq!(extracted[0].bytes, b"png-bytes");
        assert_eq!(extracted[0].alt_text, None);
        assert!(!extracted[0].decorative);
    }
}
