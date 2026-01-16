use std::collections::HashMap;
use std::io::Cursor;

use formula_model::{CellRef, HyperlinkTarget};
use quick_xml::events::Event;
use quick_xml::name::QName;
use quick_xml::Reader;

use crate::drawings::REL_TYPE_IMAGE;
use crate::path::{rels_for_part, resolve_target};
use crate::rich_data::{scan_cells_with_metadata_indices, RichDataError};
use crate::{parse_worksheet_hyperlinks, XlsxError, XlsxPackage};

const RICH_VALUE_REL_PART: &str = "xl/richData/richValueRel.xml";
const RD_RICH_VALUE_PART: &str = "xl/richData/rdrichvalue.xml";
const RD_RICH_VALUE_STRUCTURE_PART: &str = "xl/richData/rdrichvaluestructure.xml";

/// An embedded image stored inside a cell using Excel's RichData / `vm=` mechanism.
///
/// These are distinct from DrawingML images (floating/anchored shapes).
#[derive(Debug, Clone, PartialEq, Eq)]
pub struct EmbeddedCellImage {
    /// Resolved package part name for the image (e.g. `xl/media/image1.png`).
    pub image_part: String,
    /// Raw bytes for the image file.
    pub image_bytes: Vec<u8>,
    /// Rich value `CalcOrigin` flag (Excel-specific; observed values include `5` and `6`).
    ///
    /// When `rdrichvalue.xml` local-image metadata is missing we cannot recover this value. In
    /// that case we default to `0` (unknown) rather than guessing.
    ///
    /// The exact meaning of this field is not publicly documented; treat it as opaque metadata and
    /// preserve it when round-tripping.
    pub calc_origin: u32,
    /// Optional alternative text.
    pub alt_text: Option<String>,
    /// Optional hyperlink target attached to the same worksheet cell.
    pub hyperlink_target: Option<HyperlinkTarget>,
}

impl XlsxPackage {
    /// Extract embedded-in-cell images from the workbook package.
    ///
    /// Excel stores "in-cell" images by attaching RichData metadata to a worksheet cell via
    /// `c/@vm` (value-metadata index). In modern files this is resolved through `xl/metadata.xml`
    /// into a rich value index, which then ultimately references an image relationship slot in
    /// `xl/richData/richValueRel.xml`.
    ///
    /// This function is intentionally best-effort and supports multiple real-world variants:
    /// - Full RichData: `xl/metadata.xml` + `xl/richData/richValue.xml` / `rdrichvalue.xml`
    /// - Simplified: cells use `vm` to index directly into `richValueRel.xml` even when rich value
    ///   tables are missing.
    ///
    /// This API returns a mapping keyed by `(worksheet_part, cell_ref)`.
    pub fn extract_embedded_cell_images(
        &self,
    ) -> Result<HashMap<(String, CellRef), EmbeddedCellImage>, XlsxError> {
        let rich_value_rel_part = if self.part(RICH_VALUE_REL_PART).is_some() {
            Some(RICH_VALUE_REL_PART.to_string())
        } else {
            // Most workbooks use the canonical `xl/richData/richValueRel.xml` part name, but some
            // producers emit numbered variants (`richValueRel1.xml`, ...) or custom names.
            //
            // First attempt to find the numbered `richValueRel{N}.xml` pattern, then fall back to
            // any `xl/richData/*RichValueRel{N}.xml` part.
            find_lowest_numbered_part(self, "xl/richData/richValueRel", ".xml")
                .or_else(|| find_lowest_custom_rich_value_rel_part(self))
        };
        let Some(rich_value_rel_part) = rich_value_rel_part else {
            // Workbooks without in-cell images omit the entire `xl/richData/` tree.
            return Ok(HashMap::new());
        };
        let Some(rich_value_rel_bytes) = self.part(&rich_value_rel_part) else {
            return Ok(HashMap::new());
        };

        let rich_value_rel_xml = std::str::from_utf8(rich_value_rel_bytes)
            .map_err(|e| XlsxError::Invalid(format!("{rich_value_rel_part} not utf-8: {e}")))?;
        let rich_value_rel_ids = parse_rich_value_rel_ids(rich_value_rel_xml)?;
        if rich_value_rel_ids.is_empty() {
            return Ok(HashMap::new());
        }

        // Resolve relationship IDs (`rId*`) to concrete targets via the `.rels` part.
        let rich_value_rels_part = rels_for_part(&rich_value_rel_part);
        let Some(rich_value_rel_rels_bytes) = self.part(&rich_value_rels_part) else {
            // If the richValueRel part exists, we expect its .rels as well. Be defensive and
            // treat a missing rels part as "no images" rather than erroring.
            return Ok(HashMap::new());
        };

        let image_targets_by_rel_id =
            parse_rich_value_rel_image_targets(&rich_value_rel_part, rich_value_rel_rels_bytes)?;
        if image_targets_by_rel_id.is_empty() {
            return Ok(HashMap::new());
        }

        fn slot_points_to_image(
            rich_value_rel_ids: &[String],
            image_targets_by_rel_id: &HashMap<String, String>,
            slot: u32,
        ) -> bool {
            let Some(rel_id) = rich_value_rel_ids.get(slot as usize) else {
                return false;
            };
            image_targets_by_rel_id.contains_key(rel_id)
        }

        // Optional: value metadata mapping (worksheet `c/@vm` -> rich value index).
        //
        // Some simplified workbooks omit or do not populate `xl/metadata.xml`. In that case we
        // fall back to interpreting `vm` as a direct relationship-slot index.
        //
        // Some producers use numbered metadata part names like `xl/metadata1.xml`; prefer the
        // canonical `xl/metadata.xml` when present, but fall back to the lowest-numbered
        // `xl/metadata*.xml`.
        let metadata_part = if self.part("xl/metadata.xml").is_some() {
            Some("xl/metadata.xml".to_string())
        } else {
            find_lowest_numbered_part(self, "xl/metadata", ".xml")
        };
        let mut vm_to_rich_value: HashMap<u32, u32> = HashMap::new();
        if let Some(metadata_part) = metadata_part.as_deref() {
            if let Some(metadata_bytes) = self.part(metadata_part) {
                let parsed =
                    crate::rich_data::metadata::parse_value_metadata_vm_to_rich_value_index_map(
                        metadata_bytes,
                    )
                    .map_err(|e| {
                        XlsxError::Invalid(format!("failed to parse {metadata_part}: {e}"))
                    })?;
                vm_to_rich_value = parsed;
            }
        }
        let has_vm_mapping = !vm_to_rich_value.is_empty();

        // Best-effort worksheet discovery:
        //
        // `XlsxPackage::worksheet_parts()` requires `xl/_rels/workbook.xml.rels`. Some minimal or
        // malformed files omit it but still include usable worksheet parts under `xl/worksheets/*`.
        let mut worksheet_parts: Vec<String> = self
            .worksheet_parts()
            .ok()
            .map(|parts| parts.into_iter().map(|p| p.worksheet_part).collect())
            .unwrap_or_default();
        if worksheet_parts.is_empty() {
            worksheet_parts = self
                .part_names()
                .filter_map(|name| {
                    let name = name.strip_prefix('/').unwrap_or(name);
                    (name.starts_with("xl/worksheets/") && name.ends_with(".xml"))
                        .then_some(name.to_string())
                })
                .collect();
            worksheet_parts.sort();
            worksheet_parts.dedup();
        }

        // When we *don't* have `xl/metadata.xml`, we interpret `vm` as a relationship-slot index into
        // `xl/richData/richValueRel.xml`. Some producers encode this slot index as:
        // - 0-based (vm="0" for the first slot), or
        // - 1-based (vm="1" for the first slot).
        //
        // To avoid incorrectly biasing toward 1-based indexing in multi-image workbooks, choose the
        // preferred indexing scheme by scanning all `vm` cells and counting how many would resolve
        // to an image relationship under each scheme.
        let prefer_zero_based_vm_slots: bool = if has_vm_mapping {
            false
        } else {
            let mut zero_based_matches: usize = 0;
            let mut one_based_matches: usize = 0;
            for worksheet_part in &worksheet_parts {
                let Some(sheet_xml_bytes) = self.part(worksheet_part) else {
                    continue;
                };
                for (_cell_ref, vm) in parse_sheet_vm_image_cells(sheet_xml_bytes)? {
                    if slot_points_to_image(&rich_value_rel_ids, &image_targets_by_rel_id, vm) {
                        zero_based_matches = zero_based_matches.saturating_add(1);
                    }
                    if let Some(slot) = vm.checked_sub(1) {
                        if slot_points_to_image(&rich_value_rel_ids, &image_targets_by_rel_id, slot)
                        {
                            one_based_matches = one_based_matches.saturating_add(1);
                        }
                    }
                }
            }
            zero_based_matches > one_based_matches
        };

        // Optional: richValue*.xml relationship indices (rich value index -> relationship slot).
        //
        // Excel can split rich values across multiple parts (`richValue.xml`, `richValue1.xml`, ...).
        // Some producers use the pluralized `richValues*.xml` naming. `rich_value_part_suffix_index`
        // accepts both naming patterns; concatenate them in numeric-suffix order so that the
        // resulting vector index matches the global rich value index.
        let rich_value_rel_indices: Vec<Option<u32>> = {
            let mut rich_value_parts: Vec<(u32, String)> = self
                .part_names()
                .filter_map(|name| {
                    let name = name.strip_prefix('/').unwrap_or(name);
                    rich_value_part_suffix_index(name).map(|idx| (idx, name.to_string()))
                })
                .collect();

            if rich_value_parts.is_empty() {
                Vec::new()
            } else {
                rich_value_parts.sort_by(|(a_idx, a_part), (b_idx, b_part)| {
                    a_idx.cmp(b_idx).then_with(|| a_part.cmp(b_part))
                });
                rich_value_parts.dedup_by(|a, b| a.1 == b.1);

                let mut out: Vec<Option<u32>> = Vec::new();
                for (_idx, part) in rich_value_parts {
                    let Some(bytes) = self.part(&part) else {
                        continue;
                    };
                    out.extend(
                        crate::rich_data::rich_value::parse_rich_value_relationship_indices(bytes)?
                            .into_iter()
                            .map(|idx| idx.map(|idx| idx as u32)),
                    );
                }
                out
            }
        };

        // Optional: rdRichValue local-image metadata (rich value index -> relationship slot + alt/calcOrigin).
        let local_image_by_rich_value_index = match (
            self.part(RD_RICH_VALUE_PART),
            self.part(RD_RICH_VALUE_STRUCTURE_PART),
        ) {
            (Some(rd_rich_value_bytes), Some(rd_rich_value_structure_bytes)) => {
                if let Some(local_image_structure) =
                    parse_local_image_structure(rd_rich_value_structure_bytes)?
                {
                    parse_local_image_rich_values(rd_rich_value_bytes, &local_image_structure)?
                } else {
                    HashMap::new()
                }
            }
            _ => HashMap::new(),
        };

        let has_rich_value_tables =
            !local_image_by_rich_value_index.is_empty() || !rich_value_rel_indices.is_empty();

        // Best-effort fallback: some files include `xl/richData/richValueRel.xml` + rels but omit
        // `xl/metadata.xml` or the rich value tables. In that case, attempt to interpret `vm="N"`
        // as (0-based or 1-based) indices into `richValueRel.xml` and extract whatever image targets
        // we can resolve.
        let mut out = HashMap::new();
        for worksheet_part in worksheet_parts {
            let Some(sheet_xml_bytes) = self.part(&worksheet_part) else {
                continue;
            };
            let sheet_xml = std::str::from_utf8(sheet_xml_bytes)
                .map_err(|e| XlsxError::Invalid(format!("{worksheet_part} not utf-8: {e}")))?;

            let sheet_rels_part = rels_for_part(&worksheet_part);
            let sheet_rels_xml = self
                .part(&sheet_rels_part)
                .and_then(|bytes| std::str::from_utf8(bytes).ok());

            let hyperlinks =
                parse_worksheet_hyperlinks(sheet_xml, sheet_rels_xml).unwrap_or_default();

            let vm_cells = parse_sheet_vm_image_cells(sheet_xml_bytes)?;
            if vm_cells.is_empty() {
                continue;
            }

            // Excel typically stores worksheet `c/@vm` as a 1-based index into `xl/metadata.xml`
            // `<valueMetadata>` `<bk>` records, but some producers have been observed to use 0-based
            // indices instead.
            //
            // The metadata parser returns a canonical 1-based `vm -> richValue` mapping. To tolerate
            // 0-based worksheets we choose a per-worksheet offset (0 or 1) that yields the most
            // successfully-resolved images. If the result is ambiguous, we still treat `vm="0"` as
            // strong evidence of 0-based indexing.
            let vm_offset: u32 = if has_vm_mapping {
                let has_vm_zero = vm_cells.iter().any(|(_, vm)| *vm == 0);
                let mut resolved_offset_0 = 0usize;
                let mut resolved_offset_1 = 0usize;

                let resolves_image = |vm_raw: u32, offset: u32| -> bool {
                    let vm = vm_raw.saturating_add(offset);
                    let Some(&rich_value_index) = vm_to_rich_value.get(&vm) else {
                        return false;
                    };

                    let mut slot_candidates: Vec<u32> = Vec::new();
                    if let Some(local_image) =
                        local_image_by_rich_value_index.get(&rich_value_index)
                    {
                        slot_candidates.push(local_image.local_image_identifier);
                    } else if let Some(Some(rel_idx)) =
                        rich_value_rel_indices.get(rich_value_index as usize)
                    {
                        slot_candidates.push(*rel_idx);
                    } else if !has_rich_value_tables {
                        slot_candidates.push(rich_value_index);
                    } else {
                        return false;
                    }

                    for slot in slot_candidates {
                        let Some(rel_id) = rich_value_rel_ids.get(slot as usize) else {
                            continue;
                        };
                        let Some(target_part) = image_targets_by_rel_id.get(rel_id) else {
                            continue;
                        };
                        if self.part(target_part).is_some() {
                            return true;
                        }
                    }

                    false
                };

                for (_cell_ref, vm) in &vm_cells {
                    if resolves_image(*vm, 0) {
                        resolved_offset_0 += 1;
                    }
                    if resolves_image(*vm, 1) {
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
            } else {
                0
            };

            for (cell_ref, vm) in vm_cells {
                // First resolve the cell's `vm` into a rich value index when possible.
                let rich_value_index = if has_vm_mapping {
                    let vm = vm.saturating_add(vm_offset);
                    let Some(&idx) = vm_to_rich_value.get(&vm) else {
                        continue;
                    };
                    Some(idx)
                } else {
                    None
                };

                // The `CalcOrigin` value lives in the rdRichValue local-image metadata. When that
                // schema is missing, we cannot recover it; default to `0` (unknown) so callers can
                // distinguish real CalcOrigin values (observed: 5/6) from an inferred fallback.
                let mut calc_origin: u32 = 0;
                let mut alt_text: Option<String> = None;

                // Determine which relationship-slot index to use for this cell image.
                // We try, in order:
                // 1) rdRichValue local image schema (best; includes CalcOrigin + alt text).
                // 2) richValue.xml relationship index (`<v kind="rel">`).
                // 3) direct indexing: interpret rich value index as relationship slot (only when rich value tables are missing).
                // 4) last-ditch: interpret `vm` as the relationship-slot index (tolerating 1-based vs 0-based).
                let mut slot_candidates: Vec<u32> = Vec::new();
                if let Some(rich_value_index) = rich_value_index {
                    if let Some(local_image) =
                        local_image_by_rich_value_index.get(&rich_value_index)
                    {
                        slot_candidates.push(local_image.local_image_identifier);
                        calc_origin = local_image.calc_origin;
                        alt_text = local_image.alt_text.clone();
                    } else if let Some(Some(rel_idx)) =
                        rich_value_rel_indices.get(rich_value_index as usize)
                    {
                        slot_candidates.push(*rel_idx);
                    } else if !has_rich_value_tables {
                        slot_candidates.push(rich_value_index);
                    } else {
                        // We have rich value tables but couldn't map this rich value to a relationship
                        // slot; treat it as a non-image rich value.
                        continue;
                    }
                } else {
                    // No metadata mapping; fall back to interpreting `vm` as a relationship slot.
                    if prefer_zero_based_vm_slots {
                        slot_candidates.push(vm);
                        if vm > 0 {
                            slot_candidates.push(vm - 1);
                        }
                    } else {
                        if vm > 0 {
                            slot_candidates.push(vm - 1);
                        }
                        slot_candidates.push(vm);
                    }
                }

                // Resolve the first slot candidate that maps to a concrete image part.
                let mut image_part: Option<String> = None;
                for slot in slot_candidates {
                    if let Some(part) = resolve_local_image_identifier_to_image_part(
                        &rich_value_rel_ids,
                        &image_targets_by_rel_id,
                        slot,
                    ) {
                        image_part = Some(part);
                        break;
                    }
                }

                let Some(image_part) = image_part else {
                    continue;
                };
                let Some(image_bytes) = self.part(&image_part) else {
                    continue;
                };

                let hyperlink_target = hyperlinks
                    .iter()
                    .find(|link| link.range.contains(cell_ref))
                    .map(|link| link.target.clone());

                out.insert(
                    (worksheet_part.clone(), cell_ref),
                    EmbeddedCellImage {
                        image_part,
                        image_bytes: image_bytes.to_vec(),
                        calc_origin,
                        alt_text,
                        hyperlink_target,
                    },
                );
            }
        }

        Ok(out)
    }
}

fn rich_value_part_suffix_index(part_path: &str) -> Option<u32> {
    let part_path = part_path.strip_prefix('/').unwrap_or(part_path);
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

fn find_lowest_numbered_part(pkg: &XlsxPackage, prefix: &str, suffix: &str) -> Option<String> {
    let mut best: Option<(u32, String)> = None;

    for part in pkg.part_names() {
        let part = part.strip_prefix('/').unwrap_or(part);
        let Some(num) = numeric_suffix(part, prefix, suffix) else {
            continue;
        };

        match &mut best {
            Some((best_num, best_name)) => {
                if num < *best_num || (num == *best_num && part < best_name.as_str()) {
                    *best_num = num;
                    *best_name = part.to_string();
                }
            }
            None => best = Some((num, part.to_string())),
        }
    }

    best.map(|(_, name)| name)
}

fn find_lowest_custom_rich_value_rel_part(pkg: &XlsxPackage) -> Option<String> {
    let mut best: Option<(u32, String)> = None;

    for part in pkg.part_names() {
        let part = part.strip_prefix('/').unwrap_or(part);
        let Some(num) = rich_value_rel_custom_suffix_index(part) else {
            continue;
        };

        match &mut best {
            Some((best_num, best_name)) => {
                if num < *best_num || (num == *best_num && part < best_name.as_str()) {
                    *best_num = num;
                    *best_name = part.to_string();
                }
            }
            None => best = Some((num, part.to_string())),
        }
    }

    best.map(|(_, name)| name)
}

fn rich_value_rel_custom_suffix_index(part_path: &str) -> Option<u32> {
    let part_path = part_path.strip_prefix('/').unwrap_or(part_path);
    if !part_path.starts_with("xl/richData/") {
        return None;
    }

    let file_name = part_path.rsplit('/').next()?;
    let stem = crate::ascii::strip_suffix_ignore_case(file_name, ".xml")?;
    let idx = crate::ascii::rfind_ignore_case(stem, "richvaluerel")?;
    let suffix = &stem[idx + "richvaluerel".len()..];
    if !suffix.chars().all(|c| c.is_ascii_digit()) {
        return None;
    }

    if suffix.is_empty() {
        Some(0)
    } else {
        suffix.parse::<u32>().ok()
    }
}

fn numeric_suffix(part_name: &str, prefix: &str, suffix: &str) -> Option<u32> {
    let mid = part_name.strip_prefix(prefix)?.strip_suffix(suffix)?;
    if !mid.chars().all(|c| c.is_ascii_digit()) {
        return None;
    }
    if mid.is_empty() {
        return Some(0);
    }
    mid.parse::<u32>().ok()
}

fn parse_rich_value_rel_ids(xml: &str) -> Result<Vec<String>, XlsxError> {
    let mut reader = Reader::from_reader(Cursor::new(xml.as_bytes()));
    reader.config_mut().trim_text(true);

    let mut buf = Vec::new();
    let mut out = Vec::new();

    loop {
        match reader.read_event_into(&mut buf)? {
            Event::Eof => break,
            Event::Start(e) | Event::Empty(e) if e.local_name().as_ref() == b"rel" => {
                let mut rid_prefixed: Option<String> = None;
                let mut rid_unprefixed: Option<String> = None;
                for attr in e.attributes() {
                    let attr = attr?;
                    let key = attr.key.as_ref();
                    let local = key.rsplit(|b| *b == b':').next().unwrap_or(key);
                    if !local.eq_ignore_ascii_case(b"id") {
                        continue;
                    }
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
                // Preserve missing/invalid entries as placeholders so indices remain stable.
                out.push(rid_prefixed.or(rid_unprefixed).unwrap_or_default());
            }
            _ => {}
        }
        buf.clear();
    }

    Ok(out)
}

fn parse_rich_value_rel_image_targets(
    source_part: &str,
    rels_xml: &[u8],
) -> Result<HashMap<String, String>, XlsxError> {
    let relationships = crate::openxml::parse_relationships(rels_xml)?;
    let mut out = HashMap::new();
    for rel in relationships {
        let rel_id = rel.id.trim();
        if rel_id.is_empty() {
            continue;
        }
        if rel
            .target_mode
            .as_deref()
            .is_some_and(|m| m.trim().eq_ignore_ascii_case("External"))
        {
            continue;
        }
        if rel.type_uri.trim() != REL_TYPE_IMAGE {
            continue;
        }
        let target_raw = strip_uri_suffixes(&rel.target);
        if target_raw.is_empty() {
            continue;
        }
        // Be resilient to invalid/unescaped Windows-style path separators so we can match the
        // nonstandard-but-observed `media/` and `xl/` prefixes consistently.
        let target_cow: std::borrow::Cow<'_, str> = if target_raw.contains('\\') {
            std::borrow::Cow::Owned(target_raw.replace('\\', "/"))
        } else {
            std::borrow::Cow::Borrowed(target_raw)
        };
        let target = target_cow.as_ref().trim_start_matches("./");
        // Some producers emit `Target="media/image1.png"` (relative to `xl/`) rather than the more
        // common `Target="../media/image1.png"` (relative to `xl/richData/`). Make a best-effort
        // guess for this case.
        let resolved = if target.starts_with("media/") {
            // Resolve relative to `xl/` rather than `xl/richData/`, while still applying URI
            // normalization (fragment/query stripping and dot-segment resolution).
            resolve_target("xl/workbook.xml", target)
        } else if target.starts_with("xl/") {
            // Treat `xl/...` targets as absolute part names (some producers omit the leading `/`).
            let mut absolute = String::with_capacity(target.len() + 1);
            absolute.push('/');
            absolute.push_str(target);
            resolve_target(source_part, &absolute)
        } else {
            resolve_target(source_part, target)
        };
        out.insert(rel_id.to_string(), resolved);
    }
    Ok(out)
}

fn strip_uri_suffixes(target: &str) -> &str {
    let target = target.trim();
    let target = target
        .split_once('#')
        .map(|(base, _)| base)
        .unwrap_or(target);
    target
        .split_once('?')
        .map(|(base, _)| base)
        .unwrap_or(target)
}

fn parse_sheet_vm_image_cells(sheet_xml: &[u8]) -> Result<Vec<(CellRef, u32)>, XlsxError> {
    let cells_with_metadata =
        scan_cells_with_metadata_indices(sheet_xml).map_err(|err| match err {
            RichDataError::Xlsx(err) => err,
            RichDataError::XmlNonUtf8 { source, .. } => {
                XlsxError::Invalid(format!("worksheet xml contains invalid UTF-8: {source}"))
            }
            RichDataError::XmlParse { source, .. } => XlsxError::RoXml(source),
        })?;

    // Filter to `vm`-annotated cells; the rich value structure parsing in the caller will determine
    // which `vm`s are actually local images.
    Ok(cells_with_metadata
        .into_iter()
        .filter_map(|(cell, vm, _cm)| vm.map(|vm| (cell, vm)))
        .collect())
}

fn resolve_local_image_identifier_to_image_part(
    rich_value_rel_ids: &[String],
    image_targets_by_rel_id: &HashMap<String, String>,
    local_image_identifier: u32,
) -> Option<String> {
    let rel_id =
        crate::rich_data::rel_slot_get(rich_value_rel_ids, local_image_identifier as usize)?;
    image_targets_by_rel_id.get(rel_id).cloned()
}

#[derive(Debug, Clone)]
struct LocalImageStructure {
    struct_index: u32,
    local_image_identifier_pos: usize,
    calc_origin_pos: usize,
    alt_text_pos: Option<usize>,
}

#[derive(Debug, Clone)]
struct LocalImageRichValueRow {
    local_image_identifier: u32,
    calc_origin: u32,
    alt_text: Option<String>,
}

fn parse_local_image_structure(xml: &[u8]) -> Result<Option<LocalImageStructure>, XlsxError> {
    let mut reader = Reader::from_reader(Cursor::new(xml));
    reader.config_mut().trim_text(true);
    let mut buf = Vec::new();

    let mut struct_index: u32 = 0;
    let mut in_target_struct = false;
    let mut key_pos: usize = 0;

    let mut target_struct_index: Option<u32> = None;
    let mut local_image_identifier_pos: Option<usize> = None;
    let mut calc_origin_pos: Option<usize> = None;
    let mut alt_text_pos: Option<usize> = None;

    loop {
        match reader.read_event_into(&mut buf)? {
            Event::Start(e) => {
                if e.local_name().as_ref() == b"s" {
                    let t = attr_value(&e, b"t")?;
                    in_target_struct = t.as_deref() == Some("_localImage");
                    if in_target_struct {
                        target_struct_index = Some(struct_index);
                        key_pos = 0;
                        local_image_identifier_pos = None;
                        calc_origin_pos = None;
                        alt_text_pos = None;
                    }
                } else if in_target_struct && e.local_name().as_ref() == b"k" {
                    if let Some(name) = attr_value(&e, b"n")? {
                        match name.as_str() {
                            "_rvRel:LocalImageIdentifier" => {
                                local_image_identifier_pos = Some(key_pos)
                            }
                            "CalcOrigin" => calc_origin_pos = Some(key_pos),
                            "Text" => alt_text_pos = Some(key_pos),
                            _ => {}
                        }
                    }
                    key_pos += 1;
                    reader.read_to_end_into(e.name(), &mut Vec::new())?;
                }
            }
            Event::Empty(e) => {
                if e.local_name().as_ref() == b"s" {
                    // Empty struct definition.
                    struct_index = struct_index.saturating_add(1);
                } else if in_target_struct && e.local_name().as_ref() == b"k" {
                    if let Some(name) = attr_value(&e, b"n")? {
                        match name.as_str() {
                            "_rvRel:LocalImageIdentifier" => {
                                local_image_identifier_pos = Some(key_pos)
                            }
                            "CalcOrigin" => calc_origin_pos = Some(key_pos),
                            "Text" => alt_text_pos = Some(key_pos),
                            _ => {}
                        }
                    }
                    key_pos += 1;
                }
            }
            Event::End(e) => {
                if e.local_name().as_ref() == b"s" {
                    if in_target_struct {
                        break;
                    }
                    struct_index = struct_index.saturating_add(1);
                }
            }
            Event::Eof => break,
            _ => {}
        }
        buf.clear();
    }

    let Some(struct_index) = target_struct_index else {
        return Ok(None);
    };
    let Some(local_image_identifier_pos) = local_image_identifier_pos else {
        return Ok(None);
    };
    let Some(calc_origin_pos) = calc_origin_pos else {
        return Ok(None);
    };

    Ok(Some(LocalImageStructure {
        struct_index,
        local_image_identifier_pos,
        calc_origin_pos,
        alt_text_pos,
    }))
}

fn parse_local_image_rich_values(
    xml: &[u8],
    structure: &LocalImageStructure,
) -> Result<HashMap<u32, LocalImageRichValueRow>, XlsxError> {
    let mut reader = Reader::from_reader(Cursor::new(xml));
    reader.config_mut().trim_text(false);
    let mut buf = Vec::new();

    let mut out = HashMap::new();
    let mut rv_index: u32 = 0;

    loop {
        match reader.read_event_into(&mut buf)? {
            Event::Start(e) => {
                if e.local_name().as_ref() != b"rv" {
                    continue;
                }
                let s = attr_value(&e, b"s")?;
                let Some(struct_idx) = s.as_deref().and_then(|s| s.parse::<u32>().ok()) else {
                    reader.read_to_end_into(e.name(), &mut Vec::new())?;
                    rv_index = rv_index.saturating_add(1);
                    continue;
                };

                if struct_idx != structure.struct_index {
                    reader.read_to_end_into(e.name(), &mut Vec::new())?;
                    rv_index = rv_index.saturating_add(1);
                    continue;
                }

                let values = read_rv_values(&mut reader)?;
                if let Some(row) = local_image_row_from_values(&values, structure) {
                    out.insert(rv_index, row);
                }
                rv_index = rv_index.saturating_add(1);
            }
            Event::Empty(e) => {
                if e.local_name().as_ref() == b"rv" {
                    rv_index = rv_index.saturating_add(1);
                }
            }
            Event::Eof => break,
            _ => {}
        }
        buf.clear();
    }

    Ok(out)
}

fn read_rv_values(reader: &mut Reader<Cursor<&[u8]>>) -> Result<Vec<String>, XlsxError> {
    let mut buf = Vec::new();
    let mut values = Vec::new();
    loop {
        match reader.read_event_into(&mut buf)? {
            Event::Start(e) => {
                if e.local_name().as_ref() == b"v" {
                    values.push(read_text(reader, QName(b"v"))?);
                } else {
                    reader.read_to_end_into(e.name(), &mut Vec::new())?;
                }
            }
            Event::Empty(e) => {
                if e.local_name().as_ref() == b"v" {
                    values.push(String::new());
                }
            }
            Event::End(e) => {
                if e.local_name().as_ref() == b"rv" {
                    break;
                }
            }
            Event::Eof => return Err(XlsxError::Invalid("unexpected eof in <rv>".to_string())),
            _ => {}
        }
        buf.clear();
    }
    Ok(values)
}

fn local_image_row_from_values(
    values: &[String],
    structure: &LocalImageStructure,
) -> Option<LocalImageRichValueRow> {
    let local_image_identifier = values
        .get(structure.local_image_identifier_pos)?
        .trim()
        .parse::<u32>()
        .ok()?;
    let calc_origin = values
        .get(structure.calc_origin_pos)?
        .trim()
        .parse::<u32>()
        .ok()?;
    let alt_text = structure
        .alt_text_pos
        .and_then(|idx| values.get(idx))
        .map(|s| s.to_string())
        .filter(|s| !s.is_empty());

    Some(LocalImageRichValueRow {
        local_image_identifier,
        calc_origin,
        alt_text,
    })
}

fn read_text(reader: &mut Reader<Cursor<&[u8]>>, end: QName<'_>) -> Result<String, XlsxError> {
    let mut buf = Vec::new();
    let mut text = String::new();
    loop {
        match reader.read_event_into(&mut buf)? {
            Event::Text(e) => {
                text.push_str(&e.unescape()?.into_owned());
            }
            Event::CData(e) => {
                let s = std::str::from_utf8(e.as_ref())
                    .map_err(|err| XlsxError::Invalid(format!("invalid utf-8 in cdata: {err}")))?;
                text.push_str(s);
            }
            Event::End(e) if e.name() == end => break,
            Event::Eof => return Err(XlsxError::Invalid("unexpected eof".to_string())),
            _ => {}
        }
        buf.clear();
    }
    Ok(text)
}

fn attr_value(
    e: &quick_xml::events::BytesStart<'_>,
    key: &[u8],
) -> Result<Option<String>, XlsxError> {
    for attr in e.attributes().with_checks(false) {
        let attr = attr?;
        if attr.key.as_ref() == key {
            return Ok(Some(attr.unescape_value()?.into_owned()));
        }
    }
    Ok(None)
}
