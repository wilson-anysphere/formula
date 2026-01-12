use std::collections::HashMap;
use std::io::Cursor;

use formula_model::{CellRef, HyperlinkTarget};
use quick_xml::events::Event;
use quick_xml::name::QName;
use quick_xml::Reader;

use crate::path::{rels_for_part, resolve_target};
use crate::rich_data::{scan_cells_with_metadata_indices, RichDataError};
use crate::{parse_worksheet_hyperlinks, XlsxError, XlsxPackage};

const RICH_VALUE_REL_PART: &str = "xl/richData/richValueRel.xml";
const RD_RICH_VALUE_PART: &str = "xl/richData/rdrichvalue.xml";
const RD_RICH_VALUE_STRUCTURE_PART: &str = "xl/richData/rdrichvaluestructure.xml";
const REL_TYPE_IMAGE: &str =
    "http://schemas.openxmlformats.org/officeDocument/2006/relationships/image";

/// An embedded image stored inside a cell using Excel's RichData / `vm=` mechanism.
///
/// These are distinct from DrawingML images (floating/anchored shapes).
#[derive(Debug, Clone, PartialEq, Eq)]
pub struct EmbeddedCellImage {
    /// Resolved package part name for the image (e.g. `xl/media/image1.png`).
    pub image_part: String,
    /// Raw bytes for the image file.
    pub image_bytes: Vec<u8>,
    /// Rich value CalcOrigin (observed values: `5` normal, `6` decorative).
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
        let Some(rich_value_rel_bytes) = self.part(RICH_VALUE_REL_PART) else {
            // Workbooks without in-cell images omit the entire `xl/richData/` tree.
            return Ok(HashMap::new());
        };

        let rich_value_rel_xml = std::str::from_utf8(rich_value_rel_bytes)
            .map_err(|e| XlsxError::Invalid(format!("{RICH_VALUE_REL_PART} not utf-8: {e}")))?;
        let rich_value_rel_ids = parse_rich_value_rel_ids(rich_value_rel_xml)?;
        if rich_value_rel_ids.is_empty() {
            return Ok(HashMap::new());
        }

        // Resolve relationship IDs (`rId*`) to concrete targets via the `.rels` part.
        let rich_value_rels_part = rels_for_part(RICH_VALUE_REL_PART);
        let Some(rich_value_rel_rels_bytes) = self.part(&rich_value_rels_part) else {
            // If the richValueRel part exists, we expect its .rels as well. Be defensive and
            // treat a missing rels part as "no images" rather than erroring.
            return Ok(HashMap::new());
        };

        let image_targets_by_rel_id =
            parse_rich_value_rel_image_targets(rich_value_rel_rels_bytes)?;
        if image_targets_by_rel_id.is_empty() {
            return Ok(HashMap::new());
        }

        // Optional: value metadata mapping (worksheet `c/@vm` -> rich value index).
        //
        // Some simplified workbooks omit or do not populate `xl/metadata.xml`. In that case we
        // fall back to interpreting `vm` as a direct relationship-slot index.
        let mut vm_to_rich_value: HashMap<u32, u32> = HashMap::new();
        if let Some(metadata_bytes) = self.part("xl/metadata.xml") {
            let parsed = crate::rich_data::metadata::parse_value_metadata_vm_to_rich_value_index_map(
                metadata_bytes,
            )
            .map_err(|e| XlsxError::Invalid(format!("failed to parse xl/metadata.xml: {e}")))?;
            for (vm, rv_idx) in parsed {
                vm_to_rich_value.entry(vm).or_insert(rv_idx);
                // Excel has been observed to encode `vm` as both 0-based and 1-based in the wild.
                // Insert both so callers don't need to care.
                if vm > 0 {
                    vm_to_rich_value.entry(vm - 1).or_insert(rv_idx);
                }
            }
        }
        let has_vm_mapping = !vm_to_rich_value.is_empty();

        // Optional: richValue.xml relationship indices (rich value index -> relationship slot).
        let rich_value_rel_indices: Vec<Option<u32>> = match self.part("xl/richData/richValue.xml") {
            Some(bytes) => crate::rich_data::rich_value::parse_rich_value_relationship_indices(bytes)?
                .into_iter()
                .map(|idx| idx.map(|idx| idx as u32))
                .collect(),
            None => Vec::new(),
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

        let mut out = HashMap::new();
        for sheet in self.worksheet_parts()? {
            let Some(sheet_xml_bytes) = self.part(&sheet.worksheet_part) else {
                continue;
            };
            let sheet_xml = std::str::from_utf8(sheet_xml_bytes).map_err(|e| {
                XlsxError::Invalid(format!("{} not utf-8: {e}", sheet.worksheet_part))
            })?;

            let sheet_rels_part = rels_for_part(&sheet.worksheet_part);
            let sheet_rels_xml = self
                .part(&sheet_rels_part)
                .and_then(|bytes| std::str::from_utf8(bytes).ok());

            // Best-effort: if hyperlink parsing fails (malformed file), still extract images.
            let hyperlinks =
                parse_worksheet_hyperlinks(sheet_xml, sheet_rels_xml).unwrap_or_default();

            for (cell_ref, vm) in parse_sheet_vm_image_cells(sheet_xml_bytes)? {
                // First resolve the cell's `vm` into a rich value index when possible.
                let rich_value_index = if has_vm_mapping {
                    let Some(&idx) = vm_to_rich_value.get(&vm) else {
                        continue;
                    };
                    Some(idx)
                } else {
                    None
                };

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
                    if let Some(local_image) = local_image_by_rich_value_index.get(&rich_value_index)
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
                    if vm > 0 {
                        slot_candidates.push(vm - 1);
                    }
                    slot_candidates.push(vm);
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
                    (sheet.worksheet_part.clone(), cell_ref),
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

fn parse_rich_value_rel_ids(xml: &str) -> Result<Vec<String>, XlsxError> {
    let mut reader = Reader::from_reader(Cursor::new(xml.as_bytes()));
    reader.config_mut().trim_text(true);

    let mut buf = Vec::new();
    let mut out = Vec::new();

    loop {
        match reader.read_event_into(&mut buf)? {
            Event::Eof => break,
            Event::Start(e) | Event::Empty(e) if e.local_name().as_ref() == b"rel" => {
                for attr in e.attributes() {
                    let attr = attr?;
                    match attr.key.as_ref() {
                        b"r:id" | b"id" => {
                            out.push(attr.unescape_value()?.into_owned());
                            break;
                        }
                        _ => {}
                    }
                }
            }
            _ => {}
        }
        buf.clear();
    }

    Ok(out)
}

fn parse_rich_value_rel_image_targets(rels_xml: &[u8]) -> Result<HashMap<String, String>, XlsxError> {
    let relationships = crate::openxml::parse_relationships(rels_xml)?;
    let mut out = HashMap::new();
    for rel in relationships {
        if rel
            .target_mode
            .as_deref()
            .is_some_and(|m| m.trim().eq_ignore_ascii_case("External"))
        {
            continue;
        }
        if rel.type_uri != REL_TYPE_IMAGE {
            continue;
        }
        let target = strip_fragment(&rel.target);
        if target.is_empty() {
            continue;
        }
        let resolved = resolve_target(RICH_VALUE_REL_PART, target);
        out.insert(rel.id, resolved);
    }
    Ok(out)
}

fn strip_fragment(target: &str) -> &str {
    target
        .split_once('#')
        .map(|(base, _)| base)
        .unwrap_or(target)
}

fn parse_sheet_vm_image_cells(sheet_xml: &[u8]) -> Result<Vec<(CellRef, u32)>, XlsxError> {
    let cells_with_metadata =
        scan_cells_with_metadata_indices(sheet_xml).map_err(|err| match err {
            RichDataError::Xlsx(err) => err,
            RichDataError::XmlNonUtf8 { source, .. } => XlsxError::Invalid(format!(
                "worksheet xml contains invalid UTF-8: {source}"
            )),
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
    let rel_id = rich_value_rel_ids.get(local_image_identifier as usize)?;
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

fn attr_value(e: &quick_xml::events::BytesStart<'_>, key: &[u8]) -> Result<Option<String>, XlsxError> {
    for attr in e.attributes().with_checks(false) {
        let attr = attr?;
        if attr.key.as_ref() == key {
            return Ok(Some(attr.unescape_value()?.into_owned()));
        }
    }
    Ok(None)
}
