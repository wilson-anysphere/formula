use std::collections::{BTreeMap, HashMap, HashSet};
use std::io::{BufRead, BufReader, Read, Seek, Write};

use formula_model::rich_text::RichText;
use formula_model::{CellRef, CellValue, StyleTable};
use quick_xml::events::{BytesEnd, BytesStart, BytesText, Event};
use quick_xml::{Reader, Writer};
use thiserror::Error;
use zip::write::FileOptions;
use zip::{CompressionMethod, ZipArchive, ZipWriter};

use crate::openxml::{parse_relationships, resolve_target};
use crate::recalc_policy::{
    content_types_remove_calc_chain, workbook_rels_remove_calc_chain,
    workbook_xml_force_full_calc_on_load, RecalcPolicyError,
};
use crate::shared_strings::preserve::SharedStringsEditor;
use crate::styles::XlsxStylesEditor;
use crate::{parse_workbook_sheets, CellPatch, WorkbookCellPatches};
use crate::RecalcPolicy;

const REL_TYPE_STYLES: &str =
    "http://schemas.openxmlformats.org/officeDocument/2006/relationships/styles";
const REL_TYPE_SHARED_STRINGS: &str =
    "http://schemas.openxmlformats.org/officeDocument/2006/relationships/sharedStrings";

const SPREADSHEETML_NS: &str = "http://schemas.openxmlformats.org/spreadsheetml/2006/main";

#[derive(Debug, Error)]
pub enum StreamingPatchError {
    #[error("io error: {0}")]
    Io(#[from] std::io::Error),
    #[error("zip error: {0}")]
    Zip(#[from] zip::result::ZipError),
    #[error("xml error: {0}")]
    Xml(#[from] quick_xml::Error),
    #[error("xml attribute error: {0}")]
    XmlAttr(#[from] quick_xml::events::attributes::AttrError),
    #[error("unsupported cell value kind for streaming patch: {0:?}")]
    UnsupportedCellValue(CellValue),
    #[error("invalid cell reference in worksheet xml: {0}")]
    InvalidCellRef(String),
    #[error("worksheet part referenced by patch not found in input zip: {0}")]
    MissingWorksheetPart(String),
    #[error("worksheet xml is missing required <sheetData> section: {0}")]
    MissingSheetData(String),
    #[error("xlsx error: {0}")]
    Xlsx(#[from] crate::XlsxError),
}

impl From<RecalcPolicyError> for StreamingPatchError {
    fn from(value: RecalcPolicyError) -> Self {
        match value {
            RecalcPolicyError::Io(err) => StreamingPatchError::Io(err),
            RecalcPolicyError::Xml(err) => StreamingPatchError::Xml(err),
            RecalcPolicyError::XmlAttr(err) => StreamingPatchError::XmlAttr(err),
        }
    }
}

/// Patch a single cell in a worksheet part (`xl/worksheets/sheetN.xml`).
///
/// The patch is applied by rewriting only the worksheet XML part. All other ZIP entries are
/// copied from the input XLSX/XLSM streamingly.
#[derive(Debug, Clone)]
pub struct WorksheetCellPatch {
    /// ZIP entry name for the worksheet XML (e.g. `xl/worksheets/sheet1.xml`).
    pub worksheet_part: String,
    /// Cell reference to patch.
    pub cell: CellRef,
    /// New cached value for the cell.
    pub value: CellValue,
    /// Optional formula to write into the `<f>` element. Leading `=` is permitted.
    pub formula: Option<String>,
    /// Optional `xf` index to write into the cell `s` attribute.
    ///
    /// - `None`: preserve the existing `s` attribute when patching an existing cell (and omit `s`
    ///   entirely when inserting a new cell).
    /// - `Some(0)`: remove the `s` attribute.
    /// - `Some(xf)` where `xf != 0`: set/overwrite `s="xf"`.
    pub xf_index: Option<u32>,
}

impl WorksheetCellPatch {
    pub fn new(
        worksheet_part: impl Into<String>,
        cell: CellRef,
        value: CellValue,
        formula: Option<String>,
    ) -> Self {
        Self {
            worksheet_part: worksheet_part.into(),
            cell,
            value,
            formula,
            xf_index: None,
        }
    }

    pub fn with_xf_index(mut self, xf_index: Option<u32>) -> Self {
        self.xf_index = xf_index;
        self
    }
}

/// Streaming XLSX/XLSM patcher.
///
/// This function reads the input ZIP once and writes a new ZIP by stream-copying all unchanged
/// parts. Only worksheet XML parts mentioned in `cell_patches` are rewritten.
pub fn patch_xlsx_streaming<R: Read + Seek, W: Write + Seek>(
    input: R,
    output: W,
    cell_patches: &[WorksheetCellPatch],
) -> Result<(), StreamingPatchError> {
    let mut patches_by_part: HashMap<String, Vec<WorksheetCellPatch>> = HashMap::new();
    for patch in cell_patches {
        patches_by_part
            .entry(patch.worksheet_part.clone())
            .or_default()
            .push(patch.clone());
    }
    // Deterministic patching within a worksheet.
    for patches in patches_by_part.values_mut() {
        patches.sort_by_key(|p| (p.cell.row, p.cell.col));
    }

    let mut archive = ZipArchive::new(input)?;
    let mut formula_changed = cell_patches.iter().any(|p| p.formula.is_some());
    if !formula_changed {
        formula_changed = streaming_patches_remove_existing_formulas(&mut archive, &patches_by_part)?;
    }
    let recalc_policy = if formula_changed {
        // Match `XlsxPackage::apply_cell_patches` default: dropping calcChain and requesting a full
        // calc on load is the safest behavior after formula edits (including removing formulas).
        RecalcPolicy::default()
    } else {
        RecalcPolicy::PRESERVE
    };
    patch_xlsx_streaming_with_archive(
        &mut archive,
        output,
        &patches_by_part,
        &HashMap::new(),
        &HashMap::new(),
        recalc_policy,
    )?;
    Ok(())
}

/// Apply [`WorkbookCellPatches`] (the part-preserving cell patch DSL) using the streaming ZIP
/// rewriter.
///
/// Patches are keyed by a worksheet selector, matching [`XlsxPackage::apply_cell_patches`] but
/// without loading every part into memory.
///
/// Supported worksheet selectors:
/// - Worksheet (tab) name (case-insensitive, as in Excel)
/// - Worksheet part name (any key containing `/`, e.g. `xl/worksheets/sheet2.xml`)
/// - Workbook relationship id (e.g. `rId2`) when no sheet name matches
pub fn patch_xlsx_streaming_workbook_cell_patches<R: Read + Seek, W: Write + Seek>(
    input: R,
    output: W,
    patches: &WorkbookCellPatches,
) -> Result<(), StreamingPatchError> {
    if patches.is_empty() {
        return patch_xlsx_streaming(input, output, &[]);
    }

    // StyleId patches require rewriting styles.xml; callers should use the style-aware variant.
    for (_, sheet_patches) in patches.sheets() {
        for (_, patch) in sheet_patches.iter() {
            if patch.style_id().is_some_and(|id| id != 0) {
                return Err(crate::XlsxError::Invalid(
                    "style_id patches require patch_xlsx_streaming_workbook_cell_patches_with_styles"
                        .to_string(),
                )
                .into());
            }
        }
    }

    let mut archive = ZipArchive::new(input)?;

    let mut pre_read_parts: HashMap<String, Vec<u8>> = HashMap::new();
    let workbook_xml = read_zip_part(&mut archive, "xl/workbook.xml", &mut pre_read_parts)?;
    let workbook_xml = String::from_utf8(workbook_xml).map_err(crate::XlsxError::from)?;
    let workbook_sheets = parse_workbook_sheets(&workbook_xml)?;

    let workbook_rels_bytes = read_zip_part(
        &mut archive,
        "xl/_rels/workbook.xml.rels",
        &mut pre_read_parts,
    )?;
    let rels = parse_relationships(&workbook_rels_bytes)?;
    let mut rel_targets: HashMap<String, String> = HashMap::new();
    for rel in rels {
        rel_targets.insert(rel.id, resolve_target("xl/workbook.xml", &rel.target));
    }

    let mut patches_by_part: HashMap<String, Vec<WorksheetCellPatch>> = HashMap::new();
    let mut saw_formula_patch = false;
    for (sheet_selector, sheet_patches) in patches.sheets() {
        if sheet_patches.is_empty() {
            continue;
        }
        let worksheet_part = resolve_worksheet_part_for_selector(
            sheet_selector,
            &workbook_sheets,
            &rel_targets,
        )?;

        for (cell_ref, patch) in sheet_patches.iter() {
            let (value, formula) = match patch {
                CellPatch::Clear { .. } => (CellValue::Empty, None),
                CellPatch::Set { value, formula, .. } => (value.clone(), formula.clone()),
            };
            saw_formula_patch |= formula.is_some();
            let xf_index = patch.style_index();
            patches_by_part
                .entry(worksheet_part.clone())
                .or_default()
                .push(
                    WorksheetCellPatch::new(worksheet_part.clone(), cell_ref, value, formula)
                        .with_xf_index(xf_index),
                );
        }
    }

    for patches in patches_by_part.values_mut() {
        patches.sort_by_key(|p| (p.cell.row, p.cell.col));
    }

    let mut formula_changed = saw_formula_patch;
    if !formula_changed {
        formula_changed = streaming_patches_remove_existing_formulas(&mut archive, &patches_by_part)?;
    }
    let recalc_policy = if formula_changed {
        RecalcPolicy::default()
    } else {
        RecalcPolicy::PRESERVE
    };

    patch_xlsx_streaming_with_archive(
        &mut archive,
        output,
        &patches_by_part,
        &pre_read_parts,
        &HashMap::new(),
        recalc_policy,
    )?;
    Ok(())
}

/// Apply [`WorkbookCellPatches`] using the streaming ZIP rewriter, resolving `style_id` overrides
/// via `styles.xml`.
///
/// This variant updates `styles.xml` deterministically when new styles are introduced.
pub fn patch_xlsx_streaming_workbook_cell_patches_with_styles<R: Read + Seek, W: Write + Seek>(
    input: R,
    output: W,
    patches: &WorkbookCellPatches,
    style_table: &StyleTable,
) -> Result<(), StreamingPatchError> {
    if patches.is_empty() {
        return patch_xlsx_streaming(input, output, &[]);
    }

    let mut archive = ZipArchive::new(input)?;

    let mut pre_read_parts: HashMap<String, Vec<u8>> = HashMap::new();
    let mut updated_parts: HashMap<String, Vec<u8>> = HashMap::new();
    let workbook_xml = read_zip_part(&mut archive, "xl/workbook.xml", &mut pre_read_parts)?;
    let workbook_xml = String::from_utf8(workbook_xml).map_err(crate::XlsxError::from)?;
    let workbook_sheets = parse_workbook_sheets(&workbook_xml)?;

    let workbook_rels_bytes = read_zip_part(
        &mut archive,
        "xl/_rels/workbook.xml.rels",
        &mut pre_read_parts,
    )?;
    let rels = parse_relationships(&workbook_rels_bytes)?;
    let mut rel_targets: HashMap<String, String> = HashMap::new();
    let mut styles_part: Option<String> = None;
    for rel in rels {
        let resolved = resolve_target("xl/workbook.xml", &rel.target);
        if rel.type_uri == REL_TYPE_STYLES {
            styles_part = Some(resolved.clone());
        }
        rel_targets.insert(rel.id, resolved);
    }

    let mut style_id_overrides: Vec<u32> = Vec::new();
    for (_, sheet_patches) in patches.sheets() {
        for (_, patch) in sheet_patches.iter() {
            if let Some(style_id) = patch.style_id().filter(|id| *id != 0) {
                style_id_overrides.push(style_id);
            }
        }
    }

    let style_id_to_xf = if style_id_overrides.is_empty() {
        None
    } else {
        let styles_part = styles_part.ok_or_else(|| {
            crate::XlsxError::Invalid("workbook.xml.rels missing styles relationship".to_string())
        })?;

        let styles_bytes = read_zip_part(&mut archive, &styles_part, &mut pre_read_parts)?;
        let mut style_table = style_table.clone();
        let mut styles_editor =
            XlsxStylesEditor::parse_or_default(Some(styles_bytes.as_slice()), &mut style_table)
                .map_err(|e| crate::XlsxError::Invalid(format!("styles.xml error: {e}")))?;

        let before_xfs = styles_editor.styles_part().cell_xfs_count();
        let style_id_to_xf = styles_editor
            .ensure_styles_for_style_ids(style_id_overrides, &style_table)
            .map_err(|e| crate::XlsxError::Invalid(format!("styles.xml error: {e}")))?;
        let after_xfs = styles_editor.styles_part().cell_xfs_count();

        if before_xfs != after_xfs {
            updated_parts.insert(styles_part.clone(), styles_editor.to_styles_xml_bytes());
        }

        Some(style_id_to_xf)
    };

    let mut patches_by_part: HashMap<String, Vec<WorksheetCellPatch>> = HashMap::new();
    let mut saw_formula_patch = false;
    for (sheet_selector, sheet_patches) in patches.sheets() {
        if sheet_patches.is_empty() {
            continue;
        }
        let worksheet_part = resolve_worksheet_part_for_selector(
            sheet_selector,
            &workbook_sheets,
            &rel_targets,
        )?;

        for (cell_ref, patch) in sheet_patches.iter() {
            let (value, formula) = match patch {
                CellPatch::Clear { .. } => (CellValue::Empty, None),
                CellPatch::Set { value, formula, .. } => (value.clone(), formula.clone()),
            };
            saw_formula_patch |= formula.is_some();

            let xf_index = if let Some(style_id) = patch.style_id() {
                if style_id == 0 {
                    Some(0)
                } else {
                    let map = style_id_to_xf.as_ref().ok_or_else(|| {
                        crate::XlsxError::Invalid(
                            "missing style_id mapping (styles.xml not updated)".to_string(),
                        )
                    })?;
                    Some(*map.get(&style_id).ok_or_else(|| {
                        crate::XlsxError::Invalid(format!("unknown style_id {style_id}"))
                    })?)
                }
            } else {
                patch.style_index()
            };

            patches_by_part
                .entry(worksheet_part.clone())
                .or_default()
                .push(
                    WorksheetCellPatch::new(worksheet_part.clone(), cell_ref, value, formula)
                        .with_xf_index(xf_index),
                );
        }
    }

    for patches in patches_by_part.values_mut() {
        patches.sort_by_key(|p| (p.cell.row, p.cell.col));
    }

    let mut formula_changed = saw_formula_patch;
    if !formula_changed {
        formula_changed = streaming_patches_remove_existing_formulas(&mut archive, &patches_by_part)?;
    }
    let recalc_policy = if formula_changed {
        RecalcPolicy::default()
    } else {
        RecalcPolicy::PRESERVE
    };

    patch_xlsx_streaming_with_archive(
        &mut archive,
        output,
        &patches_by_part,
        &pre_read_parts,
        &updated_parts,
        recalc_policy,
    )?;
    Ok(())
}

fn resolve_worksheet_part_for_selector(
    selector: &str,
    workbook_sheets: &[crate::WorkbookSheetInfo],
    rel_targets: &HashMap<String, String>,
) -> Result<String, crate::XlsxError> {
    let selector = selector.strip_prefix('/').unwrap_or(selector);

    // Worksheet part selector: any string containing `/` cannot be an Excel sheet name, and is
    // treated as an explicit part name.
    if selector.contains('/') {
        return Ok(selector.to_string());
    }

    // Sheet name selector (case-insensitive, matching Excel).
    if let Some(sheet) = workbook_sheets
        .iter()
        .find(|s| s.name.eq_ignore_ascii_case(selector))
    {
        return rel_targets
            .get(&sheet.rel_id)
            .cloned()
            .ok_or_else(|| {
                crate::XlsxError::Invalid(format!(
                    "missing worksheet relationship for {}",
                    sheet.name
                ))
            });
    }

    // RelId selector: if no sheet name matches, treat the key as a workbook relationship Id.
    if let Some(part) = rel_targets.get(selector) {
        return Ok(part.clone());
    }

    Err(crate::XlsxError::Invalid(format!(
        "unknown sheet selector: {selector} (tried sheet name match against xl/workbook.xml and relId lookup in xl/_rels/workbook.xml.rels; worksheet part selectors contain '/' e.g. xl/worksheets/sheet2.xml)"
    )))
}

fn plan_shared_strings<R: Read + Seek>(
    archive: &mut ZipArchive<R>,
    patches_by_part: &HashMap<String, Vec<WorksheetCellPatch>>,
    pre_read_parts: &HashMap<String, Vec<u8>>,
) -> Result<
    (
        Option<String>,
        HashMap<String, HashMap<(u32, u32), u32>>,
        Option<Vec<u8>>,
    ),
    StreamingPatchError,
> {
    let any_string_patch = patches_by_part
        .values()
        .flat_map(|patches| patches.iter())
        .any(|p| matches!(p.value, CellValue::String(_) | CellValue::RichText(_)));
    if !any_string_patch {
        return Ok((None, HashMap::new(), None));
    }

    let Some(shared_strings_part) = resolve_shared_strings_part_name(archive, pre_read_parts)?
    else {
        return Ok((None, HashMap::new(), None));
    };

    let existing_types = scan_existing_cell_types(archive, patches_by_part)?;

    let mut needs_shared_strings = false;
    for (part, patches) in patches_by_part {
        for patch in patches {
            let existing_t = existing_types
                .get(part)
                .and_then(|m| m.get(&(patch.cell.row, patch.cell.col)))
                .and_then(|t| t.as_deref());
            if patch_wants_shared_string(patch, existing_t, true) {
                needs_shared_strings = true;
                break;
            }
        }
        if needs_shared_strings {
            break;
        }
    }

    if !needs_shared_strings {
        return Ok((Some(shared_strings_part), HashMap::new(), None));
    }

    let mut count_delta: i32 = 0;
    for (part, patches) in patches_by_part {
        for patch in patches {
            let existing_t = existing_types
                .get(part)
                .and_then(|m| m.get(&(patch.cell.row, patch.cell.col)))
                .and_then(|t| t.as_deref());
            let old_uses_shared = matches!(existing_t, Some("s"));
            let new_uses_shared = patch_wants_shared_string(patch, existing_t, true);
            match (old_uses_shared, new_uses_shared) {
                (true, false) => count_delta -= 1,
                (false, true) => count_delta += 1,
                _ => {}
            }
        }
    }

    let mut shared_strings_bytes = Vec::new();
    {
        let mut file = archive.by_name(&shared_strings_part)?;
        file.read_to_end(&mut shared_strings_bytes)?;
    }
    let existing_shared_indices = scan_existing_shared_string_indices(archive, patches_by_part)?;
    let mut shared_strings = SharedStringsState::from_part(&shared_strings_bytes)?;
    shared_strings.count_delta = count_delta;

    // Deterministic insertion order: sort by (worksheet part, row, col).
    let mut shared_patches: Vec<(&str, &WorksheetCellPatch)> = patches_by_part
        .iter()
        .flat_map(|(part, patches)| patches.iter().map(move |p| (part.as_str(), p)))
        .filter(|(part, patch)| {
            let existing_t = existing_types
                .get(*part)
                .and_then(|m| m.get(&(patch.cell.row, patch.cell.col)))
                .and_then(|t| t.as_deref());
            patch_wants_shared_string(patch, existing_t, true)
        })
        .collect();
    shared_patches.sort_by_key(|(part, patch)| (*part, patch.cell.row, patch.cell.col));

    let mut indices_by_part: HashMap<String, HashMap<(u32, u32), u32>> = HashMap::new();
    for (part, patch) in shared_patches {
        let existing_t = existing_types
            .get(part)
            .and_then(|m| m.get(&(patch.cell.row, patch.cell.col)))
            .and_then(|t| t.as_deref());
        let existing_idx = existing_shared_indices
            .get(part)
            .and_then(|m| m.get(&(patch.cell.row, patch.cell.col)))
            .copied()
            .flatten();

        let reuse_idx = if existing_t == Some("s") {
            existing_idx.and_then(|idx| {
                let matches = match &patch.value {
                    CellValue::String(s) => shared_strings
                        .editor
                        .rich_at(idx)
                        .map(|rt| rt.text.as_str() == s)
                        .unwrap_or(false),
                    CellValue::RichText(rich) => shared_strings
                        .editor
                        .rich_at(idx)
                        .map(|rt| rt == rich)
                        .unwrap_or(false),
                    _ => false,
                };
                matches.then_some(idx)
            })
        } else {
            None
        };

        let idx = reuse_idx.unwrap_or_else(|| match &patch.value {
            CellValue::String(s) => shared_strings.get_or_insert_plain(s),
            CellValue::RichText(rich) => shared_strings.get_or_insert_rich(rich),
            _ => 0,
        });
        indices_by_part
            .entry(part.to_string())
            .or_default()
            .insert((patch.cell.row, patch.cell.col), idx);
    }

    let updated_shared_strings = shared_strings.write_if_dirty()?;
    Ok((Some(shared_strings_part), indices_by_part, updated_shared_strings))
}

fn resolve_shared_strings_part_name<R: Read + Seek>(
    archive: &mut ZipArchive<R>,
    pre_read_parts: &HashMap<String, Vec<u8>>,
) -> Result<Option<String>, StreamingPatchError> {
    let workbook_rels = if let Some(bytes) = pre_read_parts.get("xl/_rels/workbook.xml.rels") {
        Some(bytes.clone())
    } else {
        read_zip_part_optional(archive, "xl/_rels/workbook.xml.rels")?
    };

    if let Some(bytes) = workbook_rels {
        let rels = parse_relationships(&bytes)?;
        if let Some(rel) = rels
            .iter()
            .find(|rel| rel.type_uri == REL_TYPE_SHARED_STRINGS)
        {
            return Ok(Some(resolve_target("xl/workbook.xml", &rel.target)));
        }
    }

    // Fallback: common path when workbook.xml.rels is missing the sharedStrings relationship.
    if zip_part_exists(archive, "xl/sharedStrings.xml")? {
        return Ok(Some("xl/sharedStrings.xml".to_string()));
    }

    Ok(None)
}

fn scan_existing_cell_types<R: Read + Seek>(
    archive: &mut ZipArchive<R>,
    patches_by_part: &HashMap<String, Vec<WorksheetCellPatch>>,
) -> Result<HashMap<String, HashMap<(u32, u32), Option<String>>>, StreamingPatchError> {
    let mut out: HashMap<String, HashMap<(u32, u32), Option<String>>> = HashMap::new();

    for (part, patches) in patches_by_part {
        let mut targets: HashMap<String, (u32, u32)> = HashMap::new();
        for patch in patches {
            targets.insert(patch.cell.to_a1(), (patch.cell.row, patch.cell.col));
        }
        if targets.is_empty() {
            continue;
        }

        let mut file = match archive.by_name(part) {
            Ok(file) => file,
            Err(zip::result::ZipError::FileNotFound) => {
                return Err(StreamingPatchError::MissingWorksheetPart(part.clone()));
            }
            Err(err) => return Err(err.into()),
        };

        let found = scan_worksheet_cell_types(&mut file, &targets)?;
        out.insert(part.clone(), found);
    }

    Ok(out)
}

fn scan_worksheet_cell_types<R: Read>(
    input: R,
    targets: &HashMap<String, (u32, u32)>,
) -> Result<HashMap<(u32, u32), Option<String>>, StreamingPatchError> {
    let mut out: HashMap<(u32, u32), Option<String>> = HashMap::new();
    let mut remaining = targets.len();
    if remaining == 0 {
        return Ok(out);
    }

    let mut reader = Reader::from_reader(BufReader::new(input));
    reader.config_mut().trim_text(true);
    let mut buf = Vec::new();

    loop {
        match reader.read_event_into(&mut buf)? {
            Event::Start(ref e) | Event::Empty(ref e) if local_name(e.name().as_ref()) == b"c" => {
                let mut r: Option<String> = None;
                let mut t: Option<String> = None;
                for attr in e.attributes() {
                    let attr = attr?;
                    match attr.key.as_ref() {
                        b"r" => r = Some(attr.unescape_value()?.into_owned()),
                        b"t" => t = Some(attr.unescape_value()?.into_owned()),
                        _ => {}
                    }
                }
                if let Some(r) = r {
                    if let Some(&(row, col)) = targets.get(&r) {
                        out.insert((row, col), t);
                        remaining = remaining.saturating_sub(1);
                        if remaining == 0 {
                            break;
                        }
                    }
                }
            }
            Event::Eof => break,
            _ => {}
        }
        buf.clear();
    }

    Ok(out)
}

fn scan_existing_shared_string_indices<R: Read + Seek>(
    archive: &mut ZipArchive<R>,
    patches_by_part: &HashMap<String, Vec<WorksheetCellPatch>>,
) -> Result<HashMap<String, HashMap<(u32, u32), Option<u32>>>, StreamingPatchError> {
    let mut out: HashMap<String, HashMap<(u32, u32), Option<u32>>> = HashMap::new();

    for (part, patches) in patches_by_part {
        let mut targets: HashMap<String, (u32, u32)> = HashMap::new();
        for patch in patches {
            if matches!(patch.value, CellValue::String(_) | CellValue::RichText(_)) {
                targets.insert(patch.cell.to_a1(), (patch.cell.row, patch.cell.col));
            }
        }
        if targets.is_empty() {
            continue;
        }

        let mut file = match archive.by_name(part) {
            Ok(file) => file,
            Err(zip::result::ZipError::FileNotFound) => {
                return Err(StreamingPatchError::MissingWorksheetPart(part.clone()));
            }
            Err(err) => return Err(err.into()),
        };

        let found = scan_worksheet_shared_string_indices(&mut file, &targets)?;
        out.insert(part.clone(), found);
    }

    Ok(out)
}

fn scan_worksheet_shared_string_indices<R: Read>(
    input: R,
    targets: &HashMap<String, (u32, u32)>,
) -> Result<HashMap<(u32, u32), Option<u32>>, StreamingPatchError> {
    let mut out: HashMap<(u32, u32), Option<u32>> = HashMap::new();
    let mut remaining = targets.len();
    if remaining == 0 {
        return Ok(out);
    }

    let mut reader = Reader::from_reader(BufReader::new(input));
    reader.config_mut().trim_text(true);
    let mut buf = Vec::new();

    let mut current_target: Option<(u32, u32)> = None;
    let mut current_t: Option<String> = None;
    let mut in_v = false;
    let mut current_idx: Option<u32> = None;

    loop {
        match reader.read_event_into(&mut buf)? {
            Event::Start(ref e) if local_name(e.name().as_ref()) == b"c" => {
                let mut r: Option<String> = None;
                let mut t: Option<String> = None;
                for attr in e.attributes() {
                    let attr = attr?;
                    match attr.key.as_ref() {
                        b"r" => r = Some(attr.unescape_value()?.into_owned()),
                        b"t" => t = Some(attr.unescape_value()?.into_owned()),
                        _ => {}
                    }
                }
                if let Some(r) = r {
                    if let Some(&(row, col)) = targets.get(&r) {
                        current_target = Some((row, col));
                        current_t = t;
                        current_idx = None;
                    }
                }
            }
            Event::Empty(ref e) if local_name(e.name().as_ref()) == b"c" => {
                let mut r: Option<String> = None;
                for attr in e.attributes() {
                    let attr = attr?;
                    if attr.key.as_ref() == b"r" {
                        r = Some(attr.unescape_value()?.into_owned());
                    }
                }
                if let Some(r) = r {
                    if let Some(&(row, col)) = targets.get(&r) {
                        out.insert((row, col), None);
                        remaining = remaining.saturating_sub(1);
                        if remaining == 0 {
                            break;
                        }
                    }
                }
            }
            Event::Start(ref e) if current_target.is_some() && local_name(e.name().as_ref()) == b"v" => {
                in_v = true;
            }
            Event::End(ref e) if current_target.is_some() && local_name(e.name().as_ref()) == b"v" => {
                in_v = false;
            }
            Event::Text(e) if in_v && current_target.is_some() => {
                if current_t.as_deref() == Some("s") {
                    current_idx = e.unescape()?.trim().parse::<u32>().ok();
                }
            }
            Event::End(ref e) if current_target.is_some() && local_name(e.name().as_ref()) == b"c" => {
                let coord = current_target.take().unwrap();
                out.insert(coord, current_idx);
                current_t = None;
                in_v = false;
                current_idx = None;
                remaining = remaining.saturating_sub(1);
                if remaining == 0 {
                    break;
                }
            }
            Event::Eof => break,
            _ => {}
        }
        buf.clear();
    }

    Ok(out)
}

#[derive(Debug, Clone, Copy, Default)]
pub(crate) struct WorksheetXmlMetadata {
    has_dimension: bool,
    has_sheet_pr: bool,
    existing_used_range: Option<PatchBounds>,
}

fn scan_worksheet_xml_metadata<R: Read>(
    input: R,
) -> Result<WorksheetXmlMetadata, StreamingPatchError> {
    let mut reader = Reader::from_reader(BufReader::new(input));
    reader.config_mut().trim_text(true);
    let mut buf = Vec::new();

    let mut in_sheet_data = false;
    let mut has_dimension = false;
    let mut has_sheet_pr = false;
    let mut used_range: Option<PatchBounds> = None;

    loop {
        match reader.read_event_into(&mut buf)? {
            Event::Start(ref e) if local_name(e.name().as_ref()) == b"sheetData" => {
                in_sheet_data = true;
            }
            Event::Empty(ref e) if local_name(e.name().as_ref()) == b"sheetData" => {
                in_sheet_data = false;
            }
            Event::End(ref e) if local_name(e.name().as_ref()) == b"sheetData" => {
                in_sheet_data = false;
            }
            Event::Start(ref e) | Event::Empty(ref e)
                if local_name(e.name().as_ref()) == b"dimension" =>
            {
                has_dimension = true;
            }
            Event::Start(ref e) | Event::Empty(ref e)
                if local_name(e.name().as_ref()) == b"sheetPr" =>
            {
                has_sheet_pr = true;
            }
            Event::Start(ref e) | Event::Empty(ref e)
                if local_name(e.name().as_ref()) == b"mergeCell" =>
            {
                let mut r: Option<String> = None;
                for attr in e.attributes() {
                    let attr = attr?;
                    if local_name(attr.key.as_ref()) == b"ref" {
                        r = Some(attr.unescape_value()?.into_owned());
                        break;
                    }
                }
                if let Some(r) = r {
                    let mut parts = r.split(':');
                    let start = parts.next().unwrap_or_default();
                    let start = CellRef::from_a1(start)
                        .map_err(|_| StreamingPatchError::InvalidCellRef(r.clone()))?;
                    let end = parts
                        .next()
                        .map(|p| {
                            CellRef::from_a1(p)
                                .map_err(|_| StreamingPatchError::InvalidCellRef(r.clone()))
                        })
                        .transpose()?
                        .unwrap_or(start);
                    let (min_row_0, max_row_0) = (start.row.min(end.row), start.row.max(end.row));
                    let (min_col_0, max_col_0) = (start.col.min(end.col), start.col.max(end.col));
                    used_range = Some(match used_range {
                        Some(existing) => PatchBounds {
                            min_row_0: existing.min_row_0.min(min_row_0),
                            max_row_0: existing.max_row_0.max(max_row_0),
                            min_col_0: existing.min_col_0.min(min_col_0),
                            max_col_0: existing.max_col_0.max(max_col_0),
                        },
                        None => PatchBounds {
                            min_row_0,
                            max_row_0,
                            min_col_0,
                            max_col_0,
                        },
                    });
                }
            }
            Event::Start(ref e) | Event::Empty(ref e)
                if in_sheet_data && local_name(e.name().as_ref()) == b"c" =>
            {
                let mut r: Option<String> = None;
                for attr in e.attributes() {
                    let attr = attr?;
                    if local_name(attr.key.as_ref()) == b"r" {
                        r = Some(attr.unescape_value()?.into_owned());
                        break;
                    }
                }
                if let Some(r) = r {
                    let cell_ref =
                        CellRef::from_a1(&r).map_err(|_| StreamingPatchError::InvalidCellRef(r))?;
                    used_range = Some(match used_range {
                        Some(existing) => PatchBounds {
                            min_row_0: existing.min_row_0.min(cell_ref.row),
                            max_row_0: existing.max_row_0.max(cell_ref.row),
                            min_col_0: existing.min_col_0.min(cell_ref.col),
                            max_col_0: existing.max_col_0.max(cell_ref.col),
                        },
                        None => PatchBounds {
                            min_row_0: cell_ref.row,
                            max_row_0: cell_ref.row,
                            min_col_0: cell_ref.col,
                            max_col_0: cell_ref.col,
                        },
                    });
                }
            }
            Event::Eof => break,
            _ => {}
        }
        buf.clear();
    }

    Ok(WorksheetXmlMetadata {
        has_dimension,
        has_sheet_pr,
        existing_used_range: used_range,
    })
}

fn patch_wants_shared_string(
    patch: &WorksheetCellPatch,
    existing_t: Option<&str>,
    shared_strings_available: bool,
) -> bool {
    if !shared_strings_available {
        return false;
    }

    match &patch.value {
        CellValue::String(_) => {
            if existing_t.is_some_and(should_preserve_unknown_t) {
                return false;
            }

            match existing_t {
                Some("inlineStr" | "str") => return false,
                Some("s") => return true,
                _ => {}
            }

            // Preserve streaming patcher's historical behavior for formula string results: use
            // `t="str"` unless the original cell already used shared strings.
            if patch.formula.is_some() {
                return false;
            }

            true
        }
        CellValue::RichText(_) => {
            if existing_t.is_some_and(should_preserve_unknown_t) {
                return false;
            }
            // Preserve inline strings when the existing cell already uses inline storage.
            existing_t != Some("inlineStr")
        }
        _ => false,
    }
}

fn zip_part_exists<R: Read + Seek>(
    archive: &mut ZipArchive<R>,
    name: &str,
) -> Result<bool, StreamingPatchError> {
    match archive.by_name(name) {
        Ok(_) => Ok(true),
        Err(zip::result::ZipError::FileNotFound) => Ok(false),
        Err(err) => Err(err.into()),
    }
}

fn read_zip_part_optional<R: Read + Seek>(
    archive: &mut ZipArchive<R>,
    name: &str,
) -> Result<Option<Vec<u8>>, StreamingPatchError> {
    let mut file = match archive.by_name(name) {
        Ok(file) => file,
        Err(zip::result::ZipError::FileNotFound) => return Ok(None),
        Err(err) => return Err(err.into()),
    };
    let mut buf = Vec::with_capacity(file.size() as usize);
    file.read_to_end(&mut buf)?;
    Ok(Some(buf))
}

fn patch_xlsx_streaming_with_archive<R: Read + Seek, W: Write + Seek>(
    archive: &mut ZipArchive<R>,
    output: W,
    patches_by_part: &HashMap<String, Vec<WorksheetCellPatch>>,
    pre_read_parts: &HashMap<String, Vec<u8>>,
    updated_parts: &HashMap<String, Vec<u8>>,
    recalc_policy: RecalcPolicy,
) -> Result<(), StreamingPatchError> {
    let (shared_strings_part, shared_string_indices, shared_strings_updated) =
        plan_shared_strings(archive, patches_by_part, pre_read_parts)?;

    let mut worksheet_metadata_by_part: HashMap<String, WorksheetXmlMetadata> = HashMap::new();
    for part in patches_by_part.keys() {
        let mut file = match archive.by_name(part) {
            Ok(file) => file,
            Err(zip::result::ZipError::FileNotFound) => {
                return Err(StreamingPatchError::MissingWorksheetPart(part.clone()));
            }
            Err(err) => return Err(err.into()),
        };
        worksheet_metadata_by_part.insert(part.clone(), scan_worksheet_xml_metadata(&mut file)?);
    }

    let mut missing_parts: BTreeMap<String, ()> =
        patches_by_part.keys().map(|k| (k.clone(), ())).collect();

    let mut zip = ZipWriter::new(output);
    let options = FileOptions::<()>::default().compression_method(CompressionMethod::Deflated);

    for i in 0..archive.len() {
        let mut file = archive.by_index(i)?;
        if file.is_dir() {
            continue;
        }

        let name = file.name().to_string();

        if recalc_policy.drop_calc_chain_on_formula_change && name == "xl/calcChain.xml" {
            // Drop calcChain.xml entirely when formulas change, matching the in-memory patcher.
            continue;
        }

        if let Some(patches) = patches_by_part.get(&name) {
            zip.start_file(name.clone(), options)?;
            missing_parts.remove(&name);
            let indices = shared_string_indices.get(&name);
            let worksheet_meta = worksheet_metadata_by_part
                .get(&name)
                .copied()
                .unwrap_or_default();
            patch_worksheet_xml_streaming(
                &mut file,
                &mut zip,
                &name,
                patches,
                indices,
                worksheet_meta,
            )?;
        } else if let Some(bytes) = updated_parts.get(&name) {
            zip.start_file(name.clone(), options)?;
            zip.write_all(bytes)?;
        } else if shared_strings_part.as_deref() == Some(name.as_str())
            && shared_strings_updated.is_some()
        {
            zip.start_file(name.clone(), options)?;
            zip.write_all(
                shared_strings_updated
                    .as_deref()
                    .expect("checked is_some above"),
            )?;
        } else if let Some(bytes) = pre_read_parts.get(&name) {
            if should_patch_recalc_part(&name, recalc_policy) {
                zip.start_file(name.clone(), options)?;
                let bytes = maybe_patch_recalc_part(&name, bytes, recalc_policy)?;
                zip.write_all(&bytes)?;
            } else {
                // We buffered this part earlier for metadata resolution, but it doesn't need to be
                // rewritten. Raw-copy it to avoid recompression.
                zip.raw_copy_file(file)?;
            }
        } else if let Some(updated) = patch_recalc_part_from_file(&name, &mut file, recalc_policy)?
        {
            zip.start_file(name.clone(), options)?;
            zip.write_all(&updated)?;
        } else {
            // Use raw copy to preserve bytes for unchanged parts and avoid a decompression /
            // recompression pass over large binary assets.
            zip.raw_copy_file(file)?;
        }
    }

    if let Some((missing, _)) = missing_parts.into_iter().next() {
        return Err(StreamingPatchError::MissingWorksheetPart(missing));
    }

    zip.finish()?;
    Ok(())
}

fn streaming_patches_remove_existing_formulas<R: Read + Seek>(
    archive: &mut ZipArchive<R>,
    patches_by_part: &HashMap<String, Vec<WorksheetCellPatch>>,
) -> Result<bool, StreamingPatchError> {
    for (worksheet_part, patches) in patches_by_part {
        let mut target_cells: HashSet<String> = HashSet::new();
        for patch in patches {
            if patch.formula.is_none() {
                target_cells.insert(patch.cell.to_a1());
            }
        }
        if target_cells.is_empty() {
            continue;
        }

        let file = match archive.by_name(worksheet_part) {
            Ok(file) => file,
            // If the worksheet is missing, the main streaming pass will surface this as
            // `MissingWorksheetPart`. Don't change the error type here.
            Err(zip::result::ZipError::FileNotFound) => continue,
            Err(err) => return Err(err.into()),
        };

        if worksheet_contains_formula_in_cells(file, &target_cells)? {
            return Ok(true);
        }
    }

    Ok(false)
}

fn worksheet_contains_formula_in_cells<R: Read>(
    input: R,
    target_cells: &HashSet<String>,
) -> Result<bool, StreamingPatchError> {
    let mut reader = Reader::from_reader(BufReader::new(input));
    reader.config_mut().trim_text(false);

    let mut buf = Vec::new();
    loop {
        let event = reader.read_event_into(&mut buf)?;
        match event {
            Event::Eof => break,
            Event::Start(ref e) if local_name(e.name().as_ref()) == b"c" => {
                if cell_is_in_set(e, target_cells)? {
                    if cell_contains_formula(&mut reader, &mut buf)? {
                        return Ok(true);
                    }
                }
            }
            // `<c .../>` cannot contain a formula.
            Event::Empty(ref e) if local_name(e.name().as_ref()) == b"c" => {
                if cell_is_in_set(e, target_cells)? {
                    // Target cell exists but is empty / non-formula.
                }
            }
            _ => {}
        }
        buf.clear();
    }

    Ok(false)
}

fn cell_is_in_set(
    cell_start: &BytesStart<'_>,
    target_cells: &HashSet<String>,
) -> Result<bool, StreamingPatchError> {
    for attr in cell_start.attributes() {
        let attr = attr?;
        if attr.key.as_ref() == b"r" {
            let r = attr.unescape_value()?;
            return Ok(target_cells.contains(r.as_ref()));
        }
    }
    Ok(false)
}

fn cell_contains_formula<R: BufRead>(
    reader: &mut Reader<R>,
    buf: &mut Vec<u8>,
) -> Result<bool, StreamingPatchError> {
    // We are currently positioned just after the opening `<c ...>` event; scan until the matching
    // `</c>` looking for an `<f>` element.
    let mut depth = 1usize;
    loop {
        let event = reader.read_event_into(buf)?;
        match event {
            Event::Eof => break,
            Event::Start(ref e) if local_name(e.name().as_ref()) == b"f" => return Ok(true),
            Event::Empty(ref e) if local_name(e.name().as_ref()) == b"f" => return Ok(true),
            Event::Start(_) => depth += 1,
            Event::End(_) => {
                depth = depth.saturating_sub(1);
                if depth == 0 {
                    break;
                }
            }
            _ => {}
        }
        buf.clear();
    }
    Ok(false)
}

fn maybe_patch_recalc_part(
    name: &str,
    bytes: &[u8],
    recalc_policy: RecalcPolicy,
) -> Result<Vec<u8>, StreamingPatchError> {
    match name {
        "xl/workbook.xml" if recalc_policy.force_full_calc_on_formula_change => {
            Ok(workbook_xml_force_full_calc_on_load(bytes)?)
        }
        "xl/_rels/workbook.xml.rels" if recalc_policy.drop_calc_chain_on_formula_change => {
            Ok(workbook_rels_remove_calc_chain(bytes)?)
        }
        "[Content_Types].xml" if recalc_policy.drop_calc_chain_on_formula_change => {
            Ok(content_types_remove_calc_chain(bytes)?)
        }
        _ => Ok(bytes.to_vec()),
    }
}

fn patch_recalc_part_from_file<R: Read>(
    name: &str,
    file: &mut R,
    recalc_policy: RecalcPolicy,
) -> Result<Option<Vec<u8>>, StreamingPatchError> {
    if !should_patch_recalc_part(name, recalc_policy) {
        return Ok(None);
    }

    let mut buf = Vec::new();
    file.read_to_end(&mut buf)?;
    Ok(Some(maybe_patch_recalc_part(name, &buf, recalc_policy)?))
}

fn should_patch_recalc_part(name: &str, recalc_policy: RecalcPolicy) -> bool {
    match name {
        "xl/workbook.xml" => recalc_policy.force_full_calc_on_formula_change,
        "xl/_rels/workbook.xml.rels" | "[Content_Types].xml" => {
            recalc_policy.drop_calc_chain_on_formula_change
        }
        _ => false,
    }
}

fn read_zip_part<R: Read + Seek>(
    archive: &mut ZipArchive<R>,
    name: &str,
    cache: &mut HashMap<String, Vec<u8>>,
) -> Result<Vec<u8>, StreamingPatchError> {
    if let Some(bytes) = cache.get(name) {
        return Ok(bytes.clone());
    }
    let mut file = archive.by_name(name)?;
    let mut buf = Vec::with_capacity(file.size() as usize);
    file.read_to_end(&mut buf)?;
    cache.insert(name.to_string(), buf.clone());
    Ok(buf)
}

#[derive(Debug)]
struct SharedStringsState {
    editor: SharedStringsEditor,
    // Best-effort shared-string reference count delta from cell patches.
    count_delta: i32,
}

impl SharedStringsState {
    fn from_part(bytes: &[u8]) -> Result<Self, StreamingPatchError> {
        let editor = SharedStringsEditor::parse(bytes).map_err(|e| {
            crate::XlsxError::Invalid(format!("sharedStrings.xml parse error: {e}"))
        })?;
        Ok(Self {
            editor,
            count_delta: 0,
        })
    }

    fn get_or_insert_plain(&mut self, text: &str) -> u32 {
        self.editor.get_or_insert_plain(text)
    }

    fn get_or_insert_rich(&mut self, rich: &RichText) -> u32 {
        self.editor.get_or_insert_rich(rich)
    }

    fn write_if_dirty(&self) -> Result<Option<Vec<u8>>, StreamingPatchError> {
        if !self.editor.is_dirty() {
            return Ok(None);
        }

        let count_hint = self
            .editor
            .original_count()
            .map(|base| base.saturating_add_signed(self.count_delta));
        let updated = self.editor.to_xml_bytes(count_hint).map_err(|e| {
            crate::XlsxError::Invalid(format!("sharedStrings.xml write error: {e}"))
        })?;
        Ok(Some(updated))
    }
}

fn write_shared_string_t<W: Write>(
    writer: &mut Writer<W>,
    t_tag: &str,
    text: &str,
) -> std::io::Result<()> {
    let mut t = BytesStart::new(t_tag);
    if needs_space_preserve(text) {
        t.push_attribute(("xml:space", "preserve"));
    }
    writer.write_event(Event::Start(t))?;
    writer.write_event(Event::Text(BytesText::new(text)))?;
    writer.write_event(Event::End(BytesEnd::new(t_tag)))?;
    Ok(())
}

fn write_shared_string_rpr<W: Write>(
    writer: &mut Writer<W>,
    prefix: Option<&str>,
    style: &formula_model::rich_text::RichTextRunStyle,
) -> std::io::Result<()> {
    if let Some(font) = &style.font {
        let mut rfont = BytesStart::new(prefixed_tag(prefix, "rFont"));
        rfont.push_attribute(("val", font.as_str()));
        writer.write_event(Event::Empty(rfont))?;
    }

    if let Some(size_100pt) = style.size_100pt {
        let mut sz = BytesStart::new(prefixed_tag(prefix, "sz"));
        let value = format_size_100pt(size_100pt);
        sz.push_attribute(("val", value.as_str()));
        writer.write_event(Event::Empty(sz))?;
    }

    if let Some(color) = style.color {
        let mut c = BytesStart::new(prefixed_tag(prefix, "color"));
        let value = format!("{:08X}", color.argb().unwrap_or(0));
        c.push_attribute(("rgb", value.as_str()));
        writer.write_event(Event::Empty(c))?;
    }

    if let Some(bold) = style.bold {
        let mut b = BytesStart::new(prefixed_tag(prefix, "b"));
        if !bold {
            b.push_attribute(("val", "0"));
        }
        writer.write_event(Event::Empty(b))?;
    }

    if let Some(italic) = style.italic {
        let mut i = BytesStart::new(prefixed_tag(prefix, "i"));
        if !italic {
            i.push_attribute(("val", "0"));
        }
        writer.write_event(Event::Empty(i))?;
    }

    if let Some(ul) = style.underline {
        let mut u = BytesStart::new(prefixed_tag(prefix, "u"));
        if let Some(val) = ul.to_ooxml() {
            u.push_attribute(("val", val));
        }
        writer.write_event(Event::Empty(u))?;
    }

    Ok(())
}

fn format_size_100pt(size_100pt: u16) -> String {
    let int = size_100pt / 100;
    let frac = size_100pt % 100;
    if frac == 0 {
        return int.to_string();
    }

    let mut s = format!("{int}.{frac:02}");
    while s.ends_with('0') {
        s.pop();
    }
    s
}

#[derive(Debug, Clone)]
struct CellPatchInternal {
    row_1: u32,
    col_0: u32,
    value: CellValue,
    formula: Option<String>,
    xf_index: Option<u32>,
    shared_string_idx: Option<u32>,
}

struct RowState {
    row_1: u32,
    pending: Vec<CellPatchInternal>,
    next_idx: usize,
    cell_prefix: Option<String>,
}

pub(crate) fn patch_worksheet_xml_streaming<R: Read, W: Write>(
    input: R,
    output: W,
    worksheet_part: &str,
    patches: &[WorksheetCellPatch],
    shared_string_indices: Option<&HashMap<(u32, u32), u32>>,
    worksheet_meta: WorksheetXmlMetadata,
) -> Result<(), StreamingPatchError> {
    let patch_bounds = bounds_for_patches(patches);
    let dimension_ref_to_insert = if worksheet_meta.has_dimension {
        None
    } else {
        union_bounds(worksheet_meta.existing_used_range, patch_bounds).map(bounds_to_dimension_ref)
    };
    let insert_dimension_after_sheet_pr =
        dimension_ref_to_insert.is_some() && worksheet_meta.has_sheet_pr;
    let insert_dimension_at_worksheet_start =
        dimension_ref_to_insert.is_some() && !worksheet_meta.has_sheet_pr;

    let mut patches_by_row: BTreeMap<u32, Vec<CellPatchInternal>> = BTreeMap::new();
    for patch in patches {
        let row_1 = patch.cell.row + 1;
        let col_0 = patch.cell.col;
        let shared_string_idx = shared_string_indices
            .and_then(|m| m.get(&(patch.cell.row, patch.cell.col)).copied());
        patches_by_row
            .entry(row_1)
            .or_default()
            .push(CellPatchInternal {
                row_1,
                col_0,
                value: patch.value.clone(),
                formula: patch.formula.clone(),
                xf_index: patch.xf_index,
                shared_string_idx,
            });
    }
    for row_patches in patches_by_row.values_mut() {
        row_patches.sort_by_key(|p| p.col_0);
    }

    let mut reader = Reader::from_reader(BufReader::new(input));
    reader.config_mut().trim_text(false);
    let mut writer = Writer::new(output);

    let mut buf = Vec::new();
    let mut in_sheet_data = false;
    let mut saw_sheet_data = false;
    let mut patched_dimension = false;
    let mut worksheet_prefix: Option<String> = None;
    let mut worksheet_has_default_ns = false;
    let mut sheet_prefix: Option<String> = None;
    let mut inserted_dimension = false;
    let mut pending_dimension_after_sheet_pr_end = false;

    let mut row_state: Option<RowState> = None;
    let mut in_cell = false;

    loop {
        let event = reader.read_event_into(&mut buf)?;
        match event {
            Event::Eof => break,

            Event::Start(ref e) if local_name(e.name().as_ref()) == b"worksheet" => {
                if worksheet_prefix.is_none() {
                    worksheet_prefix = element_prefix(e.name().as_ref())
                        .and_then(|p| std::str::from_utf8(p).ok())
                        .map(|s| s.to_string());
                    worksheet_has_default_ns = worksheet_has_default_spreadsheetml_ns(e)?;
                }
                writer.write_event(Event::Start(e.to_owned()))?;
                if insert_dimension_at_worksheet_start && !inserted_dimension {
                    if let Some(ref dimension_ref) = dimension_ref_to_insert {
                        let prefix = if worksheet_has_default_ns {
                            None
                        } else {
                            worksheet_prefix.as_deref()
                        };
                        write_dimension_element(&mut writer, prefix, dimension_ref)?;
                        inserted_dimension = true;
                    }
                }
            }

            Event::Start(ref e) if local_name(e.name().as_ref()) == b"sheetPr" => {
                writer.write_event(Event::Start(e.to_owned()))?;
                if insert_dimension_after_sheet_pr && !inserted_dimension {
                    pending_dimension_after_sheet_pr_end = true;
                }
            }
            Event::Empty(ref e) if local_name(e.name().as_ref()) == b"sheetPr" => {
                writer.write_event(Event::Empty(e.to_owned()))?;
                if insert_dimension_after_sheet_pr && !inserted_dimension {
                    if let Some(ref dimension_ref) = dimension_ref_to_insert {
                        let prefix = if worksheet_has_default_ns {
                            None
                        } else {
                            worksheet_prefix.as_deref()
                        };
                        write_dimension_element(&mut writer, prefix, dimension_ref)?;
                        inserted_dimension = true;
                    }
                }
            }
            Event::End(ref e) if local_name(e.name().as_ref()) == b"sheetPr" => {
                writer.write_event(Event::End(e.to_owned()))?;
                if pending_dimension_after_sheet_pr_end && !inserted_dimension {
                    if let Some(ref dimension_ref) = dimension_ref_to_insert {
                        let prefix = if worksheet_has_default_ns {
                            None
                        } else {
                            worksheet_prefix.as_deref()
                        };
                        write_dimension_element(&mut writer, prefix, dimension_ref)?;
                        inserted_dimension = true;
                    }
                    pending_dimension_after_sheet_pr_end = false;
                }
            }

            Event::Start(ref e) if local_name(e.name().as_ref()) == b"sheetData" => {
                saw_sheet_data = true;
                in_sheet_data = true;
                if sheet_prefix.is_none() {
                    sheet_prefix = element_prefix(e.name().as_ref())
                        .and_then(|p| std::str::from_utf8(p).ok())
                        .map(|s| s.to_string());
                }
                writer.write_event(Event::Start(e.to_owned()))?;
            }
            Event::Empty(ref e) if local_name(e.name().as_ref()) == b"sheetData" => {
                saw_sheet_data = true;
                if patches_by_row.is_empty() {
                    writer.write_event(Event::Empty(e.to_owned()))?;
                } else {
                    if sheet_prefix.is_none() {
                        sheet_prefix = element_prefix(e.name().as_ref())
                            .and_then(|p| std::str::from_utf8(p).ok())
                            .map(|s| s.to_string());
                    }
                    in_sheet_data = false;
                    // Expand `<sheetData/>` into `<sheetData>...</sheetData>`.
                    let sheet_data_tag = String::from_utf8_lossy(e.name().as_ref()).into_owned();
                    writer.write_event(Event::Start(e.to_owned()))?;
                    write_pending_rows(&mut writer, &mut patches_by_row, sheet_prefix.as_deref())?;
                    writer.write_event(Event::End(BytesEnd::new(sheet_data_tag.as_str())))?;
                }
            }
            Event::End(ref e) if local_name(e.name().as_ref()) == b"sheetData" => {
                // Flush any remaining patch rows at the end of sheetData.
                write_pending_rows(&mut writer, &mut patches_by_row, sheet_prefix.as_deref())?;
                in_sheet_data = false;
                writer.write_event(Event::End(e.to_owned()))?;
            }
            Event::End(ref e) if local_name(e.name().as_ref()) == b"worksheet" => {
                if !saw_sheet_data && !patches_by_row.is_empty() {
                    saw_sheet_data = true;
                    let sheet_prefix = if worksheet_has_default_ns {
                        None
                    } else {
                        worksheet_prefix.as_deref()
                    };
                    let sheet_data_tag = prefixed_tag(sheet_prefix, "sheetData");
                    writer.write_event(Event::Start(BytesStart::new(sheet_data_tag.as_str())))?;
                    write_pending_rows(&mut writer, &mut patches_by_row, sheet_prefix)?;
                    writer.write_event(Event::End(BytesEnd::new(sheet_data_tag.as_str())))?;
                }
                writer.write_event(Event::End(e.to_owned()))?;
            }

            Event::Start(ref e) if local_name(e.name().as_ref()) == b"dimension" => {
                if !patched_dimension {
                    patched_dimension = true;
                    if let Some(bounds) = patch_bounds {
                        let updated = updated_dimension_element(e, bounds)?;
                        writer.write_event(Event::Start(updated))?;
                    } else {
                        writer.write_event(Event::Start(e.to_owned()))?;
                    }
                } else {
                    writer.write_event(Event::Start(e.to_owned()))?;
                }
            }
            Event::Empty(ref e) if local_name(e.name().as_ref()) == b"dimension" => {
                if !patched_dimension {
                    patched_dimension = true;
                    if let Some(bounds) = patch_bounds {
                        let updated = updated_dimension_element(e, bounds)?;
                        writer.write_event(Event::Empty(updated))?;
                    } else {
                        writer.write_event(Event::Empty(e.to_owned()))?;
                    }
                } else {
                    writer.write_event(Event::Empty(e.to_owned()))?;
                }
            }

            Event::Start(ref e) if in_sheet_data && local_name(e.name().as_ref()) == b"row" => {
                let row_1 = parse_row_number(e)?;
                in_cell = false;

                // Insert any patch rows that should appear before this row.
                while let Some((&next_row, _)) = patches_by_row.iter().next() {
                    if next_row < row_1 {
                        let pending = patches_by_row.remove(&next_row).unwrap_or_default();
                        write_inserted_row(&mut writer, next_row, &pending, sheet_prefix.as_deref())?;
                    } else {
                        break;
                    }
                }

                let pending = patches_by_row.remove(&row_1);
                if let Some(mut pending) = pending {
                    pending.sort_by_key(|p| p.col_0);
                    let updated_row = updated_row_spans_element(e, &pending)?;
                    row_state = Some(RowState {
                        row_1,
                        pending,
                        next_idx: 0,
                        cell_prefix: None,
                    });
                    if let Some(updated) = updated_row {
                        writer.write_event(Event::Start(updated))?;
                    } else {
                        writer.write_event(Event::Start(e.to_owned()))?;
                    }
                } else {
                    writer.write_event(Event::Start(e.to_owned()))?;
                }
            }
            Event::Empty(ref e) if in_sheet_data && local_name(e.name().as_ref()) == b"row" => {
                let row_1 = parse_row_number(e)?;
                in_cell = false;

                // Insert patch rows that should appear before this row.
                while let Some((&next_row, _)) = patches_by_row.iter().next() {
                    if next_row < row_1 {
                        let pending = patches_by_row.remove(&next_row).unwrap_or_default();
                        write_inserted_row(&mut writer, next_row, &pending, sheet_prefix.as_deref())?;
                    } else {
                        break;
                    }
                }

                if let Some(mut pending) = patches_by_row.remove(&row_1) {
                    pending.sort_by_key(|p| p.col_0);
                    let updated_row = updated_row_spans_element(e, &pending)?;
                    // Expand `<row/>` into `<row>...</row>` and insert cells.
                    let row_tag = String::from_utf8_lossy(e.name().as_ref()).into_owned();
                    let row_prefix_owned = element_prefix(e.name().as_ref())
                        .and_then(|p| std::str::from_utf8(p).ok())
                        .map(|s| s.to_string());
                    let row_prefix = row_prefix_owned.as_deref().or(sheet_prefix.as_deref());
                    if let Some(updated) = updated_row {
                        writer.write_event(Event::Start(updated))?;
                    } else {
                        writer.write_event(Event::Start(e.to_owned()))?;
                    }
                    write_inserted_cells(&mut writer, &pending, row_prefix)?;
                    writer.write_event(Event::End(BytesEnd::new(row_tag.as_str())))?;
                } else {
                    writer.write_event(Event::Empty(e.to_owned()))?;
                }
            }
            Event::End(ref e) if in_sheet_data && local_name(e.name().as_ref()) == b"row" => {
                if let Some(state) = row_state.take() {
                    let prefix = state.cell_prefix.as_deref().or(sheet_prefix.as_deref());
                    write_remaining_row_cells(&mut writer, &state.pending, state.next_idx, prefix)?;
                }
                in_cell = false;
                writer.write_event(Event::End(e.to_owned()))?;
            }

            // Inside a row that needs patching, intercept cell events.
            Event::Start(ref e)
                if in_sheet_data && row_state.is_some() && local_name(e.name().as_ref()) == b"c" =>
            {
                let state = row_state.as_mut().expect("row_state just checked");
                let (cell_ref, col_0) = parse_cell_ref_and_col(e)?;
                let cell_prefix_owned = element_prefix(e.name().as_ref())
                    .and_then(|p| std::str::from_utf8(p).ok())
                    .map(|s| s.to_string());
                if state.cell_prefix.is_none() {
                    state.cell_prefix = cell_prefix_owned.clone();
                }
                let cell_prefix = cell_prefix_owned.as_deref().or(sheet_prefix.as_deref());

                // Insert any pending patches that come before this cell.
                insert_pending_before_cell(&mut writer, state, col_0, cell_prefix)?;

                if let Some(patch) = take_patch_for_col(state, col_0) {
                    patch_existing_cell(&mut reader, &mut writer, e, &cell_ref, &patch)?;
                } else {
                    writer.write_event(Event::Start(e.to_owned()))?;
                    in_cell = true;
                }
            }
            Event::Empty(ref e)
                if in_sheet_data && row_state.is_some() && local_name(e.name().as_ref()) == b"c" =>
            {
                let state = row_state.as_mut().expect("row_state just checked");
                let (cell_ref, col_0) = parse_cell_ref_and_col(e)?;
                let cell_prefix_owned = element_prefix(e.name().as_ref())
                    .and_then(|p| std::str::from_utf8(p).ok())
                    .map(|s| s.to_string());
                if state.cell_prefix.is_none() {
                    state.cell_prefix = cell_prefix_owned.clone();
                }
                let cell_prefix = cell_prefix_owned.as_deref().or(sheet_prefix.as_deref());

                insert_pending_before_cell(&mut writer, state, col_0, cell_prefix)?;

                if let Some(patch) = take_patch_for_col(state, col_0) {
                    write_patched_cell(&mut writer, Some(e), &cell_ref, &patch, cell_prefix)?;
                } else {
                    writer.write_event(Event::Empty(e.to_owned()))?;
                }
            }
            Event::End(ref e)
                if in_sheet_data && row_state.is_some() && in_cell && local_name(e.name().as_ref()) == b"c" =>
            {
                in_cell = false;
                writer.write_event(Event::End(e.to_owned()))?;
            }
            // Ensure cells are emitted before any non-cell elements (e.g. extLst) in the row.
            Event::Start(ref e)
                if in_sheet_data
                    && row_state.is_some()
                    && !in_cell
                    && local_name(e.name().as_ref()) != b"c" =>
            {
                let state = row_state.as_mut().expect("row_state just checked");
                let prefix = state.cell_prefix.clone();
                let prefix = prefix.as_deref().or(sheet_prefix.as_deref());
                insert_pending_before_non_cell(&mut writer, state, prefix)?;
                writer.write_event(Event::Start(e.to_owned()))?;
            }
            Event::Empty(ref e)
                if in_sheet_data
                    && row_state.is_some()
                    && !in_cell
                    && local_name(e.name().as_ref()) != b"c" =>
            {
                let state = row_state.as_mut().expect("row_state just checked");
                let prefix = state.cell_prefix.clone();
                let prefix = prefix.as_deref().or(sheet_prefix.as_deref());
                insert_pending_before_non_cell(&mut writer, state, prefix)?;
                writer.write_event(Event::Empty(e.to_owned()))?;
            }

            // Default passthrough.
            ev => writer.write_event(ev.into_owned())?,
        }

        buf.clear();
    }

    if !saw_sheet_data {
        return Err(StreamingPatchError::MissingSheetData(
            worksheet_part.to_string(),
        ));
    }

    Ok(())
}

fn parse_row_number(e: &BytesStart<'_>) -> Result<u32, StreamingPatchError> {
    for attr in e.attributes() {
        let attr = attr?;
        if attr.key.as_ref() == b"r" {
            let v = attr.unescape_value()?.into_owned();
            return Ok(v.parse::<u32>().unwrap_or(0));
        }
    }
    Ok(0)
}

fn parse_cell_ref_and_col(e: &BytesStart<'_>) -> Result<(CellRef, u32), StreamingPatchError> {
    for attr in e.attributes() {
        let attr = attr?;
        if attr.key.as_ref() == b"r" {
            let a1 = attr.unescape_value()?.into_owned();
            let cell_ref =
                CellRef::from_a1(&a1).map_err(|_| StreamingPatchError::InvalidCellRef(a1))?;
            return Ok((cell_ref, cell_ref.col));
        }
    }
    // Malformed cell - treat as A1 so it at least serializes.
    Ok((CellRef::new(0, 0), 0))
}

fn spans_for_patches(patches: &[CellPatchInternal]) -> Option<(u32, u32)> {
    let mut iter = patches.iter();
    let first = iter.next()?;
    let mut min_col_1 = first.col_0 + 1;
    let mut max_col_1 = first.col_0 + 1;
    for patch in iter {
        let col_1 = patch.col_0 + 1;
        min_col_1 = min_col_1.min(col_1);
        max_col_1 = max_col_1.max(col_1);
    }
    Some((min_col_1, max_col_1))
}

fn parse_row_spans(spans: &str) -> Option<(u32, u32)> {
    let (start, end) = spans.split_once(':')?;
    let start = start.parse::<u32>().ok()?;
    let end = end.parse::<u32>().ok()?;
    Some((start, end))
}

fn format_row_spans(min_col_1: u32, max_col_1: u32) -> String {
    format!("{min_col_1}:{max_col_1}")
}

fn updated_row_spans_element(
    original: &BytesStart<'_>,
    patches: &[CellPatchInternal],
) -> Result<Option<BytesStart<'static>>, StreamingPatchError> {
    let Some((patch_min_col_1, patch_max_col_1)) = spans_for_patches(patches) else {
        return Ok(None);
    };

    let spans_attr = original
        .attributes()
        .flatten()
        .find(|a| a.key.as_ref() == b"spans")
        .and_then(|a| a.unescape_value().ok())
        .map(|v| v.into_owned());
    let Some(spans_attr) = spans_attr else {
        return Ok(None);
    };
    let Some((existing_min_col_1, existing_max_col_1)) = parse_row_spans(&spans_attr) else {
        return Ok(None);
    };

    let min_col_1 = existing_min_col_1.min(patch_min_col_1);
    let max_col_1 = existing_max_col_1.max(patch_max_col_1);
    if min_col_1 == existing_min_col_1 && max_col_1 == existing_max_col_1 {
        return Ok(None);
    }

    let spans = format_row_spans(min_col_1, max_col_1);
    let tag = String::from_utf8_lossy(original.name().as_ref()).into_owned();
    let mut out = BytesStart::new(tag.as_str());
    for attr in original.attributes() {
        let attr = attr?;
        if attr.key.as_ref() == b"spans" {
            out.push_attribute(("spans", spans.as_str()));
        } else {
            out.push_attribute((attr.key.as_ref(), attr.value.as_ref()));
        }
    }

    Ok(Some(out.into_owned()))
}

fn write_dimension_element<W: Write>(
    writer: &mut Writer<W>,
    prefix: Option<&str>,
    dimension_ref: &str,
) -> Result<(), StreamingPatchError> {
    let tag = prefixed_tag(prefix, "dimension");
    let mut dimension = BytesStart::new(tag.as_str());
    dimension.push_attribute(("ref", dimension_ref));
    writer.write_event(Event::Empty(dimension))?;
    Ok(())
}

fn write_pending_rows<W: Write>(
    writer: &mut Writer<W>,
    patches_by_row: &mut BTreeMap<u32, Vec<CellPatchInternal>>,
    prefix: Option<&str>,
) -> Result<(), StreamingPatchError> {
    while let Some((&row_1, _)) = patches_by_row.iter().next() {
        let pending = patches_by_row.remove(&row_1).unwrap_or_default();
        write_inserted_row(writer, row_1, &pending, prefix)?;
    }
    Ok(())
}

fn write_inserted_row<W: Write>(
    writer: &mut Writer<W>,
    row_1: u32,
    patches: &[CellPatchInternal],
    prefix: Option<&str>,
) -> Result<(), StreamingPatchError> {
    let row_tag = prefixed_tag(prefix, "row");
    let mut row = BytesStart::new(row_tag.as_str());
    let row_num = row_1.to_string();
    row.push_attribute(("r", row_num.as_str()));
    let spans = spans_for_patches(patches).map(|(min_col_1, max_col_1)| {
        format_row_spans(min_col_1, max_col_1)
    });
    if let Some(spans) = spans.as_deref() {
        row.push_attribute(("spans", spans));
    }
    writer.write_event(Event::Start(row))?;
    write_inserted_cells(writer, patches, prefix)?;
    writer.write_event(Event::End(BytesEnd::new(row_tag.as_str())))?;
    Ok(())
}

fn write_inserted_cells<W: Write>(
    writer: &mut Writer<W>,
    patches: &[CellPatchInternal],
    prefix: Option<&str>,
) -> Result<(), StreamingPatchError> {
    for patch in patches {
        let cell_ref = CellRef::new(patch.row_1 - 1, patch.col_0);
        write_patched_cell::<W>(writer, None, &cell_ref, patch, prefix)?;
    }
    Ok(())
}

fn write_remaining_row_cells<W: Write>(
    writer: &mut Writer<W>,
    pending: &[CellPatchInternal],
    next_idx: usize,
    prefix: Option<&str>,
) -> Result<(), StreamingPatchError> {
    if next_idx >= pending.len() {
        return Ok(());
    }
    for patch in &pending[next_idx..] {
        let cell_ref = CellRef::new(patch.row_1 - 1, patch.col_0);
        write_patched_cell::<W>(writer, None, &cell_ref, patch, prefix)?;
    }
    Ok(())
}

fn insert_pending_before_cell<W: Write>(
    writer: &mut Writer<W>,
    state: &mut RowState,
    col_0: u32,
    prefix: Option<&str>,
) -> Result<(), StreamingPatchError> {
    while let Some(patch) = state.pending.get(state.next_idx) {
        if patch.col_0 < col_0 {
            let cell_ref = CellRef::new(state.row_1 - 1, patch.col_0);
            write_patched_cell::<W>(writer, None, &cell_ref, patch, prefix)?;
            state.next_idx += 1;
        } else {
            break;
        }
    }
    Ok(())
}

fn insert_pending_before_non_cell<W: Write>(
    writer: &mut Writer<W>,
    state: &mut RowState,
    prefix: Option<&str>,
) -> Result<(), StreamingPatchError> {
    if state.next_idx >= state.pending.len() {
        return Ok(());
    }
    for patch in &state.pending[state.next_idx..] {
        let cell_ref = CellRef::new(state.row_1 - 1, patch.col_0);
        write_patched_cell::<W>(writer, None, &cell_ref, patch, prefix)?;
    }
    state.next_idx = state.pending.len();
    Ok(())
}

fn take_patch_for_col(state: &mut RowState, col_0: u32) -> Option<CellPatchInternal> {
    if state.next_idx >= state.pending.len() {
        return None;
    }
    let patch = state.pending.get(state.next_idx)?;
    if patch.col_0 == col_0 {
        let taken = patch.clone();
        state.next_idx += 1;
        Some(taken)
    } else {
        None
    }
}

fn patch_existing_cell<R: BufRead, W: Write>(
    reader: &mut Reader<R>,
    writer: &mut Writer<W>,
    cell_start: &BytesStart<'_>,
    cell_ref: &CellRef,
    patch: &CellPatchInternal,
) -> Result<(), StreamingPatchError> {
    let patch_formula = patch.formula.as_deref();
    let mut existing_t: Option<String> = None;
    let style_override = patch.xf_index;

    let cell_tag = String::from_utf8_lossy(cell_start.name().as_ref()).into_owned();
    let prefix = cell_tag.rsplit_once(':').map(|(p, _)| p);
    let f_tag = prefixed_tag(prefix, "f");
    let v_tag = prefixed_tag(prefix, "v");
    let is_tag = prefixed_tag(prefix, "is");
    let t_tag = prefixed_tag(prefix, "t");

    let mut c = BytesStart::new(cell_tag.as_str());
    let mut has_r = false;
    for attr in cell_start.attributes() {
        let attr = attr?;
        if attr.key.as_ref() == b"t" {
            existing_t = Some(attr.unescape_value()?.into_owned());
            continue;
        }
        if attr.key.as_ref() == b"s" && style_override.is_some() {
            continue;
        }
        if attr.key.as_ref() == b"r" {
            has_r = true;
        }
        c.push_attribute((attr.key.as_ref(), attr.value.as_ref()));
    }
    if !has_r {
        let a1 = cell_ref.to_a1();
        c.push_attribute(("r", a1.as_str()));
    }

    let (cell_t_owned, body_kind) = cell_representation(
        &patch.value,
        patch_formula,
        existing_t.as_deref(),
        patch.shared_string_idx,
    )?;

    if let Some(xf_index) = style_override {
        if xf_index != 0 {
            let xf = xf_index.to_string();
            c.push_attribute(("s", xf.as_str()));
        }
    }

    if let Some(t) = cell_t_owned.as_deref() {
        c.push_attribute(("t", t));
    }

    writer.write_event(Event::Start(c))?;

    let mut inner_buf = Vec::new();
    let mut inner_events: Vec<Event<'static>> = Vec::new();
    loop {
        let ev = reader.read_event_into(&mut inner_buf)?;
        match ev {
            Event::End(ref e) if local_name(e.name().as_ref()) == b"c" => break,
            Event::Eof => break,
            ev => inner_events.push(ev.into_owned()),
        }
        inner_buf.clear();
    }

    write_patched_cell_children(
        writer,
        &inner_events,
        patch_formula,
        &body_kind,
        &f_tag,
        &v_tag,
        &is_tag,
        &t_tag,
    )?;
    writer.write_event(Event::End(BytesEnd::new(cell_tag.as_str())))?;
    Ok(())
}

fn write_patched_cell_children<W: Write>(
    writer: &mut Writer<W>,
    inner_events: &[Event<'static>],
    patch_formula: Option<&str>,
    body_kind: &CellBodyKind,
    f_tag: &str,
    v_tag: &str,
    is_tag: &str,
    t_tag: &str,
) -> Result<(), StreamingPatchError> {
    let mut formula_written = patch_formula.is_none();
    let mut value_written = matches!(body_kind, CellBodyKind::None);
    let mut saw_formula = false;
    let mut saw_value = false;

    let mut idx = 0usize;
    while idx < inner_events.len() {
        match &inner_events[idx] {
            Event::Start(e) if local_name(e.name().as_ref()) == b"f" => {
                saw_formula = true;
                if !formula_written {
                    if let Some(formula) = patch_formula {
                        let detach_shared = should_detach_shared_formula(e, formula);
                        let tag = String::from_utf8_lossy(e.name().as_ref()).into_owned();
                        write_formula_element(writer, Some(e), formula, detach_shared, tag.as_str())?;
                        formula_written = true;
                    }
                }
                idx = skip_owned_subtree(inner_events, idx);
                continue;
            }
            Event::Empty(e) if local_name(e.name().as_ref()) == b"f" => {
                saw_formula = true;
                if !formula_written {
                    if let Some(formula) = patch_formula {
                        let detach_shared = should_detach_shared_formula(e, formula);
                        let tag = String::from_utf8_lossy(e.name().as_ref()).into_owned();
                        write_formula_element(writer, Some(e), formula, detach_shared, tag.as_str())?;
                        formula_written = true;
                    }
                }
                idx += 1;
                continue;
            }
            Event::Start(e)
                if local_name(e.name().as_ref()) == b"v" || local_name(e.name().as_ref()) == b"is" =>
            {
                saw_value = true;

                if !formula_written {
                    if let Some(formula) = patch_formula {
                        // Original cell has no <f> before the value; insert one.
                        write_formula_element(writer, None, formula, false, f_tag)?;
                        formula_written = true;
                    }
                }
                if !value_written {
                    write_value_element(writer, body_kind, v_tag, is_tag, t_tag)?;
                    value_written = true;
                }

                idx = skip_owned_subtree(inner_events, idx);
                continue;
            }
            Event::Empty(e)
                if local_name(e.name().as_ref()) == b"v" || local_name(e.name().as_ref()) == b"is" =>
            {
                saw_value = true;

                if !formula_written {
                    if let Some(formula) = patch_formula {
                        write_formula_element(writer, None, formula, false, f_tag)?;
                        formula_written = true;
                    }
                }
                if !value_written {
                    write_value_element(writer, body_kind, v_tag, is_tag, t_tag)?;
                    value_written = true;
                }

                idx += 1;
                continue;
            }
            ev => {
                if !formula_written && !saw_formula {
                    if let Some(formula) = patch_formula {
                        write_formula_element(writer, None, formula, false, f_tag)?;
                        formula_written = true;
                    }
                }
                if !value_written && !saw_value {
                    write_value_element(writer, body_kind, v_tag, is_tag, t_tag)?;
                    value_written = true;
                }
                writer.write_event(ev.clone())?;
            }
        }
        idx += 1;
    }

    if !formula_written {
        if let Some(formula) = patch_formula {
            write_formula_element(writer, None, formula, false, f_tag)?;
        }
    }
    if !value_written {
        write_value_element(writer, body_kind, v_tag, is_tag, t_tag)?;
    }

    Ok(())
}

fn skip_owned_subtree(events: &[Event<'static>], mut idx: usize) -> usize {
    match &events[idx] {
        Event::Start(_) => {
            let mut depth = 1usize;
            idx += 1;
            while idx < events.len() {
                match &events[idx] {
                    Event::Start(_) => depth += 1,
                    Event::End(_) => {
                        depth = depth.saturating_sub(1);
                        if depth == 0 {
                            idx += 1;
                            break;
                        }
                    }
                    Event::Empty(_) => {}
                    _ => {}
                }
                idx += 1;
            }
            idx
        }
        _ => idx + 1,
    }
}

fn write_formula_element<W: Write>(
    writer: &mut Writer<W>,
    original: Option<&BytesStart<'_>>,
    formula: &str,
    detach_shared: bool,
    tag_name: &str,
) -> Result<(), StreamingPatchError> {
    let formula_display = crate::formula_text::normalize_display_formula(formula);
    let file_formula = crate::formula_text::add_xlfn_prefixes(&formula_display);

    let mut f = BytesStart::new(tag_name);
    if let Some(orig) = original {
        for attr in orig.attributes() {
            let attr = attr?;
            if detach_shared && matches!(attr.key.as_ref(), b"t" | b"ref" | b"si") {
                continue;
            }
            f.push_attribute((attr.key.as_ref(), attr.value.as_ref()));
        }
    }

    if file_formula.is_empty() {
        writer.write_event(Event::Empty(f))?;
    } else {
        writer.write_event(Event::Start(f))?;
        writer.write_event(Event::Text(BytesText::new(&file_formula)))?;
        writer.write_event(Event::End(BytesEnd::new(tag_name)))?;
    }
    Ok(())
}

fn write_value_element<W: Write>(
    writer: &mut Writer<W>,
    body_kind: &CellBodyKind,
    v_tag: &str,
    is_tag: &str,
    t_tag: &str,
) -> Result<(), StreamingPatchError> {
    match body_kind {
        CellBodyKind::V(text) => {
            writer.write_event(Event::Start(BytesStart::new(v_tag)))?;
            writer.write_event(Event::Text(BytesText::new(text)))?;
            writer.write_event(Event::End(BytesEnd::new(v_tag)))?;
        }
        CellBodyKind::InlineStr(text) => {
            writer.write_event(Event::Start(BytesStart::new(is_tag)))?;
            write_shared_string_t(writer, t_tag, text)?;
            writer.write_event(Event::End(BytesEnd::new(is_tag)))?;
        }
        CellBodyKind::InlineRich(rich) => {
            let prefix = t_tag.rsplit_once(':').map(|(p, _)| p);
            let r_tag = prefixed_tag(prefix, "r");
            let rpr_tag = prefixed_tag(prefix, "rPr");

            writer.write_event(Event::Start(BytesStart::new(is_tag)))?;
            if rich.runs.is_empty() {
                write_shared_string_t(writer, t_tag, &rich.text)?;
            } else {
                for run in &rich.runs {
                    writer.write_event(Event::Start(BytesStart::new(r_tag.as_str())))?;

                    if !run.style.is_empty() {
                        writer.write_event(Event::Start(BytesStart::new(rpr_tag.as_str())))?;
                        write_shared_string_rpr(writer, prefix, &run.style)?;
                        writer.write_event(Event::End(BytesEnd::new(rpr_tag.as_str())))?;
                    }

                    let segment = rich.slice_run_text(run);
                    write_shared_string_t(writer, t_tag, segment)?;

                    writer.write_event(Event::End(BytesEnd::new(r_tag.as_str())))?;
                }
            }
            writer.write_event(Event::End(BytesEnd::new(is_tag)))?;
        }
        CellBodyKind::None => {}
    }

    Ok(())
}

fn should_detach_shared_formula(f: &BytesStart<'_>, patch_formula: &str) -> bool {
    let trimmed = patch_formula.trim();
    let stripped = trimmed.strip_prefix('=').unwrap_or(trimmed).trim();
    if stripped.is_empty() {
        return false;
    }

    let mut is_shared = false;
    let mut has_ref = false;
    for attr in f.attributes().flatten() {
        match attr.key.as_ref() {
            b"t" if attr.value.as_ref() == b"shared" => is_shared = true,
            b"ref" => has_ref = true,
            _ => {}
        }
    }

    is_shared && !has_ref
}

fn write_patched_cell<W: Write>(
    writer: &mut Writer<W>,
    original: Option<&BytesStart<'_>>,
    cell_ref: &CellRef,
    patch: &CellPatchInternal,
    prefix: Option<&str>,
) -> Result<(), StreamingPatchError> {
    let patch_formula = patch.formula.as_deref();
    let mut existing_t: Option<String> = None;
    let shared_string_idx = patch.shared_string_idx;

    let cell_tag_owned = match original {
        Some(orig) => String::from_utf8_lossy(orig.name().as_ref()).into_owned(),
        None => prefixed_tag(prefix, "c"),
    };
    let cell_prefix = cell_tag_owned.rsplit_once(':').map(|(p, _)| p);
    let formula_tag = prefixed_tag(cell_prefix, "f");
    let v_tag = prefixed_tag(cell_prefix, "v");
    let is_tag = prefixed_tag(cell_prefix, "is");
    let t_tag = prefixed_tag(cell_prefix, "t");

    let mut c = BytesStart::new(cell_tag_owned.as_str());
    let mut has_r = false;
    let style_override = patch.xf_index;

    if let Some(orig) = original {
        for attr in orig.attributes() {
            let attr = attr?;
            if attr.key.as_ref() == b"t" {
                existing_t = Some(attr.unescape_value()?.into_owned());
                continue;
            }
            if attr.key.as_ref() == b"s" && style_override.is_some() {
                continue;
            }
            if attr.key.as_ref() == b"r" {
                has_r = true;
            }
            c.push_attribute((attr.key.as_ref(), attr.value.as_ref()));
        }
    } else {
        let a1 = cell_ref.to_a1();
        c.push_attribute(("r", a1.as_str()));
        has_r = true;
    }
    if !has_r {
        let a1 = cell_ref.to_a1();
        c.push_attribute(("r", a1.as_str()));
    }

    let (cell_t_owned, body_kind) = cell_representation(
        &patch.value,
        patch_formula,
        existing_t.as_deref(),
        shared_string_idx,
    )?;

    if let Some(xf_index) = style_override {
        if xf_index != 0 {
            let xf = xf_index.to_string();
            c.push_attribute(("s", xf.as_str()));
        }
    }

    if let Some(t) = cell_t_owned.as_deref() {
        c.push_attribute(("t", t));
    }

    writer.write_event(Event::Start(c))?;

    if let Some(formula) = patch_formula {
        write_formula_element(writer, None, formula, false, &formula_tag)?;
    }

    write_value_element(writer, &body_kind, &v_tag, &is_tag, &t_tag)?;

    writer.write_event(Event::End(BytesEnd::new(cell_tag_owned.as_str())))?;
    Ok(())
}

#[derive(Debug, Clone)]
enum CellBodyKind {
    None,
    V(String),
    InlineStr(String),
    InlineRich(RichText),
}

fn cell_representation(
    value: &CellValue,
    formula: Option<&str>,
    existing_t: Option<&str>,
    shared_string_idx: Option<u32>,
) -> Result<(Option<String>, CellBodyKind), StreamingPatchError> {
    match value {
        CellValue::Empty => Ok((None, CellBodyKind::None)),
        CellValue::Number(n) => Ok((None, CellBodyKind::V(n.to_string()))),
        CellValue::Boolean(b) => Ok((
            Some("b".to_string()),
            CellBodyKind::V(if *b { "1" } else { "0" }.to_string()),
        )),
        CellValue::Error(err) => Ok((
            Some("e".to_string()),
            CellBodyKind::V(err.as_str().to_string()),
        )),
        CellValue::String(s) => {
            if let Some(existing_t) = existing_t {
                if should_preserve_unknown_t(existing_t) {
                    return Ok((Some(existing_t.to_string()), CellBodyKind::V(s.clone())));
                }
                match existing_t {
                    "inlineStr" => {
                        return Ok((Some("inlineStr".to_string()), CellBodyKind::InlineStr(s.clone())));
                    }
                    "str" => {
                        return Ok((Some("str".to_string()), CellBodyKind::V(s.clone())));
                    }
                    _ => {}
                }
            }

            if let Some(idx) = shared_string_idx {
                return Ok((Some("s".to_string()), CellBodyKind::V(idx.to_string())));
            }

            if formula.is_some() {
                Ok((Some("str".to_string()), CellBodyKind::V(s.clone())))
            } else {
                Ok((Some("inlineStr".to_string()), CellBodyKind::InlineStr(s.clone())))
            }
        }
        CellValue::RichText(rich) => {
            if let Some(existing_t) = existing_t {
                if should_preserve_unknown_t(existing_t) {
                    return Ok((Some(existing_t.to_string()), CellBodyKind::V(rich.text.clone())));
                }
                if existing_t == "inlineStr" {
                    return Ok((
                        Some("inlineStr".to_string()),
                        CellBodyKind::InlineRich(rich.clone()),
                    ));
                }
            }

            if let Some(idx) = shared_string_idx {
                return Ok((Some("s".to_string()), CellBodyKind::V(idx.to_string())));
            }

            Ok((
                Some("inlineStr".to_string()),
                CellBodyKind::InlineRich(rich.clone()),
            ))
        }
        other => Err(StreamingPatchError::UnsupportedCellValue(other.clone())),
    }
}

fn should_preserve_unknown_t(t: &str) -> bool {
    // Preserve less-common or unknown SpreadsheetML cell types (e.g. `t="d"`). When patching
    // string cells, rewriting these as `inlineStr`/`str` can change semantics or cause Excel to
    // re-interpret values. Keep the original `t` and write the patched value into `<v>` instead.
    !matches!(t, "s" | "b" | "e" | "n" | "str" | "inlineStr")
}

fn needs_space_preserve(s: &str) -> bool {
    s.starts_with(char::is_whitespace) || s.ends_with(char::is_whitespace)
}

fn local_name(name: &[u8]) -> &[u8] {
    match name.iter().rposition(|b| *b == b':') {
        Some(idx) => &name[idx + 1..],
        None => name,
    }
}

fn element_prefix(name: &[u8]) -> Option<&[u8]> {
    name.iter().rposition(|b| *b == b':').map(|idx| &name[..idx])
}

fn prefixed_tag(prefix: Option<&str>, local: &str) -> String {
    match prefix {
        Some(prefix) => format!("{prefix}:{local}"),
        None => local.to_string(),
    }
}

fn worksheet_has_default_spreadsheetml_ns(e: &BytesStart<'_>) -> Result<bool, StreamingPatchError> {
    for attr in e.attributes() {
        let attr = attr?;
        if attr.key.as_ref() == b"xmlns" && attr.value.as_ref() == SPREADSHEETML_NS.as_bytes() {
            return Ok(true);
        }
    }
    Ok(false)
}

#[derive(Debug, Clone, Copy)]
struct PatchBounds {
    min_row_0: u32,
    max_row_0: u32,
    min_col_0: u32,
    max_col_0: u32,
}

fn bounds_for_patches(patches: &[WorksheetCellPatch]) -> Option<PatchBounds> {
    let mut iter = patches.iter();
    let first = iter.next()?;
    let mut min_row_0 = first.cell.row;
    let mut max_row_0 = first.cell.row;
    let mut min_col_0 = first.cell.col;
    let mut max_col_0 = first.cell.col;

    for patch in iter {
        min_row_0 = min_row_0.min(patch.cell.row);
        max_row_0 = max_row_0.max(patch.cell.row);
        min_col_0 = min_col_0.min(patch.cell.col);
        max_col_0 = max_col_0.max(patch.cell.col);
    }

    Some(PatchBounds {
        min_row_0,
        max_row_0,
        min_col_0,
        max_col_0,
    })
}

fn union_bounds(a: Option<PatchBounds>, b: Option<PatchBounds>) -> Option<PatchBounds> {
    match (a, b) {
        (None, None) => None,
        (Some(a), None) | (None, Some(a)) => Some(a),
        (Some(a), Some(b)) => Some(PatchBounds {
            min_row_0: a.min_row_0.min(b.min_row_0),
            max_row_0: a.max_row_0.max(b.max_row_0),
            min_col_0: a.min_col_0.min(b.min_col_0),
            max_col_0: a.max_col_0.max(b.max_col_0),
        }),
    }
}

fn bounds_to_dimension_ref(bounds: PatchBounds) -> String {
    let start = CellRef::new(bounds.min_row_0, bounds.min_col_0);
    let end = CellRef::new(bounds.max_row_0, bounds.max_col_0);
    if start == end {
        start.to_a1()
    } else {
        format!("{}:{}", start.to_a1(), end.to_a1())
    }
}

fn updated_dimension_element(
    original: &BytesStart<'_>,
    bounds: PatchBounds,
) -> Result<BytesStart<'static>, StreamingPatchError> {
    let original_ref = original
        .attributes()
        .flatten()
        .find(|a| a.key.as_ref() == b"ref")
        .and_then(|a| a.unescape_value().ok())
        .map(|v| v.into_owned());

    let mut min_row_0 = bounds.min_row_0;
    let mut max_row_0 = bounds.max_row_0;
    let mut min_col_0 = bounds.min_col_0;
    let mut max_col_0 = bounds.max_col_0;

    if let Some(existing) = original_ref.as_deref() {
        if let Some((start, end)) = parse_dimension_ref(existing) {
            min_row_0 = min_row_0.min(start.row);
            max_row_0 = max_row_0.max(end.row);
            min_col_0 = min_col_0.min(start.col);
            max_col_0 = max_col_0.max(end.col);
        }
    }

    let start = CellRef::new(min_row_0, min_col_0);
    let end = CellRef::new(max_row_0, max_col_0);
    let updated_ref = if start == end {
        start.to_a1()
    } else {
        format!("{}:{}", start.to_a1(), end.to_a1())
    };

    // Preserve attribute ordering where possible by rewriting `ref` in-place.
    let tag = String::from_utf8_lossy(original.name().as_ref()).into_owned();
    let mut out = BytesStart::new(tag.as_str());
    for attr in original.attributes() {
        let attr = attr?;
        if attr.key.as_ref() == b"ref" {
            out.push_attribute(("ref", updated_ref.as_str()));
        } else {
            out.push_attribute((attr.key.as_ref(), attr.value.as_ref()));
        }
    }

    if original
        .attributes()
        .flatten()
        .all(|a| a.key.as_ref() != b"ref")
    {
        out.push_attribute(("ref", updated_ref.as_str()));
    }

    Ok(out.into_owned())
}

fn parse_dimension_ref(s: &str) -> Option<(CellRef, CellRef)> {
    let mut parts = s.split(':');
    let start = parts.next()?;
    let start = CellRef::from_a1(start).ok()?;
    let end = parts
        .next()
        .and_then(|p| CellRef::from_a1(p).ok())
        .unwrap_or(start);
    Some((start, end))
}
