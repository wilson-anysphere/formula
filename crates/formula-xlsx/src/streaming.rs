use std::borrow::Cow;
use std::collections::{BTreeMap, BTreeSet, HashMap, HashSet};
use std::io::{BufRead, BufReader, Read, Seek, Write};

use formula_model::rich_text::RichText;
use formula_model::{CellRef, CellValue, ColProperties, ErrorValue, StyleTable};
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
use crate::zip_util::{open_zip_part, read_zip_file_bytes_with_limit, DEFAULT_MAX_ZIP_PART_BYTES};
use crate::RecalcPolicy;
use crate::WorkbookKind;
use crate::{parse_workbook_sheets, CellPatch, WorkbookCellPatches};

const REL_TYPE_STYLES: &str =
    "http://schemas.openxmlformats.org/officeDocument/2006/relationships/styles";
const REL_TYPE_SHARED_STRINGS: &str =
    "http://schemas.openxmlformats.org/officeDocument/2006/relationships/sharedStrings";

const SPREADSHEETML_NS: &str = "http://schemas.openxmlformats.org/spreadsheetml/2006/main";

fn canonicalize_zip_entry_name<'a>(name: &'a str) -> Cow<'a, str> {
    // ZIP entry names in valid XLSX/XLSM packages should not start with `/` and should use `/`
    // separators. Some producers emit non-canonical names with a leading slash and/or Windows-style
    // separators (`\`). Normalize those forms for matching against patch targets.
    let trimmed = name.trim_start_matches(|c| matches!(c, '/' | '\\'));
    if trimmed.contains('\\') {
        Cow::Owned(trimmed.replace('\\', "/"))
    } else {
        Cow::Borrowed(trimmed)
    }
}

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

    /// Optional cell `vm` attribute override.
    ///
    /// SpreadsheetML uses `c/@vm` for RichData-backed cell content (e.g. images-in-cell).
    ///
    /// - `None`: preserve the existing attribute when patching an existing cell (and omit it when
    ///   inserting a new cell).
    /// - `Some(Some(n))`: set/overwrite `vm="n"`.
    /// - `Some(None)`: remove the attribute.
    pub vm: Option<Option<u32>>,

    /// Optional cell `cm` attribute override.
    ///
    /// Some RichData-backed cell content also requires `c/@cm`.
    ///
    /// - `None`: preserve the existing attribute when patching an existing cell (and omit it when
    ///   inserting a new cell).
    /// - `Some(Some(n))`: set/overwrite `cm="n"`.
    /// - `Some(None)`: remove the attribute.
    pub cm: Option<Option<u32>>,
}

/// Override behavior for arbitrary (non-worksheet) OPC parts while streaming-patching.
///
/// This is primarily intended for losslessly updating "sidecar" XML parts (e.g.
/// `xl/formula/power-query.xml`) without inflating the entire [`crate::XlsxPackage`] in memory.
#[derive(Debug, Clone, PartialEq, Eq)]
pub enum PartOverride {
    /// Write these bytes to the output ZIP for this entry.
    ///
    /// If the part exists in the input, it will be replaced in-place. If it does not exist, it
    /// will be appended to the output ZIP after all input entries have been copied.
    Replace(Vec<u8>),
    /// Remove the part from the output ZIP (skip copying it if it exists).
    Remove,
    /// Ensure the part exists with these bytes.
    ///
    /// Behaves like [`PartOverride::Replace`] when the part is present in the input, and appends
    /// the part when it is missing.
    Add(Vec<u8>),
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
            vm: None,
            cm: None,
        }
    }

    pub fn with_xf_index(mut self, xf_index: Option<u32>) -> Self {
        self.xf_index = xf_index;
        self
    }

    pub fn with_vm(mut self, vm: Option<Option<u32>>) -> Self {
        self.vm = vm;
        self
    }

    pub fn with_cm(mut self, cm: Option<Option<u32>>) -> Self {
        self.cm = cm;
        self
    }

    pub fn set_vm(self, vm: u32) -> Self {
        self.with_vm(Some(Some(vm)))
    }

    pub fn clear_vm(self) -> Self {
        self.with_vm(Some(None))
    }

    pub fn set_cm(self, cm: u32) -> Self {
        self.with_cm(Some(Some(cm)))
    }

    pub fn clear_cm(self) -> Self {
        self.with_cm(Some(None))
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
    patch_xlsx_streaming_with_recalc_policy(input, output, cell_patches, RecalcPolicy::default())
}

/// Streaming XLSX/XLSM patcher with a configurable [`RecalcPolicy`].
///
/// `policy_on_formula_change` is applied **only** when the patch set changes formulas (including
/// removing formulas). When no formulas change, [`RecalcPolicy::PRESERVE`] is used regardless of
/// the provided policy.
pub fn patch_xlsx_streaming_with_recalc_policy<R: Read + Seek, W: Write + Seek>(
    input: R,
    output: W,
    cell_patches: &[WorksheetCellPatch],
    policy_on_formula_change: RecalcPolicy,
) -> Result<(), StreamingPatchError> {
    let mut patches_by_part: HashMap<String, Vec<WorksheetCellPatch>> = HashMap::new();
    for patch in cell_patches {
        // ZIP entry names in valid XLSX/XLSM packages should not start with `/` (or `\`), but
        // tolerate callers that include it by normalizing the patch target part name.
        let worksheet_part = patch
            .worksheet_part
            .trim_start_matches(|c| c == '/' || c == '\\')
            .to_string();
        let mut patch = patch.clone();
        patch.worksheet_part = worksheet_part.clone();
        patches_by_part
            .entry(worksheet_part)
            .or_default()
            .push(patch);
    }
    let mut archive = ZipArchive::new(input)?;
    let part_names = list_zip_part_names(&mut archive)?;

    // Remap patch targets to the exact ZIP entry names present in the container when possible.
    // This keeps downstream lookups (which often key by `ZipFile::name()`) working even when a
    // producer used non-canonical naming (case differences, `\` separators, leading separators,
    // percent-encoding).
    if !patches_by_part.is_empty() {
        let mut remapped: HashMap<String, Vec<WorksheetCellPatch>> = HashMap::new();
        for (candidate, patches) in std::mem::take(&mut patches_by_part) {
            let resolved = find_zip_part_name(&part_names, &candidate).unwrap_or(candidate);
            let entry = remapped.entry(resolved.clone()).or_default();
            for mut patch in patches {
                patch.worksheet_part = resolved.clone();
                entry.push(patch);
            }
        }
        patches_by_part = remapped;
    }

    // Deterministic patching within a worksheet.
    for patches in patches_by_part.values_mut() {
        patches.sort_by_key(|p| (p.cell.row, p.cell.col));
    }

    // `vm` (cell value-metadata) is preserved by default for fidelity. Callers can explicitly
    // override/clear it via `WorksheetCellPatch`.
    let mut formula_changed = cell_patches
        .iter()
        .any(|p| formula_is_material(p.formula.as_deref()));
    if !formula_changed {
        formula_changed =
            streaming_patches_remove_existing_formulas(&mut archive, &patches_by_part)?;
    }
    let recalc_policy = if formula_changed {
        policy_on_formula_change
    } else {
        RecalcPolicy::PRESERVE
    };
    patch_xlsx_streaming_with_archive(
        &mut archive,
        output,
        &patches_by_part,
        &HashMap::new(),
        &HashMap::new(),
        &HashMap::new(),
        &HashMap::new(),
        recalc_policy,
    )?;
    Ok(())
}

/// Remove macro-related parts from an XLSX/XLSM archive streamingly.
///
/// This matches the semantics of [`crate::XlsxPackage::remove_vba_project`], but avoids inflating
/// every ZIP entry into memory. Unchanged parts are preserved byte-for-byte via `raw_copy_file`.
///
/// This is used when saving a macro-enabled workbook (`.xlsm`) as `.xlsx`.
pub fn strip_vba_project_streaming<R: Read + Seek, W: Write + Seek>(
    input: R,
    output: W,
) -> Result<(), StreamingPatchError> {
    strip_vba_project_streaming_with_kind(input, output, WorkbookKind::Workbook)
}

/// Remove macro-related parts from an XLSX/XLSM archive streamingly, rewriting the workbook
/// "main" content type in `[Content_Types].xml` based on `target_kind`.
pub fn strip_vba_project_streaming_with_kind<R: Read + Seek, W: Write + Seek>(
    input: R,
    output: W,
    target_kind: WorkbookKind,
) -> Result<(), StreamingPatchError> {
    let mut archive = ZipArchive::new(input)?;
    macro_strip_streaming::strip_vba_project_streaming_with_archive(
        &mut archive,
        output,
        target_kind,
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
    patch_xlsx_streaming_workbook_cell_patches_with_recalc_policy(
        input,
        output,
        patches,
        RecalcPolicy::default(),
    )
}

/// Apply [`WorkbookCellPatches`] (the part-preserving cell patch DSL) using the streaming ZIP
/// rewriter with a configurable [`RecalcPolicy`].
///
/// `policy_on_formula_change` is applied **only** when the patch set changes formulas (including
/// removing formulas). When no formulas change, [`RecalcPolicy::PRESERVE`] is used regardless of
/// the provided policy.
pub fn patch_xlsx_streaming_workbook_cell_patches_with_recalc_policy<
    R: Read + Seek,
    W: Write + Seek,
>(
    input: R,
    output: W,
    patches: &WorkbookCellPatches,
    policy_on_formula_change: RecalcPolicy,
) -> Result<(), StreamingPatchError> {
    if patches.is_empty() {
        return patch_xlsx_streaming_with_recalc_policy(
            input,
            output,
            &[],
            policy_on_formula_change,
        );
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
    let part_names = list_zip_part_names(&mut archive)?;

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
        if rel
            .target_mode
            .as_deref()
            .is_some_and(|mode| mode.trim().eq_ignore_ascii_case("External"))
        {
            // Workbook relationships may reference external resources (typically hyperlinks).
            // These do not correspond to OPC part names and should not participate in internal
            // target resolution.
            continue;
        }
        rel_targets.insert(rel.id, resolve_target("xl/workbook.xml", &rel.target));
    }

    let mut patches_by_part: HashMap<String, Vec<WorksheetCellPatch>> = HashMap::new();
    let mut col_properties_by_part: HashMap<String, BTreeMap<u32, ColProperties>> = HashMap::new();
    let mut saw_formula_patch = false;
    for (sheet_selector, sheet_patches) in patches.sheets() {
        if sheet_patches.is_empty() {
            continue;
        }
        let worksheet_part =
            resolve_worksheet_part_for_selector(sheet_selector, &workbook_sheets, &rel_targets)?;
        let worksheet_part = find_zip_part_name(&part_names, &worksheet_part)
            .unwrap_or_else(|| worksheet_part.clone());

        if let Some(cols) = sheet_patches.col_properties() {
            col_properties_by_part.insert(worksheet_part.clone(), cols.clone());
        }

        for (cell_ref, patch) in sheet_patches.iter() {
            let (value, formula) = match patch {
                CellPatch::Clear { .. } => (CellValue::Empty, None),
                CellPatch::Set { value, formula, .. } => (value.clone(), formula.clone()),
            };
            saw_formula_patch |= formula_is_material(formula.as_deref());
            let xf_index = patch.style_index();
            patches_by_part
                .entry(worksheet_part.clone())
                .or_default()
                .push(
                    WorksheetCellPatch::new(worksheet_part.clone(), cell_ref, value, formula)
                        .with_xf_index(xf_index)
                        .with_vm(patch.vm_override())
                        .with_cm(patch.cm_override()),
                );
        }
    }

    for patches in patches_by_part.values_mut() {
        patches.sort_by_key(|p| (p.cell.row, p.cell.col));
    }

    let mut formula_changed = saw_formula_patch;
    if !formula_changed {
        formula_changed =
            streaming_patches_remove_existing_formulas(&mut archive, &patches_by_part)?;
    }
    let recalc_policy = if formula_changed {
        policy_on_formula_change
    } else {
        RecalcPolicy::PRESERVE
    };

    patch_xlsx_streaming_with_archive(
        &mut archive,
        output,
        &patches_by_part,
        &col_properties_by_part,
        &pre_read_parts,
        &HashMap::new(),
        &HashMap::new(),
        recalc_policy,
    )?;
    Ok(())
}

/// Apply [`WorkbookCellPatches`] using the streaming ZIP rewriter, plus arbitrary part overrides.
///
/// This variant supports replacing/adding/removing non-worksheet parts (e.g.
/// `xl/formula/power-query.xml`) without inflating the entire workbook package into an
/// [`crate::XlsxPackage`].
pub fn patch_xlsx_streaming_workbook_cell_patches_with_part_overrides<
    R: Read + Seek,
    W: Write + Seek,
>(
    input: R,
    output: W,
    patches: &WorkbookCellPatches,
    part_overrides: &HashMap<String, PartOverride>,
) -> Result<(), StreamingPatchError> {
    patch_xlsx_streaming_workbook_cell_patches_with_part_overrides_and_recalc_policy(
        input,
        output,
        patches,
        part_overrides,
        RecalcPolicy::default(),
    )
}

/// Apply [`WorkbookCellPatches`] using the streaming ZIP rewriter, plus arbitrary part overrides,
/// with a configurable [`RecalcPolicy`].
///
/// `policy_on_formula_change` is applied **only** when the patch set changes formulas (including
/// removing formulas). When no formulas change, [`RecalcPolicy::PRESERVE`] is used regardless of
/// the provided policy.
pub fn patch_xlsx_streaming_workbook_cell_patches_with_part_overrides_and_recalc_policy<
    R: Read + Seek,
    W: Write + Seek,
>(
    input: R,
    output: W,
    patches: &WorkbookCellPatches,
    part_overrides: &HashMap<String, PartOverride>,
    policy_on_formula_change: RecalcPolicy,
) -> Result<(), StreamingPatchError> {
    if patches.is_empty() {
        let mut archive = ZipArchive::new(input)?;
        patch_xlsx_streaming_with_archive(
            &mut archive,
            output,
            &HashMap::new(),
            &HashMap::new(),
            &HashMap::new(),
            &HashMap::new(),
            part_overrides,
            RecalcPolicy::PRESERVE,
        )?;
        return Ok(());
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
    let part_names = list_zip_part_names(&mut archive)?;

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
        if rel
            .target_mode
            .as_deref()
            .is_some_and(|mode| mode.trim().eq_ignore_ascii_case("External"))
        {
            continue;
        }
        rel_targets.insert(rel.id, resolve_target("xl/workbook.xml", &rel.target));
    }

    let mut patches_by_part: HashMap<String, Vec<WorksheetCellPatch>> = HashMap::new();
    let mut col_properties_by_part: HashMap<String, BTreeMap<u32, ColProperties>> = HashMap::new();
    let mut saw_formula_patch = false;
    for (sheet_selector, sheet_patches) in patches.sheets() {
        if sheet_patches.is_empty() {
            continue;
        }
        let worksheet_part =
            resolve_worksheet_part_for_selector(sheet_selector, &workbook_sheets, &rel_targets)?;
        let worksheet_part = find_zip_part_name(&part_names, &worksheet_part)
            .unwrap_or_else(|| worksheet_part.clone());

        if let Some(cols) = sheet_patches.col_properties() {
            col_properties_by_part.insert(worksheet_part.clone(), cols.clone());
        }

        for (cell_ref, patch) in sheet_patches.iter() {
            let (value, formula) = match patch {
                CellPatch::Clear { .. } => (CellValue::Empty, None),
                CellPatch::Set { value, formula, .. } => (value.clone(), formula.clone()),
            };
            saw_formula_patch |= formula_is_material(formula.as_deref());
            let xf_index = patch.style_index();
            patches_by_part
                .entry(worksheet_part.clone())
                .or_default()
                .push(
                    WorksheetCellPatch::new(worksheet_part.clone(), cell_ref, value, formula)
                        .with_xf_index(xf_index)
                        .with_vm(patch.vm_override())
                        .with_cm(patch.cm_override()),
                );
        }
    }

    for patches in patches_by_part.values_mut() {
        patches.sort_by_key(|p| (p.cell.row, p.cell.col));
    }

    let mut formula_changed = saw_formula_patch;
    if !formula_changed {
        formula_changed =
            streaming_patches_remove_existing_formulas(&mut archive, &patches_by_part)?;
    }
    let recalc_policy = if formula_changed {
        policy_on_formula_change
    } else {
        RecalcPolicy::PRESERVE
    };

    patch_xlsx_streaming_with_archive(
        &mut archive,
        output,
        &patches_by_part,
        &col_properties_by_part,
        &pre_read_parts,
        &HashMap::new(),
        part_overrides,
        recalc_policy,
    )?;
    Ok(())
}

mod macro_strip_streaming {
    use std::collections::{BTreeMap, BTreeSet, HashMap, VecDeque};
    use std::io::{Read, Seek, Write};

    use quick_xml::events::{BytesStart, Event};
    use quick_xml::{Reader as XmlReader, Writer as XmlWriter};
    use zip::write::FileOptions;
    use zip::{CompressionMethod, ZipArchive, ZipWriter};

    use crate::WorkbookKind;

    use super::{read_zip_part, StreamingPatchError};

    const CUSTOM_UI_REL_TYPES: [&str; 2] = [
        "http://schemas.microsoft.com/office/2006/relationships/ui/extensibility",
        "http://schemas.microsoft.com/office/2007/relationships/ui/extensibility",
    ];

    const RELATIONSHIPS_NS: &[u8] =
        b"http://schemas.openxmlformats.org/officeDocument/2006/relationships";

    fn canonical_part_name(name: &str) -> String {
        // Normalize producer bugs so macro stripping is robust:
        // - strip leading separators (`/` or `\`, including percent-encoded)
        // - normalize `\` to `/`
        // - ASCII-lowercase
        // - percent-decode valid `%xx` sequences
        //
        // This matches `zip_part_names_equivalent`.
        String::from_utf8_lossy(&crate::zip_util::zip_part_name_lookup_key(name)).into_owned()
    }
    fn find_part_name(part_names: &BTreeSet<String>, candidate: &str) -> Option<String> {
        let candidate = canonical_part_name(candidate);
        part_names.get(&candidate).cloned()
    }
    pub(super) fn strip_vba_project_streaming_with_archive<R: Read + Seek, W: Write + Seek>(
        archive: &mut ZipArchive<R>,
        output: W,
        target_kind: WorkbookKind,
    ) -> Result<(), StreamingPatchError> {
        let part_names = list_part_names(archive)?;

        // Cache of XML parts we need to parse/patch while planning.
        let mut read_cache: HashMap<String, Vec<u8>> = HashMap::new();
        let mut delete_parts = compute_macro_delete_set(archive, &part_names, &mut read_cache)?;

        let mut updated_parts: HashMap<String, Vec<u8>> = HashMap::new();
        plan_relationship_part_updates(
            archive,
            &part_names,
            &mut delete_parts,
            &mut read_cache,
            &mut updated_parts,
        )?;
        plan_content_types_update(
            archive,
            &mut delete_parts,
            &mut read_cache,
            &mut updated_parts,
            target_kind,
        )?;

        let mut zip = ZipWriter::new(output);
        let options = FileOptions::<()>::default().compression_method(CompressionMethod::Deflated);

        for i in 0..archive.len() {
            let file = archive.by_index(i)?;
            if file.is_dir() {
                continue;
            }

            let name = file.name().to_string();
            let canonical_name = canonical_part_name(&name);
            if delete_parts.contains(&canonical_name) {
                continue;
            }

            if let Some(bytes) = updated_parts.get(&canonical_name) {
                zip.start_file(name, options)?;
                zip.write_all(bytes)?;
            } else {
                zip.raw_copy_file(file)?;
            }
        }

        zip.finish()?;
        Ok(())
    }

    fn list_part_names<R: Read + Seek>(
        archive: &mut ZipArchive<R>,
    ) -> Result<BTreeSet<String>, StreamingPatchError> {
        let mut out = BTreeSet::new();
        for i in 0..archive.len() {
            let file = archive.by_index(i)?;
            if file.is_dir() {
                continue;
            }
            out.insert(canonical_part_name(file.name()));
        }
        Ok(out)
    }

    fn compute_macro_delete_set<R: Read + Seek>(
        archive: &mut ZipArchive<R>,
        part_names: &BTreeSet<String>,
        read_cache: &mut HashMap<String, Vec<u8>>,
    ) -> Result<BTreeSet<String>, StreamingPatchError> {
        let mut delete = BTreeSet::new();

        // VBA project payloads.
        delete.insert("xl/vbaproject.bin".to_string());
        delete.insert("xl/vbadata.xml".to_string());
        delete.insert("xl/vbaprojectsignature.bin".to_string());

        // Ribbon customizations.
        for name in part_names {
            if name.starts_with("customui/") {
                delete.insert(name.clone());
            }
        }

        // ActiveX + legacy form controls.
        for name in part_names {
            if name.starts_with("xl/activex/")
                || name.starts_with("xl/ctrlprops/")
                || name.starts_with("xl/controls/")
            {
                delete.insert(name.clone());
            }
        }

        // Legacy macro surfaces beyond VBA:
        // - Excel 4.0 macro sheets (XLM) stored under `xl/macrosheets/**`
        // - Dialog sheets stored under `xl/dialogsheets/**`
        for name in part_names {
            if name.starts_with("xl/macrosheets/") || name.starts_with("xl/dialogsheets/") {
                delete.insert(name.clone());
            }
        }

        // Parts referenced by `xl/_rels/vbaProject.bin.rels` (e.g. signature payloads).
        if let Some(rels_part) = find_part_name(part_names, "xl/_rels/vbaProject.bin.rels") {
            let rels_bytes = read_zip_part(archive, &rels_part, read_cache)?;
            let targets = parse_internal_relationship_targets(
                &rels_bytes,
                "xl/vbaProject.bin",
                &rels_part,
                part_names,
            )?;
            delete.extend(targets);
        }

        // ActiveX controls embedded into VML drawings can reference OLE/ActiveX binaries.
        delete.extend(find_vml_ole_object_targets(
            archive, part_names, read_cache,
        )?);

        // Delete parts referenced exclusively by deleted macro parts (e.g. `xl/embeddings/*`).
        let graph = RelationshipGraph::build(archive, part_names, read_cache)?;
        delete_orphan_targets(&graph, &mut delete);

        // If a part is deleted, its relationship part must also be deleted.
        let rels_to_remove: Vec<String> = delete
            .iter()
            .filter(|name| !name.ends_with(".rels"))
            .map(|name| crate::path::rels_for_part(name))
            .collect();
        delete.extend(rels_to_remove);

        Ok(delete)
    }

    fn find_vml_ole_object_targets<R: Read + Seek>(
        archive: &mut ZipArchive<R>,
        part_names: &BTreeSet<String>,
        read_cache: &mut HashMap<String, Vec<u8>>,
    ) -> Result<BTreeSet<String>, StreamingPatchError> {
        let mut out = BTreeSet::new();

        for vml_part in part_names {
            if !vml_part.ends_with(".vml") {
                continue;
            }

            // Only VML drawings can contain `<o:OLEObject>` control shapes.
            if !vml_part.starts_with("xl/drawings/") {
                continue;
            }

            let vml_bytes = read_zip_part(archive, vml_part, read_cache)?;
            let rel_ids = parse_vml_ole_object_relationship_ids(&vml_bytes)?;
            if rel_ids.is_empty() {
                continue;
            }

            let rels_part = crate::path::rels_for_part(vml_part);
            let Some(rels_part) = find_part_name(part_names, &rels_part) else {
                continue;
            };
            let rels_bytes = read_zip_part(archive, &rels_part, read_cache)?;

            out.extend(parse_relationship_targets_for_ids(
                &rels_bytes,
                vml_part,
                &rel_ids,
                part_names,
            )?);
        }

        Ok(out)
    }

    fn parse_vml_ole_object_relationship_ids(
        xml: &[u8],
    ) -> Result<BTreeSet<String>, StreamingPatchError> {
        let mut reader = XmlReader::from_reader(xml);
        reader.config_mut().trim_text(false);
        let mut buf = Vec::new();
        let mut namespace_context = NamespaceContext::default();
        let mut ids = BTreeSet::new();

        loop {
            match reader.read_event_into(&mut buf)? {
                Event::Eof => break,
                Event::Start(ref e) => {
                    let changes = namespace_context.apply_namespace_decls(e)?;
                    if crate::openxml::local_name(e.name().as_ref())
                        .eq_ignore_ascii_case(b"OLEObject")
                    {
                        collect_relationship_id_attrs(e, &namespace_context, &mut ids)?;
                    }
                    namespace_context.push(changes);
                }
                Event::Empty(ref e) => {
                    let changes = namespace_context.apply_namespace_decls(e)?;
                    if crate::openxml::local_name(e.name().as_ref())
                        .eq_ignore_ascii_case(b"OLEObject")
                    {
                        collect_relationship_id_attrs(e, &namespace_context, &mut ids)?;
                    }
                    namespace_context.rollback(changes);
                }
                Event::End(_) => namespace_context.pop(),
                _ => {}
            }
            buf.clear();
        }

        Ok(ids)
    }

    fn parse_relationship_targets_for_ids(
        xml: &[u8],
        source_part: &str,
        ids: &BTreeSet<String>,
        part_names: &BTreeSet<String>,
    ) -> Result<BTreeSet<String>, StreamingPatchError> {
        let mut reader = XmlReader::from_reader(xml);
        reader.config_mut().trim_text(false);
        let mut buf = Vec::new();
        let mut out = BTreeSet::new();

        loop {
            match reader.read_event_into(&mut buf)? {
                Event::Eof => break,
                Event::Start(ref e) | Event::Empty(ref e)
                    if crate::openxml::local_name(e.name().as_ref()) == b"Relationship" =>
                {
                    let mut id = None;
                    let mut target = None;
                    let mut target_mode = None;
                    for attr in e.attributes().with_checks(false) {
                        let attr = attr?;
                        match crate::openxml::local_name(attr.key.as_ref()) {
                            b"Id" => id = Some(attr.unescape_value()?.into_owned()),
                            b"Target" => target = Some(attr.unescape_value()?.into_owned()),
                            b"TargetMode" => {
                                target_mode = Some(attr.unescape_value()?.into_owned())
                            }
                            _ => {}
                        }
                    }

                    if target_mode
                        .as_deref()
                        .is_some_and(|mode| mode.trim().eq_ignore_ascii_case("External"))
                    {
                        continue;
                    }

                    let Some(id) = id else {
                        continue;
                    };
                    if !ids.contains(&id) {
                        continue;
                    }

                    let Some(target) = target else {
                        continue;
                    };
                    let target = strip_fragment(&target);
                    let resolved = resolve_target_for_source(source_part, target);
                    let resolved = canonical_part_name(&resolved);

                    // Worksheet OLE objects are stored under `xl/embeddings/` and referenced from
                    // `<oleObjects>` in sheet XML (valid in `.xlsx`). For macro stripping we only
                    // delete embedding binaries referenced by VML `<o:OLEObject>` control shapes.
                    if resolved.starts_with("xl/embeddings/") && part_names.contains(&resolved) {
                        out.insert(resolved);
                    }
                }
                _ => {}
            }
            buf.clear();
        }

        Ok(out)
    }

    fn delete_orphan_targets(graph: &RelationshipGraph, delete: &mut BTreeSet<String>) {
        let mut queue: VecDeque<String> = delete.iter().cloned().collect();
        while let Some(source) = queue.pop_front() {
            let Some(targets) = graph.outgoing.get(&source) else {
                continue;
            };
            for target in targets {
                if delete.contains(target) {
                    continue;
                }
                let Some(inbound) = graph.inbound.get(target) else {
                    continue;
                };
                // Only delete parts that are referenced exclusively by parts we're already deleting.
                if inbound.iter().all(|src| delete.contains(src)) {
                    delete.insert(target.clone());
                    queue.push_back(target.clone());
                }
            }
        }
    }

    fn plan_relationship_part_updates<R: Read + Seek>(
        archive: &mut ZipArchive<R>,
        part_names: &BTreeSet<String>,
        delete_parts: &mut BTreeSet<String>,
        read_cache: &mut HashMap<String, Vec<u8>>,
        updated_parts: &mut HashMap<String, Vec<u8>>,
    ) -> Result<(), StreamingPatchError> {
        let rels_names: Vec<String> = part_names
            .iter()
            .filter(|name| name.ends_with(".rels"))
            .cloned()
            .collect();

        for rels_name in rels_names {
            if delete_parts.contains(&rels_name) {
                continue;
            }

            let Some(source_part) = source_part_from_rels_part(&rels_name) else {
                continue;
            };

            // If the relationship source is gone (or will be deleted), remove the `.rels` part.
            if !source_part.is_empty()
                && (!part_names.contains(&source_part) || delete_parts.contains(&source_part))
            {
                delete_parts.insert(rels_name);
                continue;
            }

            let bytes = read_zip_part(archive, &rels_name, read_cache)?;
            let (updated, removed_ids) = strip_deleted_relationships(
                &rels_name,
                &source_part,
                &bytes,
                delete_parts,
                part_names,
            )?;

            if let Some(updated) = updated {
                updated_parts.insert(rels_name.clone(), updated);
            }

            if !removed_ids.is_empty() {
                strip_source_relationship_references(
                    archive,
                    &source_part,
                    &removed_ids,
                    delete_parts,
                    read_cache,
                    updated_parts,
                )?;
            }
        }

        Ok(())
    }

    fn plan_content_types_update<R: Read + Seek>(
        archive: &mut ZipArchive<R>,
        delete_parts: &mut BTreeSet<String>,
        read_cache: &mut HashMap<String, Vec<u8>>,
        updated_parts: &mut HashMap<String, Vec<u8>>,
        target_kind: WorkbookKind,
    ) -> Result<(), StreamingPatchError> {
        let ct_name = canonical_part_name("[Content_Types].xml");
        if !delete_parts.contains(&ct_name) && zip_part_exists(archive, &ct_name)? {
            let existing = read_zip_part(archive, &ct_name, read_cache)?;
            if let Some(updated) = strip_content_types(&existing, delete_parts, target_kind)? {
                updated_parts.insert(ct_name, updated);
            }
        }

        Ok(())
    }

    fn zip_part_exists<R: Read + Seek>(
        archive: &mut ZipArchive<R>,
        name: &str,
    ) -> Result<bool, StreamingPatchError> {
        super::zip_part_exists(archive, name)
    }

    fn strip_deleted_relationships(
        rels_part_name: &str,
        source_part: &str,
        xml: &[u8],
        delete_parts: &BTreeSet<String>,
        part_names: &BTreeSet<String>,
    ) -> Result<(Option<Vec<u8>>, BTreeSet<String>), StreamingPatchError> {
        let mut reader = XmlReader::from_reader(xml);
        reader.config_mut().trim_text(false);
        let mut writer = XmlWriter::new(Vec::with_capacity(xml.len()));

        let mut buf = Vec::new();
        let mut changed = false;
        let mut removed_ids = BTreeSet::new();
        let mut skip_depth = 0usize;

        loop {
            let ev = reader.read_event_into(&mut buf)?;

            if skip_depth > 0 {
                match ev {
                    Event::Start(_) => skip_depth += 1,
                    Event::End(_) => {
                        skip_depth -= 1;
                    }
                    Event::Eof => break,
                    _ => {}
                }
                buf.clear();
                continue;
            }

            match ev {
                Event::Eof => break,
                Event::Empty(e)
                    if crate::openxml::local_name(e.name().as_ref()) == b"Relationship" =>
                {
                    if should_remove_relationship(
                        rels_part_name,
                        source_part,
                        &e,
                        delete_parts,
                        part_names,
                    )? {
                        changed = true;
                        if let Some(id) = relationship_id(&e)? {
                            removed_ids.insert(id);
                        }
                        buf.clear();
                        continue;
                    }
                    writer.write_event(Event::Empty(e.to_owned()))?;
                }
                Event::Start(e)
                    if crate::openxml::local_name(e.name().as_ref()) == b"Relationship" =>
                {
                    if should_remove_relationship(
                        rels_part_name,
                        source_part,
                        &e,
                        delete_parts,
                        part_names,
                    )? {
                        changed = true;
                        if let Some(id) = relationship_id(&e)? {
                            removed_ids.insert(id);
                        }
                        skip_depth = 1;
                        buf.clear();
                        continue;
                    }
                    writer.write_event(Event::Start(e.to_owned()))?;
                }
                other => writer.write_event(other.into_owned())?,
            }

            buf.clear();
        }

        let updated = if changed {
            Some(writer.into_inner())
        } else {
            None
        };

        Ok((updated, removed_ids))
    }

    fn should_remove_relationship(
        rels_part_name: &str,
        source_part: &str,
        e: &BytesStart<'_>,
        delete_parts: &BTreeSet<String>,
        part_names: &BTreeSet<String>,
    ) -> Result<bool, StreamingPatchError> {
        let mut target = None;
        let mut target_mode = None;
        let mut rel_type = None;

        for attr in e.attributes().with_checks(false) {
            let attr = attr?;
            match crate::openxml::local_name(attr.key.as_ref()) {
                b"Target" => target = Some(attr.unescape_value()?.into_owned()),
                b"TargetMode" => target_mode = Some(attr.unescape_value()?.into_owned()),
                b"Type" => rel_type = Some(attr.unescape_value()?.into_owned()),
                _ => {}
            }
        }

        if target_mode
            .as_deref()
            .is_some_and(|mode| mode.eq_ignore_ascii_case("External"))
        {
            return Ok(false);
        }

        if rels_part_name == "_rels/.rels"
            && rel_type
                .as_deref()
                .is_some_and(|ty| CUSTOM_UI_REL_TYPES.iter().any(|known| ty == *known))
        {
            return Ok(true);
        }

        let Some(target) = target else {
            return Ok(false);
        };

        let target = strip_fragment(&target);
        let resolved = resolve_target_best_effort(source_part, rels_part_name, target, part_names);
        Ok(delete_parts.contains(&resolved))
    }

    fn relationship_id(e: &BytesStart<'_>) -> Result<Option<String>, StreamingPatchError> {
        for attr in e.attributes().with_checks(false) {
            let attr = attr?;
            if crate::openxml::local_name(attr.key.as_ref()).eq_ignore_ascii_case(b"Id") {
                return Ok(Some(attr.unescape_value()?.into_owned()));
            }
        }
        Ok(None)
    }

    fn strip_source_relationship_references<R: Read + Seek>(
        archive: &mut ZipArchive<R>,
        source_part: &str,
        removed_ids: &BTreeSet<String>,
        delete_parts: &BTreeSet<String>,
        read_cache: &mut HashMap<String, Vec<u8>>,
        updated_parts: &mut HashMap<String, Vec<u8>>,
    ) -> Result<(), StreamingPatchError> {
        if source_part.is_empty()
            || !(source_part.ends_with(".xml") || source_part.ends_with(".vml"))
            || removed_ids.is_empty()
            || delete_parts.contains(source_part)
        {
            return Ok(());
        }

        let xml = read_zip_part(archive, source_part, read_cache)?;
        if let Some(updated) = strip_relationship_id_references(&xml, removed_ids)? {
            updated_parts.insert(source_part.to_string(), updated);
        }

        Ok(())
    }

    fn strip_relationship_id_references(
        xml: &[u8],
        removed_ids: &BTreeSet<String>,
    ) -> Result<Option<Vec<u8>>, StreamingPatchError> {
        let mut reader = XmlReader::from_reader(xml);
        reader.config_mut().trim_text(false);
        let mut writer = XmlWriter::new(Vec::with_capacity(xml.len()));

        let mut buf = Vec::new();
        let mut changed = false;
        let mut skip_depth = 0usize;
        let mut namespace_context = NamespaceContext::default();

        loop {
            let ev = reader.read_event_into(&mut buf)?;

            if skip_depth > 0 {
                match ev {
                    Event::Start(_) => skip_depth += 1,
                    Event::End(_) => skip_depth -= 1,
                    Event::Eof => break,
                    _ => {}
                }
                buf.clear();
                continue;
            }

            match ev {
                Event::Eof => break,
                Event::Start(e) => {
                    let changes = namespace_context.apply_namespace_decls(&e)?;
                    if element_has_removed_relationship_id(&e, &namespace_context, removed_ids)? {
                        changed = true;
                        namespace_context.rollback(changes);
                        skip_depth = 1;
                        buf.clear();
                        continue;
                    }
                    namespace_context.push(changes);
                    writer.write_event(Event::Start(e.to_owned()))?;
                }
                Event::Empty(e) => {
                    let changes = namespace_context.apply_namespace_decls(&e)?;
                    if element_has_removed_relationship_id(&e, &namespace_context, removed_ids)? {
                        changed = true;
                        namespace_context.rollback(changes);
                        buf.clear();
                        continue;
                    }
                    namespace_context.rollback(changes);
                    writer.write_event(Event::Empty(e.to_owned()))?;
                }
                Event::End(e) => {
                    namespace_context.pop();
                    writer.write_event(Event::End(e.to_owned()))?;
                }
                other => writer.write_event(other.into_owned())?,
            }

            buf.clear();
        }

        if changed {
            Ok(Some(writer.into_inner()))
        } else {
            Ok(None)
        }
    }

    fn element_has_removed_relationship_id(
        e: &BytesStart<'_>,
        namespace_context: &NamespaceContext,
        removed_ids: &BTreeSet<String>,
    ) -> Result<bool, StreamingPatchError> {
        for attr in e.attributes().with_checks(false) {
            let attr = attr?;
            let key = attr.key.as_ref();

            if key == b"xmlns" || key.starts_with(b"xmlns:") {
                continue;
            }

            let (prefix, local) = split_prefixed_name(key);
            let namespace_uri = prefix.and_then(|p| namespace_context.namespace_for_prefix(p));

            if !is_relationship_id_attribute(namespace_uri, local) {
                continue;
            }
            let value = attr.unescape_value()?;
            if removed_ids.contains(value.as_ref()) {
                return Ok(true);
            }
        }
        Ok(false)
    }

    fn split_prefixed_name(name: &[u8]) -> (Option<&[u8]>, &[u8]) {
        match name.iter().position(|b| *b == b':') {
            Some(idx) => (Some(&name[..idx]), &name[idx + 1..]),
            None => (None, name),
        }
    }

    fn is_relationship_id_attribute(namespace_uri: Option<&[u8]>, local_name: &[u8]) -> bool {
        // Be defensive: VML/Office markup commonly uses `o:relid`, but some documents use other
        // prefixes or even no prefix at all. If the local-name is `relid` we treat it as a
        // relationship pointer regardless of namespace.
        if local_name.eq_ignore_ascii_case(b"relid") {
            return true;
        }

        match namespace_uri {
            Some(ns) if ns == RELATIONSHIPS_NS => {
                local_name.eq_ignore_ascii_case(b"id")
                    || local_name.eq_ignore_ascii_case(b"embed")
                    || local_name.eq_ignore_ascii_case(b"link")
            }
            _ => false,
        }
    }

    fn strip_content_types(
        xml: &[u8],
        delete_parts: &BTreeSet<String>,
        target_kind: WorkbookKind,
    ) -> Result<Option<Vec<u8>>, StreamingPatchError> {
        let mut reader = XmlReader::from_reader(xml);
        reader.config_mut().trim_text(false);
        let mut writer = XmlWriter::new(Vec::with_capacity(xml.len()));

        let mut buf = Vec::new();
        let mut changed = false;
        let mut skip_depth = 0usize;

        loop {
            let ev = reader.read_event_into(&mut buf)?;

            if skip_depth > 0 {
                match ev {
                    Event::Start(_) => skip_depth += 1,
                    Event::End(_) => skip_depth -= 1,
                    Event::Eof => break,
                    _ => {}
                }
                buf.clear();
                continue;
            }

            match ev {
                Event::Eof => break,
                Event::Empty(e) if crate::openxml::local_name(e.name().as_ref()) == b"Override" => {
                    if let Some(updated) = patched_override(&e, delete_parts, target_kind)? {
                        if updated.is_none() {
                            changed = true;
                            buf.clear();
                            continue;
                        }
                        if let Some(updated) = updated {
                            changed = true;
                            writer.write_event(Event::Empty(updated))?;
                            buf.clear();
                            continue;
                        }
                    }
                    writer.write_event(Event::Empty(e.to_owned()))?;
                }
                Event::Start(e) if crate::openxml::local_name(e.name().as_ref()) == b"Override" => {
                    // `<Override>` parts are expected to be empty, but handle the non-empty form.
                    if let Some(updated) = patched_override(&e, delete_parts, target_kind)? {
                        if updated.is_none() {
                            changed = true;
                            skip_depth = 1;
                            buf.clear();
                            continue;
                        }
                        if let Some(updated) = updated {
                            changed = true;
                            writer.write_event(Event::Start(updated))?;
                            buf.clear();
                            continue;
                        }
                    }
                    writer.write_event(Event::Start(e.to_owned()))?;
                }
                other => writer.write_event(other.into_owned())?,
            }

            buf.clear();
        }

        if changed {
            Ok(Some(writer.into_inner()))
        } else {
            Ok(None)
        }
    }

    // Returns:
    // - Ok(None) -> keep original
    // - Ok(Some(None)) -> remove element
    // - Ok(Some(Some(updated))) -> replace element
    fn patched_override(
        e: &BytesStart<'_>,
        delete_parts: &BTreeSet<String>,
        target_kind: WorkbookKind,
    ) -> Result<Option<Option<BytesStart<'static>>>, StreamingPatchError> {
        let mut part_name = None;
        let mut content_type = None;

        for attr in e.attributes().with_checks(false) {
            let attr = attr?;
            match crate::openxml::local_name(attr.key.as_ref()) {
                key if key.eq_ignore_ascii_case(b"PartName") => {
                    part_name = Some(attr.unescape_value()?.into_owned())
                }
                key if key.eq_ignore_ascii_case(b"ContentType") => {
                    content_type = Some(attr.unescape_value()?.into_owned())
                }
                _ => {}
            }
        }

        let Some(part_name) = part_name else {
            return Ok(None);
        };

        let normalized = part_name.strip_prefix('/').unwrap_or(part_name.as_str());
        let normalized = canonical_part_name(normalized);
        if delete_parts.contains(&normalized) {
            return Ok(Some(None));
        }

        if content_type
            .as_deref()
            .is_some_and(|ty| ty.contains("macroEnabled.main+xml"))
        {
            let workbook_main_type = target_kind.macro_free_kind().workbook_content_type();

            // Preserve the original element's qualified name (including any namespace prefix).
            let tag_name = e.name();
            let tag_name = std::str::from_utf8(tag_name.as_ref()).unwrap_or("Override");
            let mut updated = BytesStart::new(tag_name);

            // Preserve all attributes verbatim (including any prefixes/ordering), except for
            // `ContentType`, which is rewritten to the non-macro workbook content type.
            let mut saw_content_type = false;
            for attr in e.attributes().with_checks(false) {
                let attr = attr?;
                if crate::openxml::local_name(attr.key.as_ref())
                    .eq_ignore_ascii_case(b"ContentType")
                {
                    saw_content_type = true;
                    updated.push_attribute((attr.key.as_ref(), workbook_main_type.as_bytes()));
                } else {
                    updated.push_attribute((attr.key.as_ref(), attr.value.as_ref()));
                }
            }
            if !saw_content_type {
                updated.push_attribute(("ContentType", workbook_main_type));
            }

            return Ok(Some(Some(updated.into_owned())));
        }

        Ok(None)
    }

    fn parse_internal_relationship_targets(
        xml: &[u8],
        source_part: &str,
        rels_part: &str,
        part_names: &BTreeSet<String>,
    ) -> Result<Vec<String>, StreamingPatchError> {
        let mut reader = XmlReader::from_reader(xml);
        reader.config_mut().trim_text(false);
        let mut buf = Vec::new();
        let mut out = Vec::new();

        loop {
            match reader.read_event_into(&mut buf)? {
                Event::Eof => break,
                Event::Start(ref e) | Event::Empty(ref e)
                    if crate::openxml::local_name(e.name().as_ref()) == b"Relationship" =>
                {
                    let mut target = None;
                    let mut target_mode = None;
                    for attr in e.attributes().with_checks(false) {
                        let attr = attr?;
                        match crate::openxml::local_name(attr.key.as_ref()) {
                            b"Target" => target = Some(attr.unescape_value()?.into_owned()),
                            b"TargetMode" => {
                                target_mode = Some(attr.unescape_value()?.into_owned())
                            }
                            _ => {}
                        }
                    }

                    if target_mode
                        .as_deref()
                        .is_some_and(|mode| mode.eq_ignore_ascii_case("External"))
                    {
                        continue;
                    }

                    let Some(target) = target else {
                        continue;
                    };
                    let target = strip_fragment(&target);
                    out.push(resolve_target_best_effort(
                        source_part,
                        rels_part,
                        target,
                        part_names,
                    ));
                }
                _ => {}
            }
            buf.clear();
        }

        Ok(out)
    }

    fn strip_fragment(target: &str) -> &str {
        target
            .split_once('#')
            .map(|(base, _)| base)
            .unwrap_or(target)
    }

    fn resolve_target_for_source(source_part: &str, target: &str) -> String {
        if source_part.is_empty() {
            crate::path::resolve_target("", target)
        } else {
            crate::path::resolve_target(source_part, target)
        }
    }

    fn resolve_target_best_effort(
        source_part: &str,
        rels_part: &str,
        target: &str,
        part_names: &BTreeSet<String>,
    ) -> String {
        // Match the in-memory macro stripper: prefer the standard source-relative resolution, but
        // fall back to interpreting the target as relative to the `.rels` directory when the
        // canonical path doesn't exist (common in some producers for workbook-level parts).
        let direct = resolve_target_for_source(source_part, target);
        if let Some(found) = find_part_name(part_names, &direct) {
            return found;
        }

        let rels_relative = crate::path::resolve_target(rels_part, target);
        if let Some(found) = find_part_name(part_names, &rels_relative) {
            return found;
        }

        let direct_canonical = canonical_part_name(&direct);
        if !direct_canonical.starts_with("xl/") {
            let xl_prefixed = format!("xl/{direct_canonical}");
            if let Some(found) = find_part_name(part_names, &xl_prefixed) {
                return found;
            }
        }

        direct_canonical
    }

    fn source_part_from_rels_part(rels_part: &str) -> Option<String> {
        if rels_part == "_rels/.rels" {
            return Some(String::new());
        }

        if let Some(rels_file) = rels_part.strip_prefix("_rels/") {
            let rels_file = rels_file.strip_suffix(".rels")?;
            return Some(rels_file.to_string());
        }

        let (dir, rels_file) = rels_part.rsplit_once("/_rels/")?;
        let rels_file = rels_file.strip_suffix(".rels")?;

        if dir.is_empty() {
            return Some(rels_file.to_string());
        }

        Some(format!("{dir}/{rels_file}"))
    }

    struct RelationshipGraph {
        outgoing: BTreeMap<String, BTreeSet<String>>,
        inbound: BTreeMap<String, BTreeSet<String>>,
    }

    impl RelationshipGraph {
        fn build<R: Read + Seek>(
            archive: &mut ZipArchive<R>,
            part_names: &BTreeSet<String>,
            read_cache: &mut HashMap<String, Vec<u8>>,
        ) -> Result<Self, StreamingPatchError> {
            let mut outgoing: BTreeMap<String, BTreeSet<String>> = BTreeMap::new();
            let mut inbound: BTreeMap<String, BTreeSet<String>> = BTreeMap::new();

            for rels_part in part_names.iter().filter(|name| name.ends_with(".rels")) {
                let Some(source_part) = source_part_from_rels_part(rels_part) else {
                    continue;
                };
                let source_part = if source_part.is_empty() {
                    source_part
                } else {
                    match find_part_name(part_names, &source_part) {
                        Some(found) => found,
                        None => continue,
                    }
                };

                // Ignore orphan `.rels` parts; they'll be removed during cleanup.
                // (The existence check above uses `find_part_name` and already filtered missing ones.)

                let bytes = read_zip_part(archive, rels_part, read_cache)?;
                let targets = parse_internal_relationship_targets(
                    &bytes,
                    &source_part,
                    rels_part,
                    part_names,
                )?;
                for target in targets {
                    outgoing
                        .entry(source_part.clone())
                        .or_default()
                        .insert(target.clone());
                    inbound
                        .entry(target)
                        .or_default()
                        .insert(source_part.clone());
                }
            }

            Ok(Self { outgoing, inbound })
        }
    }

    #[derive(Debug, Default)]
    struct NamespaceContext {
        /// prefix -> namespace URI
        prefixes: BTreeMap<Vec<u8>, Vec<u8>>,
        /// Stack of prefix changes for each started element that was written.
        stack: Vec<Vec<(Vec<u8>, Option<Vec<u8>>)>>,
    }

    impl NamespaceContext {
        fn apply_namespace_decls(
            &mut self,
            e: &BytesStart<'_>,
        ) -> Result<Vec<(Vec<u8>, Option<Vec<u8>>)>, StreamingPatchError> {
            let mut changes = Vec::new();

            for attr in e.attributes().with_checks(false) {
                let attr = attr?;
                let key = attr.key.as_ref();

                // Default namespace (`xmlns="..."`) affects element names, but not attributes.
                if key == b"xmlns" {
                    continue;
                }

                let Some(prefix) = key.strip_prefix(b"xmlns:") else {
                    continue;
                };

                let uri = attr.unescape_value()?.into_owned().into_bytes();
                let old = self.prefixes.insert(prefix.to_vec(), uri);
                changes.push((prefix.to_vec(), old));
            }

            Ok(changes)
        }

        fn rollback(&mut self, changes: Vec<(Vec<u8>, Option<Vec<u8>>)>) {
            for (prefix, old) in changes.into_iter().rev() {
                match old {
                    Some(uri) => {
                        self.prefixes.insert(prefix, uri);
                    }
                    None => {
                        self.prefixes.remove(&prefix);
                    }
                }
            }
        }

        fn push(&mut self, changes: Vec<(Vec<u8>, Option<Vec<u8>>)>) {
            self.stack.push(changes);
        }

        fn pop(&mut self) {
            if let Some(changes) = self.stack.pop() {
                self.rollback(changes);
            }
        }

        fn namespace_for_prefix(&self, prefix: &[u8]) -> Option<&[u8]> {
            self.prefixes.get(prefix).map(Vec::as_slice)
        }
    }

    fn collect_relationship_id_attrs(
        e: &BytesStart<'_>,
        namespace_context: &NamespaceContext,
        out: &mut BTreeSet<String>,
    ) -> Result<(), StreamingPatchError> {
        for attr in e.attributes().with_checks(false) {
            let attr = attr?;
            let key = attr.key.as_ref();

            if key == b"xmlns" || key.starts_with(b"xmlns:") {
                continue;
            }

            let (prefix, local) = split_prefixed_name(key);
            let namespace_uri = prefix.and_then(|p| namespace_context.namespace_for_prefix(p));
            if !is_relationship_id_attribute(namespace_uri, local) {
                continue;
            }
            out.insert(attr.unescape_value()?.into_owned());
        }

        Ok(())
    }
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
    patch_xlsx_streaming_workbook_cell_patches_with_styles_and_recalc_policy(
        input,
        output,
        patches,
        style_table,
        RecalcPolicy::default(),
    )
}

/// Apply [`WorkbookCellPatches`] using the streaming ZIP rewriter, resolving `style_id` overrides
/// via `styles.xml`, with a configurable [`RecalcPolicy`].
///
/// `policy_on_formula_change` is applied **only** when the patch set changes formulas (including
/// removing formulas). When no formulas change, [`RecalcPolicy::PRESERVE`] is used regardless of
/// the provided policy.
pub fn patch_xlsx_streaming_workbook_cell_patches_with_styles_and_recalc_policy<
    R: Read + Seek,
    W: Write + Seek,
>(
    input: R,
    output: W,
    patches: &WorkbookCellPatches,
    style_table: &StyleTable,
    policy_on_formula_change: RecalcPolicy,
) -> Result<(), StreamingPatchError> {
    let part_overrides = HashMap::new();
    patch_xlsx_streaming_workbook_cell_patches_with_styles_and_part_overrides_and_recalc_policy(
        input,
        output,
        patches,
        style_table,
        &part_overrides,
        policy_on_formula_change,
    )
}

/// Apply [`WorkbookCellPatches`] using the streaming ZIP rewriter, resolving `style_id` overrides
/// via `styles.xml`, plus arbitrary part overrides.
///
/// This variant updates `styles.xml` deterministically when new styles are introduced and applies
/// [`PartOverride::{Replace,Add,Remove}`] to non-worksheet parts without inflating the entire
/// workbook package into a [`crate::XlsxPackage`].
pub fn patch_xlsx_streaming_workbook_cell_patches_with_styles_and_part_overrides<
    R: Read + Seek,
    W: Write + Seek,
>(
    input: R,
    output: W,
    patches: &WorkbookCellPatches,
    style_table: &StyleTable,
    part_overrides: &HashMap<String, PartOverride>,
) -> Result<(), StreamingPatchError> {
    patch_xlsx_streaming_workbook_cell_patches_with_styles_and_part_overrides_and_recalc_policy(
        input,
        output,
        patches,
        style_table,
        part_overrides,
        RecalcPolicy::default(),
    )
}

/// Apply [`WorkbookCellPatches`] using the streaming ZIP rewriter, resolving `style_id` overrides
/// via `styles.xml`, plus arbitrary part overrides, with a configurable [`RecalcPolicy`].
///
/// `policy_on_formula_change` is applied **only** when the patch set changes formulas (including
/// removing formulas). When no formulas change, [`RecalcPolicy::PRESERVE`] is used regardless of
/// the provided policy.
pub fn patch_xlsx_streaming_workbook_cell_patches_with_styles_and_part_overrides_and_recalc_policy<
    R: Read + Seek,
    W: Write + Seek,
>(
    input: R,
    output: W,
    patches: &WorkbookCellPatches,
    style_table: &StyleTable,
    part_overrides: &HashMap<String, PartOverride>,
    policy_on_formula_change: RecalcPolicy,
) -> Result<(), StreamingPatchError> {
    if patches.is_empty() {
        let mut archive = ZipArchive::new(input)?;
        patch_xlsx_streaming_with_archive(
            &mut archive,
            output,
            &HashMap::new(),
            &HashMap::new(),
            &HashMap::new(),
            &HashMap::new(),
            part_overrides,
            RecalcPolicy::PRESERVE,
        )?;
        return Ok(());
    }

    let mut archive = ZipArchive::new(input)?;
    let part_names = list_zip_part_names(&mut archive)?;

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
        if rel
            .target_mode
            .as_deref()
            .is_some_and(|mode| mode.trim().eq_ignore_ascii_case("External"))
        {
            continue;
        }
        let resolved = resolve_target("xl/workbook.xml", &rel.target);
        let resolved = find_zip_part_name(&part_names, &resolved).unwrap_or(resolved);
        if rel.type_uri == REL_TYPE_STYLES {
            styles_part.get_or_insert(resolved.clone());
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
    let mut col_properties_by_part: HashMap<String, BTreeMap<u32, ColProperties>> = HashMap::new();
    let mut saw_formula_patch = false;
    for (sheet_selector, sheet_patches) in patches.sheets() {
        if sheet_patches.is_empty() {
            continue;
        }
        let worksheet_part =
            resolve_worksheet_part_for_selector(sheet_selector, &workbook_sheets, &rel_targets)?;
        let worksheet_part = find_zip_part_name(&part_names, &worksheet_part)
            .unwrap_or_else(|| worksheet_part.clone());

        if let Some(cols) = sheet_patches.col_properties() {
            col_properties_by_part.insert(worksheet_part.clone(), cols.clone());
        }

        for (cell_ref, patch) in sheet_patches.iter() {
            let (value, formula) = match patch {
                CellPatch::Clear { .. } => (CellValue::Empty, None),
                CellPatch::Set { value, formula, .. } => (value.clone(), formula.clone()),
            };
            saw_formula_patch |= formula_is_material(formula.as_deref());

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
                        .with_xf_index(xf_index)
                        .with_vm(patch.vm_override())
                        .with_cm(patch.cm_override()),
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
        policy_on_formula_change
    } else {
        RecalcPolicy::PRESERVE
    };

    patch_xlsx_streaming_with_archive(
        &mut archive,
        output,
        &patches_by_part,
        &col_properties_by_part,
        &pre_read_parts,
        &updated_parts,
        part_overrides,
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
        .find(|s| formula_model::sheet_name_eq_case_insensitive(&s.name, selector))
    {
        return rel_targets.get(&sheet.rel_id).cloned().ok_or_else(|| {
            crate::XlsxError::Invalid(format!("missing worksheet relationship for {}", sheet.name))
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

#[cfg(test)]
mod selector_tests {
    use super::*;
    use formula_model::SheetVisibility;

    #[test]
    fn resolve_worksheet_part_for_selector_matches_unicode_sheet_names_case_insensitive_like_excel() {
        let sheets = vec![crate::WorkbookSheetInfo {
            name: "Straße".to_string(),
            sheet_id: 1,
            rel_id: "rId1".to_string(),
            visibility: SheetVisibility::Visible,
        }];
        let rel_targets: HashMap<String, String> =
            [("rId1".to_string(), "xl/worksheets/sheet1.xml".to_string())].into();

        let resolved = resolve_worksheet_part_for_selector("STRASSE", &sheets, &rel_targets)
            .expect("should match sheet name case-insensitively");
        assert_eq!(resolved, "xl/worksheets/sheet1.xml");
    }
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
        .any(|p| {
            matches!(
                p.value,
                CellValue::String(_)
                    | CellValue::RichText(_)
                    | CellValue::Entity(_)
                    | CellValue::Record(_)
            )
        });
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

    let shared_strings_bytes = {
        let mut file = open_zip_part(archive, &shared_strings_part)?;
        read_zip_file_bytes_with_limit(&mut file, &shared_strings_part, DEFAULT_MAX_ZIP_PART_BYTES)?
    };
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
                    CellValue::Entity(entity) => shared_strings
                        .editor
                        .rich_at(idx)
                        .map(|rt| rt.text.as_str() == entity.display_value.as_str())
                        .unwrap_or(false),
                    CellValue::Record(record) => {
                        let display = record.to_string();
                        shared_strings
                            .editor
                            .rich_at(idx)
                            .map(|rt| rt.text.as_str() == display.as_str())
                            .unwrap_or(false)
                    }
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
            CellValue::Entity(entity) => {
                shared_strings.get_or_insert_plain(entity.display_value.as_str())
            }
            CellValue::Record(record) => {
                let display = record.to_string();
                shared_strings.get_or_insert_plain(display.as_str())
            }
            CellValue::RichText(rich) => shared_strings.get_or_insert_rich(rich),
            _ => 0,
        });
        indices_by_part
            .entry(part.to_string())
            .or_default()
            .insert((patch.cell.row, patch.cell.col), idx);
    }

    let updated_shared_strings = shared_strings.write_if_dirty()?;
    Ok((
        Some(shared_strings_part),
        indices_by_part,
        updated_shared_strings,
    ))
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
        if let Some(rel) = rels.iter().find(|rel| {
            rel.type_uri == REL_TYPE_SHARED_STRINGS
                && !rel
                    .target_mode
                    .as_deref()
                    .is_some_and(|mode| mode.trim().eq_ignore_ascii_case("External"))
        }) {
            let resolved = resolve_target("xl/workbook.xml", &rel.target);
            match open_zip_part(archive, &resolved) {
                Ok(file) => {
                    let name = file.name();
                    return Ok(Some(canonicalize_zip_entry_name(name).into_owned()));
                }
                Err(zip::result::ZipError::FileNotFound) => {
                    return Ok(Some(resolved));
                }
                Err(err) => return Err(err.into()),
            }
        }
    }

    // Fallback: common path when workbook.xml.rels is missing the sharedStrings relationship.
    match open_zip_part(archive, "xl/sharedStrings.xml") {
        Ok(file) => {
            let name = file.name();
            return Ok(Some(canonicalize_zip_entry_name(name).into_owned()));
        }
        Err(zip::result::ZipError::FileNotFound) => {}
        Err(err) => return Err(err.into()),
    };

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

        let mut file = match open_zip_part(archive, part) {
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
            if matches!(
                patch.value,
                CellValue::String(_)
                    | CellValue::RichText(_)
                    | CellValue::Entity(_)
                    | CellValue::Record(_)
            ) {
                targets.insert(patch.cell.to_a1(), (patch.cell.row, patch.cell.col));
            }
        }
        if targets.is_empty() {
            continue;
        }

        let mut file = match open_zip_part(archive, part) {
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
            Event::Start(ref e)
                if current_target.is_some() && local_name(e.name().as_ref()) == b"v" =>
            {
                in_v = true;
            }
            Event::End(ref e)
                if current_target.is_some() && local_name(e.name().as_ref()) == b"v" =>
            {
                in_v = false;
            }
            Event::Text(e) if in_v && current_target.is_some() => {
                if current_t.as_deref() == Some("s") {
                    current_idx = e.unescape()?.trim().parse::<u32>().ok();
                }
            }
            Event::End(ref e)
                if current_target.is_some() && local_name(e.name().as_ref()) == b"c" =>
            {
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
    sheet_uses_row_spans: bool,
    existing_used_range: Option<PatchBounds>,
}

fn scan_worksheet_xml_metadata<R: Read>(
    input: R,
    target_cells: Option<&HashSet<(u32, u32)>>,
) -> Result<(WorksheetXmlMetadata, HashSet<(u32, u32)>), StreamingPatchError> {
    let mut reader = Reader::from_reader(BufReader::new(input));
    reader.config_mut().trim_text(true);
    let mut buf = Vec::new();

    let mut in_sheet_data = false;
    let mut has_dimension = false;
    let mut has_sheet_pr = false;
    let mut sheet_uses_row_spans = false;
    let mut used_range: Option<PatchBounds> = None;
    let mut found_target_cells: HashSet<(u32, u32)> =
        HashSet::with_capacity(target_cells.map_or(0, HashSet::len));

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
                if in_sheet_data && local_name(e.name().as_ref()) == b"row" =>
            {
                if !sheet_uses_row_spans {
                    for attr in e.attributes() {
                        let attr = attr?;
                        if local_name(attr.key.as_ref()) == b"spans" {
                            sheet_uses_row_spans = true;
                            break;
                        }
                    }
                }
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
                    if target_cells
                        .is_some_and(|targets| targets.contains(&(cell_ref.row, cell_ref.col)))
                    {
                        found_target_cells.insert((cell_ref.row, cell_ref.col));
                    }
                }
            }
            Event::Eof => break,
            _ => {}
        }
        buf.clear();
    }

    Ok((
        WorksheetXmlMetadata {
            has_dimension,
            has_sheet_pr,
            sheet_uses_row_spans,
            existing_used_range: used_range,
        },
        found_target_cells,
    ))
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
        CellValue::String(_) | CellValue::Entity(_) | CellValue::Record(_) => {
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
            if formula_is_material(patch.formula.as_deref()) {
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
    match open_zip_part(archive, name) {
        Ok(_) => Ok(true),
        Err(zip::result::ZipError::FileNotFound) => Ok(false),
        Err(err) => Err(err.into()),
    }
}

fn read_zip_part_optional<R: Read + Seek>(
    archive: &mut ZipArchive<R>,
    name: &str,
) -> Result<Option<Vec<u8>>, StreamingPatchError> {
    let mut file = match open_zip_part(archive, name) {
        Ok(file) => file,
        Err(zip::result::ZipError::FileNotFound) => return Ok(None),
        Err(err) => return Err(err.into()),
    };
    Ok(Some(read_zip_file_bytes_with_limit(
        &mut file,
        name,
        DEFAULT_MAX_ZIP_PART_BYTES,
    )?))
}

fn patch_xlsx_streaming_with_archive<R: Read + Seek, W: Write + Seek>(
    archive: &mut ZipArchive<R>,
    output: W,
    patches_by_part: &HashMap<String, Vec<WorksheetCellPatch>>,
    col_properties_by_part: &HashMap<String, BTreeMap<u32, ColProperties>>,
    pre_read_parts: &HashMap<String, Vec<u8>>,
    updated_parts: &HashMap<String, Vec<u8>>,
    part_overrides: &HashMap<String, PartOverride>,
    recalc_policy: RecalcPolicy,
) -> Result<(), StreamingPatchError> {
    let repair_overrides = plan_package_repair_overrides(
        archive,
        pre_read_parts,
        updated_parts,
        part_overrides,
        recalc_policy,
    )?;

    let (shared_strings_part, shared_string_indices, shared_strings_updated) =
        plan_shared_strings(archive, patches_by_part, pre_read_parts)?;

    // Most streaming patching use-cases want to preserve `vm`/`cm` attributes on existing cells to
    // avoid accidentally dropping RichData references (e.g. images-in-cell).
    //
    // Some callers patch incomplete workbook packages missing `[Content_Types].xml` but still
    // containing `xl/workbook.xml`. In that mode, we drop `vm` when the cached value changes to
    // avoid leaving a dangling value-metadata pointer.
    //
    // NOTE: `patch_xlsx_streaming_workbook_cell_patches*` pre-reads `xl/workbook.xml`; keep its
    // default behavior (preserve `vm`) even when `[Content_Types].xml` is absent.
    let workbook_pre_read = pre_read_parts.contains_key("xl/workbook.xml");
    let workbook_present = workbook_pre_read || zip_part_exists(archive, "xl/workbook.xml")?;
    let content_types_present = zip_part_exists(archive, "[Content_Types].xml")?;
    let drop_vm_on_value_change = !workbook_pre_read && workbook_present && !content_types_present;

    let mut non_material_targets_by_part: HashMap<String, HashSet<(u32, u32)>> = HashMap::new();
    for (part, patches) in patches_by_part {
        let mut targets = HashSet::new();
        for patch in patches {
            if patch_is_material_for_insertion(patch) {
                continue;
            }
            targets.insert((patch.cell.row, patch.cell.col));
        }
        if !targets.is_empty() {
            non_material_targets_by_part.insert(part.clone(), targets);
        }
    }

    let mut worksheet_metadata_by_part: HashMap<String, WorksheetXmlMetadata> = HashMap::new();
    let mut existing_non_material_cells_by_part: HashMap<String, HashSet<(u32, u32)>> =
        HashMap::new();
    for part in patches_by_part.keys() {
        let mut file = match open_zip_part(archive, part) {
            Ok(file) => file,
            Err(zip::result::ZipError::FileNotFound) => {
                return Err(StreamingPatchError::MissingWorksheetPart(part.clone()));
            }
            Err(err) => return Err(err.into()),
        };
        let target_cells = non_material_targets_by_part.get(part);
        let (metadata, found_target_cells) = scan_worksheet_xml_metadata(&mut file, target_cells)?;
        worksheet_metadata_by_part.insert(part.clone(), metadata);
        if !found_target_cells.is_empty() {
            existing_non_material_cells_by_part.insert(part.clone(), found_target_cells);
        }
    }

    // Drop patches that are guaranteed to be a no-op:
    // a non-material patch targeting a missing cell cannot reference an existing `<c>` element,
    // and since it will not insert a new cell, it cannot change the worksheet XML.
    let mut effective_patches_by_part: HashMap<String, Vec<WorksheetCellPatch>> = HashMap::new();
    for (part, patches) in patches_by_part {
        let existing_cells = existing_non_material_cells_by_part.get(part);
        let mut filtered = Vec::new();
        for patch in patches {
            if patch_is_material_for_insertion(patch) {
                filtered.push(patch.clone());
                continue;
            }
            if existing_cells.is_some_and(|cells| cells.contains(&(patch.cell.row, patch.cell.col)))
            {
                filtered.push(patch.clone());
            }
        }
        if !filtered.is_empty() {
            effective_patches_by_part.insert(part.clone(), filtered);
        }
    }

    let mut missing_parts: BTreeMap<String, ()> = effective_patches_by_part
        .keys()
        .chain(col_properties_by_part.keys())
        .map(|k| (k.clone(), ()))
        .collect();

    let mut zip = ZipWriter::new(output);
    let options = FileOptions::<()>::default().compression_method(CompressionMethod::Deflated);
    let mut applied_part_overrides: HashSet<String> = HashSet::new();
    let mut repair_override_keys: Vec<&String> = repair_overrides.keys().collect();
    repair_override_keys.sort();
    let mut part_override_keys: Vec<&String> = part_overrides.keys().collect();
    part_override_keys.sort();

    fn find_part_override<'a>(
        part_name: &str,
        overrides: &'a HashMap<String, PartOverride>,
        sorted_keys: &[&'a String],
    ) -> Option<(&'a str, &'a PartOverride)> {
        if let Some((key, op)) = overrides.get_key_value(part_name) {
            return Some((key.as_str(), op));
        }
        for key in sorted_keys {
            if crate::zip_util::zip_part_names_equivalent(key.as_str(), part_name) {
                let op = overrides
                    .get(key.as_str())
                    .expect("override key came from override map");
                return Some((key.as_str(), op));
            }
        }
        None
    }

    for i in 0..archive.len() {
        let mut file = archive.by_index(i)?;
        if file.is_dir() {
            continue;
        }

        let raw_name = file.name().to_string();
        let canonical_name_cow = canonicalize_zip_entry_name(&raw_name);
        let canonical_name = canonical_name_cow.as_ref();

        // Track worksheet patch targets so we can report `MissingWorksheetPart` accurately even if
        // a caller overrides the part (e.g. removes or replaces it).
        missing_parts.remove(canonical_name);

        if recalc_policy.drop_calc_chain_on_formula_change && canonical_name == "xl/calcChain.xml" {
            // Drop calcChain.xml entirely when formulas change, matching the in-memory patcher.
            continue;
        }

        if let Some((override_key, override_op)) =
            find_part_override(canonical_name, &repair_overrides, &repair_override_keys)
        {
            applied_part_overrides.insert(override_key.to_string());
            match override_op {
                PartOverride::Remove => {
                    continue;
                }
                PartOverride::Replace(bytes) | PartOverride::Add(bytes) => {
                    zip.start_file(raw_name.clone(), options)?;
                    zip.write_all(bytes)?;
                    continue;
                }
            }
        }

        if let Some((override_key, override_op)) =
            find_part_override(canonical_name, part_overrides, &part_override_keys)
        {
            applied_part_overrides.insert(override_key.to_string());
            match override_op {
                PartOverride::Remove => {
                    continue;
                }
                PartOverride::Replace(bytes) | PartOverride::Add(bytes) => {
                    zip.start_file(raw_name.clone(), options)?;
                    zip.write_all(bytes)?;
                    continue;
                }
            }
        }

        let col_properties = col_properties_by_part.get(canonical_name);
        if col_properties.is_some() || effective_patches_by_part.contains_key(canonical_name) {
            let patches = effective_patches_by_part
                .get(canonical_name)
                .map(Vec::as_slice)
                .unwrap_or(&[]);
            zip.start_file(raw_name.clone(), options)?;
            let indices = shared_string_indices.get(canonical_name);
            let worksheet_meta = worksheet_metadata_by_part
                .get(canonical_name)
                .copied()
                .unwrap_or_default();
            patch_worksheet_xml_streaming(
                &mut file,
                &mut zip,
                canonical_name,
                patches,
                indices,
                col_properties,
                worksheet_meta,
                drop_vm_on_value_change,
                recalc_policy,
            )?;
        } else if let Some(bytes) = updated_parts.get(canonical_name) {
            zip.start_file(raw_name.clone(), options)?;
            zip.write_all(bytes)?;
        } else if shared_strings_part.as_deref() == Some(canonical_name)
            && shared_strings_updated.is_some()
        {
            zip.start_file(raw_name.clone(), options)?;
            zip.write_all(
                shared_strings_updated
                    .as_deref()
                    .expect("checked is_some above"),
            )?;
        } else if let Some(bytes) = pre_read_parts.get(canonical_name) {
            if should_patch_recalc_part(canonical_name, recalc_policy) {
                zip.start_file(raw_name.clone(), options)?;
                let bytes = maybe_patch_recalc_part(canonical_name, bytes, recalc_policy)?;
                zip.write_all(&bytes)?;
            } else {
                // We buffered this part earlier for metadata resolution, but it doesn't need to be
                // rewritten. Raw-copy it to avoid recompression.
                zip.raw_copy_file(file)?;
            }
        } else if let Some(updated) =
            patch_recalc_part_from_file(canonical_name, &mut file, recalc_policy)?
        {
            zip.start_file(raw_name.clone(), options)?;
            zip.write_all(&updated)?;
        } else {
            // Use raw copy to preserve bytes for unchanged parts and avoid a decompression /
            // recompression pass over large binary assets.
            zip.raw_copy_file(file)?;
        }
    }

    // Append any missing override parts deterministically.
    if !part_overrides.is_empty() || !repair_overrides.is_empty() {
        let mut names: BTreeSet<&str> = BTreeSet::new();
        for (name, op) in part_overrides.iter() {
            if matches!(op, PartOverride::Add(_) | PartOverride::Replace(_)) {
                names.insert(name.as_str());
            }
        }
        for (name, op) in repair_overrides.iter() {
            if matches!(op, PartOverride::Add(_) | PartOverride::Replace(_)) {
                names.insert(name.as_str());
            }
        }

        for name in names {
            if applied_part_overrides.contains(name) {
                continue;
            }
            let override_op = repair_overrides
                .get(name)
                .or_else(|| part_overrides.get(name));
            let Some(override_op) = override_op else {
                continue;
            };
            let bytes = match override_op {
                PartOverride::Replace(bytes) | PartOverride::Add(bytes) => bytes,
                PartOverride::Remove => continue,
            };
            zip.start_file(name.to_string(), options)?;
            zip.write_all(bytes)?;
        }
    }

    if let Some((missing, _)) = missing_parts.into_iter().next() {
        return Err(StreamingPatchError::MissingWorksheetPart(missing));
    }

    zip.finish()?;
    Ok(())
}

fn plan_package_repair_overrides<R: Read + Seek>(
    archive: &mut ZipArchive<R>,
    pre_read_parts: &HashMap<String, Vec<u8>>,
    updated_parts: &HashMap<String, Vec<u8>>,
    part_overrides: &HashMap<String, PartOverride>,
    recalc_policy: RecalcPolicy,
) -> Result<HashMap<String, PartOverride>, StreamingPatchError> {
    fn find_part_name<'a>(part_names: &'a BTreeSet<String>, candidate: &str) -> Option<&'a str> {
        if let Some(name) = part_names.get(candidate) {
            return Some(name.as_str());
        }
        part_names
            .iter()
            .find(|name| crate::zip_util::zip_part_names_equivalent(name.as_str(), candidate))
            .map(String::as_str)
    }

    fn effective_part_bytes<R: Read + Seek>(
        archive: &mut ZipArchive<R>,
        name: &str,
        pre_read_parts: &HashMap<String, Vec<u8>>,
        updated_parts: &HashMap<String, Vec<u8>>,
        part_overrides: &HashMap<String, PartOverride>,
        recalc_policy: RecalcPolicy,
    ) -> Result<Option<Vec<u8>>, StreamingPatchError> {
        let override_op = part_overrides.get(name).or_else(|| {
            part_overrides
                .iter()
                .find(|(k, _)| crate::zip_util::zip_part_names_equivalent(k.as_str(), name))
                .map(|(_, op)| op)
        });
        if let Some(op) = override_op {
            match op {
                PartOverride::Remove => return Ok(None),
                PartOverride::Replace(bytes) | PartOverride::Add(bytes) => {
                    let mut bytes = bytes.clone();
                    if should_patch_recalc_part(name, recalc_policy) {
                        bytes = maybe_patch_recalc_part(name, &bytes, recalc_policy)?;
                    }
                    return Ok(Some(bytes));
                }
            }
        }

        if let Some(bytes) = updated_parts.get(name) {
            let mut bytes = bytes.clone();
            if should_patch_recalc_part(name, recalc_policy) {
                bytes = maybe_patch_recalc_part(name, &bytes, recalc_policy)?;
            }
            return Ok(Some(bytes));
        }

        if let Some(bytes) = pre_read_parts.get(name) {
            let mut bytes = bytes.clone();
            if should_patch_recalc_part(name, recalc_policy) {
                bytes = maybe_patch_recalc_part(name, &bytes, recalc_policy)?;
            }
            return Ok(Some(bytes));
        }

        let Some(mut bytes) = read_zip_part_optional(archive, name)? else {
            return Ok(None);
        };
        if should_patch_recalc_part(name, recalc_policy) {
            bytes = maybe_patch_recalc_part(name, &bytes, recalc_policy)?;
        }
        Ok(Some(bytes))
    }

    // Determine the effective part name set after applying part overrides.
    let mut part_names: BTreeSet<String> = BTreeSet::new();
    for name in archive.file_names() {
        // `file_names()` includes directory entries; ignore them for content detection.
        if name.ends_with('/') {
            continue;
        }
        part_names.insert(canonicalize_zip_entry_name(name).into_owned());
    }
    for (name, op) in part_overrides {
        let canonical = canonicalize_zip_entry_name(name);
        match op {
            PartOverride::Remove => {
                part_names.remove(canonical.as_ref());
            }
            PartOverride::Replace(_) | PartOverride::Add(_) => {
                part_names.insert(canonical.into_owned());
            }
        }
    }

    // Detect whether the package contains image payloads that require `<Default>` entries in
    // `[Content_Types].xml`. This matches the conservative behavior of
    // `XlsxPackage::write_to(...)`: only add defaults for extensions that appear in the package.
    let mut has_vba_project = false;
    let mut has_vba_signature = false;
    let mut has_vba_data = false;
    let mut needs_png = false;
    let mut needs_jpg = false;
    let mut needs_jpeg = false;
    let mut needs_gif = false;
    let mut needs_webp = false;
    for name in &part_names {
        if !has_vba_project
            && crate::zip_util::zip_part_names_equivalent(name.as_str(), "xl/vbaProject.bin")
        {
            has_vba_project = true;
        }
        if !has_vba_signature
            && crate::zip_util::zip_part_names_equivalent(
                name.as_str(),
                "xl/vbaProjectSignature.bin",
            )
        {
            has_vba_signature = true;
        }
        if !has_vba_data
            && crate::zip_util::zip_part_names_equivalent(name.as_str(), "xl/vbaData.xml")
        {
            has_vba_data = true;
        }

        let lower = name.to_ascii_lowercase();
        if lower.ends_with(".png") {
            needs_png = true;
        } else if lower.ends_with(".jpg") {
            needs_jpg = true;
        } else if lower.ends_with(".jpeg") {
            needs_jpeg = true;
        } else if lower.ends_with(".gif") {
            needs_gif = true;
        } else if lower.ends_with(".webp") {
            needs_webp = true;
        }
    }

    if !has_vba_project && !(needs_png || needs_jpg || needs_jpeg || needs_gif || needs_webp) {
        return Ok(HashMap::new());
    }

    let content_types_key = find_part_name(&part_names, "[Content_Types].xml").map(ToString::to_string);
    let workbook_rels_key = has_vba_project
        .then(|| find_part_name(&part_names, "xl/_rels/workbook.xml.rels").map(ToString::to_string))
        .flatten();
    let vba_project_rels_key = (has_vba_project && has_vba_signature)
        .then(|| find_part_name(&part_names, "xl/_rels/vbaProject.bin.rels").map(ToString::to_string))
        .flatten();

    let content_types_original = effective_part_bytes(
        archive,
        "[Content_Types].xml",
        pre_read_parts,
        updated_parts,
        part_overrides,
        recalc_policy,
    )?;
    let workbook_rels_original = if has_vba_project {
        effective_part_bytes(
            archive,
            "xl/_rels/workbook.xml.rels",
            pre_read_parts,
            updated_parts,
            part_overrides,
            recalc_policy,
        )?
    } else {
        None
    };
    let vba_project_rels_original = if has_vba_project && has_vba_signature {
        effective_part_bytes(
            archive,
            "xl/_rels/vbaProject.bin.rels",
            pre_read_parts,
            updated_parts,
            part_overrides,
            recalc_policy,
        )?
    } else {
        None
    };

    // Run the existing in-memory repair logic on a minimal part map and convert the modified parts
    // into streaming part overrides.
    let mut parts: BTreeMap<String, Vec<u8>> = BTreeMap::new();
    if let Some(bytes) = content_types_original.clone() {
        parts.insert("[Content_Types].xml".to_string(), bytes);
    }
    if let Some(bytes) = workbook_rels_original.clone() {
        parts.insert("xl/_rels/workbook.xml.rels".to_string(), bytes);
    }
    if let Some(bytes) = vba_project_rels_original.clone() {
        parts.insert("xl/_rels/vbaProject.bin.rels".to_string(), bytes);
    }

    if has_vba_project {
        // Stub out macro payloads so `macro_repair` can check presence without inflating them.
        parts.insert("xl/vbaProject.bin".to_string(), Vec::new());
        if has_vba_signature {
            parts.insert("xl/vbaProjectSignature.bin".to_string(), Vec::new());
        }
        if has_vba_data {
            parts.insert("xl/vbaData.xml".to_string(), Vec::new());
        }

        crate::macro_repair::ensure_xlsm_content_types(&mut parts)?;
        crate::macro_repair::ensure_workbook_rels_has_vba(&mut parts)?;
        crate::macro_repair::ensure_vba_project_rels_has_signature(&mut parts)?;
    }

    if needs_png {
        crate::package::ensure_content_types_default(&mut parts, "png", "image/png")?;
    }
    if needs_jpg {
        crate::package::ensure_content_types_default(&mut parts, "jpg", "image/jpeg")?;
    }
    if needs_jpeg {
        crate::package::ensure_content_types_default(&mut parts, "jpeg", "image/jpeg")?;
    }
    if needs_gif {
        crate::package::ensure_content_types_default(&mut parts, "gif", "image/gif")?;
    }
    if needs_webp {
        crate::package::ensure_content_types_default(&mut parts, "webp", "image/webp")?;
    }

    let mut overrides: HashMap<String, PartOverride> = HashMap::new();

    if let Some(original) = content_types_original {
        if let Some(updated) = parts.get("[Content_Types].xml") {
            if updated.as_slice() != original.as_slice() {
                overrides.insert(
                    content_types_key.unwrap_or_else(|| "[Content_Types].xml".to_string()),
                    PartOverride::Replace(updated.clone()),
                );
            }
        }
    }

    if let Some(original) = workbook_rels_original {
        if let Some(updated) = parts.get("xl/_rels/workbook.xml.rels") {
            if updated.as_slice() != original.as_slice() {
                overrides.insert(
                    workbook_rels_key.unwrap_or_else(|| "xl/_rels/workbook.xml.rels".to_string()),
                    PartOverride::Replace(updated.clone()),
                );
            }
        }
    }

    if has_vba_project && has_vba_signature {
        if let Some(updated) = parts.get("xl/_rels/vbaProject.bin.rels") {
            let original = vba_project_rels_original.as_deref();
            if original != Some(updated.as_slice()) {
                overrides.insert(
                    vba_project_rels_key.unwrap_or_else(|| "xl/_rels/vbaProject.bin.rels".to_string()),
                    PartOverride::Replace(updated.clone()),
                );
            }
        }
    }

    Ok(overrides)
}

fn streaming_patches_remove_existing_formulas<R: Read + Seek>(
    archive: &mut ZipArchive<R>,
    patches_by_part: &HashMap<String, Vec<WorksheetCellPatch>>,
) -> Result<bool, StreamingPatchError> {
    for (worksheet_part, patches) in patches_by_part {
        let mut target_cells: HashSet<String> = HashSet::new();
        for patch in patches {
            if !formula_is_material(patch.formula.as_deref()) {
                target_cells.insert(patch.cell.to_a1());
            }
        }
        if target_cells.is_empty() {
            continue;
        }

        let file = match open_zip_part(archive, worksheet_part) {
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

fn patch_recalc_part_from_file(
    name: &str,
    file: &mut zip::read::ZipFile<'_>,
    recalc_policy: RecalcPolicy,
) -> Result<Option<Vec<u8>>, StreamingPatchError> {
    if !should_patch_recalc_part(name, recalc_policy) {
        return Ok(None);
    }

    let buf = read_zip_file_bytes_with_limit(file, name, DEFAULT_MAX_ZIP_PART_BYTES)?;
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

fn list_zip_part_names<R: Read + Seek>(
    archive: &mut ZipArchive<R>,
) -> Result<HashSet<String>, StreamingPatchError> {
    let mut out = HashSet::new();
    for i in 0..archive.len() {
        let file = archive.by_index(i)?;
        if file.is_dir() {
            continue;
        }
        let name = file.name();
        out.insert(canonicalize_zip_entry_name(name).into_owned());
    }
    Ok(out)
}

fn find_zip_part_name(part_names: &HashSet<String>, candidate: &str) -> Option<String> {
    if part_names.contains(candidate) {
        return Some(candidate.to_string());
    }
    part_names
        .iter()
        .find(|name| crate::zip_util::zip_part_names_equivalent(name.as_str(), candidate))
        .cloned()
}

fn read_zip_part<R: Read + Seek>(
    archive: &mut ZipArchive<R>,
    name: &str,
    cache: &mut HashMap<String, Vec<u8>>,
) -> Result<Vec<u8>, StreamingPatchError> {
    if let Some(bytes) = cache.get(name) {
        return Ok(bytes.clone());
    }
    let mut file = open_zip_part(archive, name)?;
    let buf = read_zip_file_bytes_with_limit(&mut file, name, DEFAULT_MAX_ZIP_PART_BYTES)?;
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
    vm: Option<Option<u32>>,
    cm: Option<Option<u32>>,
    shared_string_idx: Option<u32>,
    clear_cached_value: bool,
    material_for_insertion: bool,
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
    col_properties: Option<&BTreeMap<u32, ColProperties>>,
    worksheet_meta: WorksheetXmlMetadata,
    drop_vm_on_value_change: bool,
    recalc_policy: RecalcPolicy,
) -> Result<(), StreamingPatchError> {
    let has_cell_patches = !patches.is_empty();
    let patch_bounds = bounds_for_patches(patches);

    // Only insert `<dimension>` when applying cell patches. Column-metadata-only edits should not
    // introduce unrelated structural changes.
    let dimension_ref_to_insert = if has_cell_patches && !worksheet_meta.has_dimension {
        union_bounds(worksheet_meta.existing_used_range, patch_bounds).map(bounds_to_dimension_ref)
    } else {
        None
    };
    let insert_dimension_after_sheet_pr =
        dimension_ref_to_insert.is_some() && worksheet_meta.has_sheet_pr;
    let insert_dimension_at_worksheet_start =
        dimension_ref_to_insert.is_some() && !worksheet_meta.has_sheet_pr;

    let mut patches_by_row: BTreeMap<u32, Vec<CellPatchInternal>> = BTreeMap::new();
    for patch in patches {
        let row_1 = patch.cell.row + 1;
        let col_0 = patch.cell.col;
        let shared_string_idx =
            shared_string_indices.and_then(|m| m.get(&(patch.cell.row, patch.cell.col)).copied());
        let clear_cached_value = recalc_policy.clear_cached_values_on_formula_change
            && formula_is_material(patch.formula.as_deref());
        let material_for_insertion = patch_is_material_for_insertion(patch);
        patches_by_row
            .entry(row_1)
            .or_default()
            .push(CellPatchInternal {
                row_1,
                col_0,
                value: patch.value.clone(),
                formula: patch.formula.clone(),
                xf_index: patch.xf_index,
                vm: patch.vm,
                cm: patch.cm,
                shared_string_idx,
                clear_cached_value,
                material_for_insertion,
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
    let mut cols_written = false;

    let mut row_state: Option<RowState> = None;
    let mut in_cell = false;

    loop {
        let event = reader.read_event_into(&mut buf)?;
        match event {
            Event::Eof => break,
            Event::Start(ref e)
                if col_properties.is_some() && local_name(e.name().as_ref()) == b"cols" =>
            {
                let name = e.name();
                let prefix =
                    element_prefix(name.as_ref()).and_then(|p| std::str::from_utf8(p).ok());
                let mut attrs_by_col = {
                    let mut cols_buf = Vec::new();
                    parse_cols_attribute_map_from_reader(&mut reader, &mut cols_buf)?
                };
                crate::patch::merge_col_properties_into_attrs_by_col(
                    &mut attrs_by_col,
                    col_properties.expect("checked is_some above"),
                );
                if !cols_written {
                    let cols_xml = crate::patch::render_cols_xml_from_attrs_by_col(prefix, &attrs_by_col);
                    if !cols_xml.is_empty() {
                        writer.get_mut().write_all(cols_xml.as_bytes())?;
                    }
                    cols_written = true;
                }
            }
            Event::Empty(ref e)
                if col_properties.is_some() && local_name(e.name().as_ref()) == b"cols" =>
            {
                let name = e.name();
                let prefix =
                    element_prefix(name.as_ref()).and_then(|p| std::str::from_utf8(p).ok());
                let mut attrs_by_col = BTreeMap::new();
                crate::patch::merge_col_properties_into_attrs_by_col(
                    &mut attrs_by_col,
                    col_properties.expect("checked is_some above"),
                );
                if !cols_written {
                    let cols_xml = crate::patch::render_cols_xml_from_attrs_by_col(prefix, &attrs_by_col);
                    if !cols_xml.is_empty() {
                        writer.get_mut().write_all(cols_xml.as_bytes())?;
                    }
                    cols_written = true;
                }
            }

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
                if let Some(col_properties) = col_properties {
                    if !cols_written {
                        let name = e.name();
                        let prefix =
                            element_prefix(name.as_ref()).and_then(|p| std::str::from_utf8(p).ok());
                        let cols_xml = crate::patch::render_cols_xml(col_properties, prefix);
                        if !cols_xml.is_empty() {
                            writer.get_mut().write_all(cols_xml.as_bytes())?;
                            cols_written = true;
                        }
                    }
                }
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
                if let Some(col_properties) = col_properties {
                    if !cols_written {
                        let name = e.name();
                        let prefix =
                            element_prefix(name.as_ref()).and_then(|p| std::str::from_utf8(p).ok());
                        let cols_xml = crate::patch::render_cols_xml(col_properties, prefix);
                        if !cols_xml.is_empty() {
                            writer.get_mut().write_all(cols_xml.as_bytes())?;
                            cols_written = true;
                        }
                    }
                }
                saw_sheet_data = true;
                if patch_bounds.is_none() {
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
                    write_pending_rows(
                        &mut writer,
                        &mut patches_by_row,
                        sheet_prefix.as_deref(),
                        worksheet_meta.sheet_uses_row_spans,
                    )?;
                    writer.write_event(Event::End(BytesEnd::new(sheet_data_tag.as_str())))?;
                }
            }
            Event::End(ref e) if local_name(e.name().as_ref()) == b"sheetData" => {
                // Flush any remaining patch rows at the end of sheetData.
                write_pending_rows(
                    &mut writer,
                    &mut patches_by_row,
                    sheet_prefix.as_deref(),
                    worksheet_meta.sheet_uses_row_spans,
                )?;
                in_sheet_data = false;
                writer.write_event(Event::End(e.to_owned()))?;
            }
            Event::End(ref e) if local_name(e.name().as_ref()) == b"worksheet" => {
                if let Some(col_properties) = col_properties {
                    if !cols_written {
                        let prefix = if worksheet_has_default_ns {
                            None
                        } else {
                            worksheet_prefix.as_deref()
                        };
                        let cols_xml = crate::patch::render_cols_xml(col_properties, prefix);
                        if !cols_xml.is_empty() {
                            writer.get_mut().write_all(cols_xml.as_bytes())?;
                            cols_written = true;
                        }
                    }
                }
                if !saw_sheet_data && !patches_by_row.is_empty() {
                    saw_sheet_data = true;
                    let sheet_prefix = if worksheet_has_default_ns {
                        None
                    } else {
                        worksheet_prefix.as_deref()
                    };
                    let sheet_data_tag = prefixed_tag(sheet_prefix, "sheetData");
                    writer.write_event(Event::Start(BytesStart::new(sheet_data_tag.as_str())))?;
                    write_pending_rows(
                        &mut writer,
                        &mut patches_by_row,
                        sheet_prefix,
                        worksheet_meta.sheet_uses_row_spans,
                    )?;
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
                        write_inserted_row(
                            &mut writer,
                            next_row,
                            &pending,
                            sheet_prefix.as_deref(),
                            worksheet_meta.sheet_uses_row_spans,
                        )?;
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
                        write_inserted_row(
                            &mut writer,
                            next_row,
                            &pending,
                            sheet_prefix.as_deref(),
                            worksheet_meta.sheet_uses_row_spans,
                        )?;
                    } else {
                        break;
                    }
                }

                if let Some(mut pending) = patches_by_row.remove(&row_1) {
                    pending.sort_by_key(|p| p.col_0);
                    if pending.iter().any(|p| p.material_for_insertion) {
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
                        // No material patches; preserve the empty row unchanged.
                        writer.write_event(Event::Empty(e.to_owned()))?;
                    }
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
                if in_sheet_data
                    && row_state.is_some()
                    && local_name(e.name().as_ref()) == b"c" =>
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
                    patch_existing_cell(
                        &mut reader,
                        &mut writer,
                        e,
                        &cell_ref,
                        &patch,
                        drop_vm_on_value_change,
                    )?;
                } else {
                    writer.write_event(Event::Start(e.to_owned()))?;
                    in_cell = true;
                }
            }
            Event::Empty(ref e)
                if in_sheet_data
                    && row_state.is_some()
                    && local_name(e.name().as_ref()) == b"c" =>
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
                if in_sheet_data
                    && row_state.is_some()
                    && in_cell
                    && local_name(e.name().as_ref()) == b"c" =>
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

fn parse_cols_attribute_map_from_reader<R: BufRead>(
    reader: &mut Reader<R>,
    buf: &mut Vec<u8>,
) -> Result<BTreeMap<u32, BTreeMap<String, String>>, StreamingPatchError> {
    let mut attrs_by_col: BTreeMap<u32, BTreeMap<String, String>> = BTreeMap::new();

    // We're inside `<cols>` (the start tag has been consumed); consume events until `</cols>`.
    let mut depth: usize = 1;
    loop {
        match reader.read_event_into(buf)? {
            Event::Eof => {
                return Err(StreamingPatchError::Xlsx(crate::XlsxError::Invalid(
                    "unexpected EOF while parsing <cols> section".to_string(),
                )))
            }
            Event::Start(e) => {
                if local_name(e.name().as_ref()) == b"col" {
                    parse_col_element_attrs(&e, &mut attrs_by_col)?;
                }
                depth += 1;
            }
            Event::Empty(e) => {
                if local_name(e.name().as_ref()) == b"col" {
                    parse_col_element_attrs(&e, &mut attrs_by_col)?;
                }
            }
            Event::End(e) => {
                depth = depth.saturating_sub(1);
                if depth == 0 && local_name(e.name().as_ref()) == b"cols" {
                    break;
                }
            }
            _ => {}
        }
        buf.clear();
    }

    Ok(attrs_by_col)
}

fn parse_col_element_attrs(
    e: &BytesStart<'_>,
    attrs_by_col: &mut BTreeMap<u32, BTreeMap<String, String>>,
) -> Result<(), StreamingPatchError> {
    let mut min: Option<u32> = None;
    let mut max: Option<u32> = None;
    let mut element_attrs: Vec<(String, String)> = Vec::new();

    for attr in e.attributes() {
        let attr = attr?;
        let key_bytes = attr.key.as_ref();
        let key = match std::str::from_utf8(key_bytes) {
            Ok(s) => s.to_string(),
            Err(_) => continue,
        };
        let val = attr.unescape_value()?.into_owned();
        match key_bytes {
            b"min" => min = val.parse().ok(),
            b"max" => max = val.parse().ok(),
            _ => element_attrs.push((key, val)),
        }
    }

    let Some(min_1) = min else {
        return Ok(());
    };
    let max_1 = max.unwrap_or(min_1).min(formula_model::EXCEL_MAX_COLS);
    if min_1 == 0 || max_1 == 0 || min_1 > formula_model::EXCEL_MAX_COLS {
        return Ok(());
    }

    for col_1 in min_1..=max_1 {
        let entry = attrs_by_col.entry(col_1).or_default();
        for (k, v) in &element_attrs {
            entry.insert(k.clone(), v.clone());
        }
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
    let mut iter = patches.iter().filter(|p| p.material_for_insertion);
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
    sheet_uses_row_spans: bool,
) -> Result<(), StreamingPatchError> {
    while let Some((&row_1, _)) = patches_by_row.iter().next() {
        let pending = patches_by_row.remove(&row_1).unwrap_or_default();
        write_inserted_row(writer, row_1, &pending, prefix, sheet_uses_row_spans)?;
    }
    Ok(())
}

fn write_inserted_row<W: Write>(
    writer: &mut Writer<W>,
    row_1: u32,
    patches: &[CellPatchInternal],
    prefix: Option<&str>,
    sheet_uses_row_spans: bool,
) -> Result<(), StreamingPatchError> {
    if !patches.iter().any(|p| p.material_for_insertion) {
        return Ok(());
    }
    let row_tag = prefixed_tag(prefix, "row");
    let mut row = BytesStart::new(row_tag.as_str());
    let row_num = row_1.to_string();
    row.push_attribute(("r", row_num.as_str()));
    let spans = if sheet_uses_row_spans {
        spans_for_patches(patches)
            .map(|(min_col_1, max_col_1)| format_row_spans(min_col_1, max_col_1))
    } else {
        None
    };
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
        if !patch.material_for_insertion {
            continue;
        }
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
        if !patch.material_for_insertion {
            continue;
        }
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
            if patch.material_for_insertion {
                let cell_ref = CellRef::new(state.row_1 - 1, patch.col_0);
                write_patched_cell::<W>(writer, None, &cell_ref, patch, prefix)?;
            }
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
        if patch.material_for_insertion {
            let cell_ref = CellRef::new(state.row_1 - 1, patch.col_0);
            write_patched_cell::<W>(writer, None, &cell_ref, patch, prefix)?;
        }
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
    drop_vm_on_value_change: bool,
) -> Result<(), StreamingPatchError> {
    let patch_formula = match patch.formula.as_deref() {
        Some(formula) if formula_is_material(Some(formula)) => Some(formula),
        _ => None,
    };
    let style_override = patch.xf_index;

    let cell_tag = String::from_utf8_lossy(cell_start.name().as_ref()).into_owned();
    let prefix = cell_tag.rsplit_once(':').map(|(p, _)| p);
    let f_tag = prefixed_tag(prefix, "f");
    let v_tag = prefixed_tag(prefix, "v");
    let is_tag = prefixed_tag(prefix, "is");
    let t_tag = prefixed_tag(prefix, "t");

    let mut existing_t: Option<String> = None;
    let mut original_has_vm = false;
    for attr in cell_start.attributes() {
        let attr = attr?;
        if attr.key.as_ref() == b"t" {
            existing_t = Some(attr.unescape_value()?.into_owned());
        } else if attr.key.as_ref() == b"vm" {
            original_has_vm = true;
        }
    }

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

    let clear_cached_value = patch.clear_cached_value && patch_formula.is_some();
    let value_eq = cell_value_semantics_eq(
        existing_t.as_deref(),
        &inner_events,
        &patch.value,
        patch.shared_string_idx,
    )?;
    let update_value = !value_eq || clear_cached_value;

    // `vm="..."` points into `xl/metadata.xml` value metadata (rich values / images-in-cell).
    //
    // We generally preserve it for fidelity. The main exception is the in-cell image placeholder
    // representation (commonly `t="e"` + `<v>#VALUE!</v>`, but some producers use a numeric `0`):
    // when a patch changes the cached value away from the placeholder semantics, we must drop `vm`
    // to avoid leaving a dangling rich-data pointer.
    //
    // Additionally, when patching incomplete workbook packages (see `drop_vm_on_value_change`),
    // drop `vm` whenever the cached value semantics change (unless the caller explicitly overrides
    // `vm`).
    let patch_is_rich_value_placeholder =
        matches!(&patch.value, CellValue::Error(ErrorValue::Value))
            || matches!(&patch.value, CellValue::Number(n) if *n == 0.0);
    let existing_is_rich_value_placeholder = if original_has_vm {
        cell_is_rich_value_placeholder(existing_t.as_deref(), &inner_events)?
    } else {
        false
    };
    let drop_vm = if patch.vm.is_none() {
        (existing_is_rich_value_placeholder && !patch_is_rich_value_placeholder)
            || (drop_vm_on_value_change && (clear_cached_value || !value_eq))
    } else {
        false
    };

    let mut c = BytesStart::new(cell_tag.as_str());
    let mut has_r = false;
    for attr in cell_start.attributes() {
        let attr = attr?;
        if attr.key.as_ref() == b"t" && update_value {
            continue;
        }
        if attr.key.as_ref() == b"s" && style_override.is_some() {
            continue;
        }
        if attr.key.as_ref() == b"vm" && (drop_vm || patch.vm.is_some()) {
            continue;
        }
        if attr.key.as_ref() == b"cm" && patch.cm.is_some() {
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

    let (cell_t_owned, body_kind) = if update_value {
        cell_representation(
            &patch.value,
            patch_formula,
            existing_t.as_deref(),
            patch.shared_string_idx,
        )?
    } else {
        (None, CellBodyKind::None)
    };

    if let Some(xf_index) = style_override {
        if xf_index != 0 {
            let xf = xf_index.to_string();
            c.push_attribute(("s", xf.as_str()));
        }
    }

    if update_value {
        if let Some(t) = cell_t_owned.as_deref() {
            c.push_attribute(("t", t));
        }
    }

    let vm_value = patch.vm.flatten().map(|vm| vm.to_string());
    if let Some(vm) = vm_value.as_deref() {
        c.push_attribute(("vm", vm));
    }
    let cm_value = patch.cm.flatten().map(|cm| cm.to_string());
    if let Some(cm) = cm_value.as_deref() {
        c.push_attribute(("cm", cm));
    }

    writer.write_event(Event::Start(c))?;

    write_patched_cell_children(
        writer,
        &inner_events,
        patch_formula,
        update_value,
        &body_kind,
        clear_cached_value,
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
    update_value: bool,
    body_kind: &CellBodyKind,
    clear_cached_value: bool,
    f_tag: &str,
    v_tag: &str,
    is_tag: &str,
    t_tag: &str,
) -> Result<(), StreamingPatchError> {
    let mut formula_written = patch_formula.is_none();
    let mut value_written =
        !update_value || matches!(body_kind, CellBodyKind::None) || clear_cached_value;
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
                        write_formula_element(
                            writer,
                            Some(e),
                            formula,
                            detach_shared,
                            tag.as_str(),
                        )?;
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
                        write_formula_element(
                            writer,
                            Some(e),
                            formula,
                            detach_shared,
                            tag.as_str(),
                        )?;
                        formula_written = true;
                    }
                }
                idx += 1;
                continue;
            }
            Event::Start(e)
                if local_name(e.name().as_ref()) == b"v"
                    || local_name(e.name().as_ref()) == b"is" =>
            {
                saw_value = true;

                if !formula_written {
                    if let Some(formula) = patch_formula {
                        // Original cell has no <f> before the value; insert one.
                        write_formula_element(writer, None, formula, false, f_tag)?;
                        formula_written = true;
                    }
                }
                if update_value && !value_written {
                    write_value_element(writer, body_kind, v_tag, is_tag, t_tag)?;
                    value_written = true;
                }

                if update_value || clear_cached_value {
                    idx = skip_owned_subtree(inner_events, idx);
                } else {
                    idx = write_owned_subtree(writer, inner_events, idx)?;
                }
                continue;
            }
            Event::Empty(e)
                if local_name(e.name().as_ref()) == b"v"
                    || local_name(e.name().as_ref()) == b"is" =>
            {
                saw_value = true;

                if !formula_written {
                    if let Some(formula) = patch_formula {
                        write_formula_element(writer, None, formula, false, f_tag)?;
                        formula_written = true;
                    }
                }
                if update_value && !value_written {
                    write_value_element(writer, body_kind, v_tag, is_tag, t_tag)?;
                    value_written = true;
                }

                if !update_value && !clear_cached_value {
                    writer.write_event(Event::Empty(e.clone()))?;
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
                if update_value && !value_written && !saw_value {
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
    if update_value && !value_written {
        write_value_element(writer, body_kind, v_tag, is_tag, t_tag)?;
    }

    Ok(())
}

fn write_owned_subtree<W: Write>(
    writer: &mut Writer<W>,
    events: &[Event<'static>],
    mut idx: usize,
) -> Result<usize, StreamingPatchError> {
    match &events[idx] {
        Event::Start(_) => {
            let mut depth = 0usize;
            while idx < events.len() {
                let ev = events[idx].clone();
                match &ev {
                    Event::Start(_) => depth += 1,
                    Event::End(_) => {
                        depth = depth.saturating_sub(1);
                    }
                    _ => {}
                }
                writer.write_event(ev)?;
                idx += 1;
                if depth == 0 {
                    break;
                }
            }
            Ok(idx)
        }
        _ => {
            writer.write_event(events[idx].clone())?;
            Ok(idx + 1)
        }
    }
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
    let patch_formula = match patch.formula.as_deref() {
        Some(formula) if formula_is_material(Some(formula)) => Some(formula),
        _ => None,
    };
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
    // This codepath only handles `<c .../>` (empty) cells, so there is no cached `<v>` value to
    // inspect. The embedded-image placeholder representation we special-case elsewhere is
    // `t="e"` + `<v>#VALUE!</v>`, so here we only drop `vm` when explicitly overridden by the
    // patch.
    let drop_vm = false;

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
            if attr.key.as_ref() == b"vm" && (drop_vm || patch.vm.is_some()) {
                continue;
            }
            if attr.key.as_ref() == b"cm" && patch.cm.is_some() {
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

    let vm_value = patch.vm.flatten().map(|vm| vm.to_string());
    if let Some(vm) = vm_value.as_deref() {
        c.push_attribute(("vm", vm));
    }
    let cm_value = patch.cm.flatten().map(|cm| cm.to_string());
    if let Some(cm) = cm_value.as_deref() {
        c.push_attribute(("cm", cm));
    }

    writer.write_event(Event::Start(c))?;

    if let Some(formula) = patch_formula {
        write_formula_element(writer, None, formula, false, &formula_tag)?;
    }

    if !(patch.clear_cached_value && patch_formula.is_some()) {
        write_value_element(writer, &body_kind, &v_tag, &is_tag, &t_tag)?;
    }

    writer.write_event(Event::End(BytesEnd::new(cell_tag_owned.as_str())))?;
    Ok(())
}

fn extract_cell_v_text(events: &[Event<'static>]) -> Result<Option<String>, StreamingPatchError> {
    let mut in_v = false;
    let mut out = String::new();

    for ev in events {
        match ev {
            Event::Start(e) if local_name(e.name().as_ref()) == b"v" => {
                in_v = true;
                out.clear();
            }
            Event::End(e) if local_name(e.name().as_ref()) == b"v" => {
                if in_v {
                    return Ok(Some(out));
                }
                in_v = false;
            }
            Event::Empty(e) if local_name(e.name().as_ref()) == b"v" => {
                return Ok(Some(String::new()))
            }
            Event::Text(t) if in_v => out.push_str(&t.unescape()?.into_owned()),
            Event::CData(t) if in_v => out.push_str(&String::from_utf8_lossy(t.as_ref())),
            _ => {}
        }
    }

    Ok(None)
}

#[allow(dead_code)]
fn extract_cell_inline_string_text(
    events: &[Event<'static>],
) -> Result<Option<String>, StreamingPatchError> {
    fn is_visible_inline_string_t(stack: &[Vec<u8>]) -> bool {
        // stack: ["is", ... , "t"]
        if !stack.last().is_some_and(|n| n.as_slice() == b"t") {
            return false;
        }
        // `<rPh>` phonetic guide text is not visible.
        if stack.iter().any(|n| n.as_slice() == b"rPh") {
            return false;
        }
        // Visible text lives in either:
        // - <is><t>...</t></is>
        // - <is><r><t>...</t></r></is>
        if stack.len() == 2 && stack[0].as_slice() == b"is" {
            return true;
        }
        if stack.len() >= 3 && stack[0].as_slice() == b"is" && stack[1].as_slice() == b"r" {
            return true;
        }
        false
    }

    let mut in_is = false;
    let mut stack: Vec<Vec<u8>> = Vec::new();
    let mut in_visible_t = false;
    let mut out = String::new();

    for ev in events {
        match ev {
            Event::Start(e) => {
                let name = local_name(e.name().as_ref()).to_vec();

                if !in_is {
                    if name.as_slice() == b"is" {
                        in_is = true;
                        stack.clear();
                        stack.push(b"is".to_vec());
                        out.clear();
                        in_visible_t = false;
                    }
                    continue;
                }

                // Inside `<is>...</is>`.
                stack.push(name.clone());
                if name.as_slice() == b"t" && is_visible_inline_string_t(&stack) {
                    in_visible_t = true;
                }
            }
            Event::End(e) => {
                if !in_is {
                    continue;
                }
                let name = local_name(e.name().as_ref()).to_vec();
                if name.as_slice() == b"t" && in_visible_t {
                    in_visible_t = false;
                }
                if name.as_slice() == b"is" {
                    return Ok(Some(out));
                }

                // Best-effort: assume well-formed nesting and pop once.
                stack.pop();
                if stack.is_empty() {
                    in_is = false;
                    in_visible_t = false;
                }
            }
            Event::Empty(e) => {
                let name = local_name(e.name().as_ref()).to_vec();
                if !in_is {
                    if name.as_slice() == b"is" {
                        return Ok(Some(String::new()));
                    }
                    continue;
                }
                // empty <t/> contributes no visible text
            }
            Event::Text(t) if in_visible_t => out.push_str(&t.unescape()?.into_owned()),
            Event::CData(t) if in_visible_t => out.push_str(&String::from_utf8_lossy(t.as_ref())),
            _ => {}
        }
    }

    Ok(None)
}

fn cell_is_rich_value_placeholder(
    existing_t: Option<&str>,
    inner_events: &[Event<'static>],
) -> Result<bool, StreamingPatchError> {
    let Some(v) = extract_cell_v_text(inner_events)? else {
        return Ok(false);
    };
    let v = v.trim();

    let existing_t = existing_t.map(|t| t.trim()).filter(|t| !t.is_empty());
    if existing_t.is_some_and(|t| t.eq_ignore_ascii_case("e")) {
        return Ok(v.eq_ignore_ascii_case(ErrorValue::Value.as_str()));
    }

    // Some in-cell image placeholder workbooks store the cached value as a number `0` (with no
    // `t=` attribute, which implies SpreadsheetML numeric cells).
    if existing_t.is_none() {
        if v.eq_ignore_ascii_case(ErrorValue::Value.as_str()) {
            return Ok(true);
        }
        if let Ok(n) = v.parse::<f64>() {
            return Ok(n == 0.0);
        }
    }

    Ok(false)
}

#[allow(dead_code)]
fn cell_value_semantics_eq(
    existing_t: Option<&str>,
    inner_events: &[Event<'static>],
    patch_value: &CellValue,
    patch_shared_string_idx: Option<u32>,
) -> Result<bool, StreamingPatchError> {
    // Mirror `cell_representation`'s "degrade richer types to strings" behavior so style-only
    // patches on those values do not unnecessarily drop `vm`.
    match patch_value {
        CellValue::Entity(entity) => {
            return cell_value_semantics_eq(
                existing_t,
                inner_events,
                &CellValue::String(entity.display_value.clone()),
                patch_shared_string_idx,
            );
        }
        CellValue::Record(record) => {
            return cell_value_semantics_eq(
                existing_t,
                inner_events,
                &CellValue::String(record.to_string()),
                patch_shared_string_idx,
            );
        }
        CellValue::Image(image) => {
            if let Some(alt) = image.alt_text.as_deref().filter(|s| !s.is_empty()) {
                return cell_value_semantics_eq(
                    existing_t,
                    inner_events,
                    &CellValue::String(alt.to_string()),
                    patch_shared_string_idx,
                );
            }
            return cell_value_semantics_eq(
                existing_t,
                inner_events,
                &CellValue::Empty,
                patch_shared_string_idx,
            );
        }
        _ => {}
    }

    // Cached value semantics live in either `<v>` (numbers, bools, errors, shared string indices,
    // `t="str"`, etc) or `<is>` (inline strings).
    let v_text = extract_cell_v_text(inner_events)?;

    match patch_value {
        CellValue::Empty => {
            if existing_t == Some("inlineStr") {
                return Ok(extract_cell_inline_string_text(inner_events)?
                    .unwrap_or_default()
                    .is_empty());
            }
            Ok(v_text.unwrap_or_default().is_empty())
        }
        CellValue::Number(n) => {
            if matches!(existing_t, Some("b" | "e" | "s" | "str" | "inlineStr")) {
                return Ok(false);
            }
            let Some(v) = v_text else {
                return Ok(false);
            };
            Ok(v.trim().parse::<f64>().ok() == Some(*n))
        }
        CellValue::Boolean(b) => {
            if existing_t != Some("b") {
                return Ok(false);
            }
            let Some(v) = v_text else {
                return Ok(false);
            };
            let normalized = v.trim();
            let existing = normalized == "1" || normalized.eq_ignore_ascii_case("true");
            Ok(existing == *b)
        }
        CellValue::Error(err) => {
            if existing_t != Some("e") {
                return Ok(false);
            }
            let Some(v) = v_text else {
                return Ok(false);
            };
            Ok(v.trim().parse::<formula_model::ErrorValue>().ok() == Some(*err))
        }
        CellValue::String(s) => {
            match existing_t {
                Some("inlineStr") => {
                    Ok(extract_cell_inline_string_text(inner_events)?.unwrap_or_default() == *s)
                }
                Some("s") => {
                    let Some(idx_text) = v_text else {
                        return Ok(false);
                    };
                    let existing_idx = idx_text.trim().parse::<u32>().ok();
                    Ok(existing_idx.is_some_and(|idx| patch_shared_string_idx == Some(idx)))
                }
                Some("str") => Ok(v_text.unwrap_or_default() == *s),
                // Treat unknown/other `t=` values as raw `<v>`-text comparisons.
                Some(_) => Ok(v_text.unwrap_or_default() == *s),
                None => Ok(false),
            }
        }
        CellValue::RichText(rich) => {
            // Best-effort: compare rich text values by their visible text, and shared-string index
            // when available.
            let s = rich.text.as_str();
            match existing_t {
                Some("inlineStr") => {
                    Ok(extract_cell_inline_string_text(inner_events)?.unwrap_or_default() == s)
                }
                Some("s") => {
                    let Some(idx_text) = v_text else {
                        return Ok(false);
                    };
                    let existing_idx = idx_text.trim().parse::<u32>().ok();
                    Ok(existing_idx.is_some_and(|idx| patch_shared_string_idx == Some(idx)))
                }
                Some("str") => Ok(v_text.unwrap_or_default() == s),
                Some(_) => Ok(v_text.unwrap_or_default() == s),
                None => Ok(false),
            }
        }
        _other => {
            // Treat other value types (unsupported for streaming patching) as changed.
            // This is intentionally conservative; it will preserve the existing behavior
            // of dropping `vm` for full-package patches.
            Ok(false)
        }
    }
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
                        return Ok((
                            Some("inlineStr".to_string()),
                            CellBodyKind::InlineStr(s.clone()),
                        ));
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

            if formula_is_material(formula) {
                Ok((Some("str".to_string()), CellBodyKind::V(s.clone())))
            } else {
                Ok((
                    Some("inlineStr".to_string()),
                    CellBodyKind::InlineStr(s.clone()),
                ))
            }
        }
        CellValue::Entity(entity) => {
            let degraded = CellValue::String(entity.display_value.clone());
            cell_representation(&degraded, formula, existing_t, shared_string_idx)
        }
        CellValue::Record(record) => {
            let degraded = CellValue::String(record.to_string());
            cell_representation(&degraded, formula, existing_t, shared_string_idx)
        }
        CellValue::Image(image) => {
            if let Some(alt) = image.alt_text.as_deref().filter(|s| !s.is_empty()) {
                let degraded = CellValue::String(alt.to_string());
                cell_representation(&degraded, formula, existing_t, shared_string_idx)
            } else {
                Ok((None, CellBodyKind::None))
            }
        }
        CellValue::RichText(rich) => {
            if let Some(existing_t) = existing_t {
                if should_preserve_unknown_t(existing_t) {
                    return Ok((
                        Some(existing_t.to_string()),
                        CellBodyKind::V(rich.text.clone()),
                    ));
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
    name.iter()
        .rposition(|b| *b == b':')
        .map(|idx| &name[..idx])
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

fn patch_is_material_for_insertion(patch: &WorksheetCellPatch) -> bool {
    is_material_cell_patch_for_insertion(&patch.value, patch.formula.as_deref(), patch.xf_index)
        || patch.vm.flatten().is_some()
        || patch.cm.flatten().is_some()
}

fn is_material_cell_patch_for_insertion(
    value: &CellValue,
    formula: Option<&str>,
    xf_index: Option<u32>,
) -> bool {
    !matches!(value, CellValue::Empty)
        || formula_is_material(formula)
        || xf_index.is_some_and(|xf| xf != 0)
}

fn formula_is_material(formula: Option<&str>) -> bool {
    let Some(formula) = formula else {
        return false;
    };
    !crate::formula_text::normalize_display_formula(formula).is_empty()
}

fn bounds_for_patches(patches: &[WorksheetCellPatch]) -> Option<PatchBounds> {
    let mut iter = patches
        .iter()
        .filter(|p| patch_is_material_for_insertion(p));
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

#[cfg(test)]
mod tests {
    use super::*;

    use std::io::{Cursor, Write};

    use formula_model::{CellRef, CellValue, Style, StyleTable};
    use zip::write::FileOptions;
    use zip::ZipWriter;

    #[test]
    fn streaming_patch_with_styles_ignores_external_workbook_styles_relationship() {
        let workbook_xml = r#"<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<workbook xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main"
 xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">
  <sheets>
    <sheet name="Sheet1" sheetId="1" r:id="rId1"/>
  </sheets>
</workbook>"#;

        // Internal styles relationship appears before an external one. We should ignore the
        // external relationship so it cannot shadow the real `xl/styles.xml` part.
        let workbook_rels = r#"<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
  <Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/worksheet" Target="worksheets/sheet1.xml"/>
  <Relationship Id="rId2" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/styles" Target="styles.xml"/>
  <Relationship Id="rId3" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/styles" Target="https://example.com/styles.xml" TargetMode="External"/>
</Relationships>"#;

        let worksheet_xml = r#"<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main">
  <dimension ref="A1"/>
  <sheetData/>
</worksheet>"#;

        // Minimal styles.xml. The streaming style-aware patcher should load this part (not the
        // external URI) and append a new `<xf>` when a new style_id is introduced.
        let styles_xml = r#"<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<styleSheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main"/>"#;

        let cursor = Cursor::new(Vec::new());
        let mut zip = ZipWriter::new(cursor);
        let options =
            FileOptions::<()>::default().compression_method(zip::CompressionMethod::Deflated);

        zip.start_file("xl/workbook.xml", options).unwrap();
        zip.write_all(workbook_xml.as_bytes()).unwrap();

        zip.start_file("xl/_rels/workbook.xml.rels", options)
            .unwrap();
        zip.write_all(workbook_rels.as_bytes()).unwrap();

        zip.start_file("xl/worksheets/sheet1.xml", options).unwrap();
        zip.write_all(worksheet_xml.as_bytes()).unwrap();

        zip.start_file("xl/styles.xml", options).unwrap();
        zip.write_all(styles_xml.as_bytes()).unwrap();

        let input_bytes = zip.finish().unwrap().into_inner();

        let mut style_table = StyleTable::default();
        let style_id = style_table.intern(Style {
            number_format: Some("0".to_string()),
            ..Default::default()
        });
        assert_eq!(style_id, 1);

        let mut patches = WorkbookCellPatches::default();
        patches.set_cell(
            "Sheet1",
            CellRef::from_a1("A1").unwrap(),
            CellPatch::set_value_with_style_id(CellValue::Number(1.0), style_id),
        );

        let mut output = Cursor::new(Vec::new());
        patch_xlsx_streaming_workbook_cell_patches_with_styles(
            Cursor::new(input_bytes),
            &mut output,
            &patches,
            &style_table,
        )
        .expect("streaming patch should succeed");
    }

    #[test]
    fn streaming_patch_ignores_external_workbook_shared_strings_relationship() {
        let workbook_xml = r#"<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<workbook xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main"
 xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">
  <sheets>
    <sheet name="Sheet1" sheetId="1" r:id="rId1"/>
  </sheets>
</workbook>"#;

        // External sharedStrings relationship is listed first and should be ignored.
        let workbook_rels = r#"<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
  <Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/worksheet" Target="worksheets/sheet1.xml"/>
  <Relationship Id="rId2" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/sharedStrings" Target="https://example.com/sharedStrings.xml" TargetMode="External"/>
  <Relationship Id="rId3" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/sharedStrings" Target="sharedStrings.xml"/>
</Relationships>"#;

        let worksheet_xml = r#"<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main">
  <dimension ref="A1"/>
  <sheetData/>
</worksheet>"#;

        let shared_strings_xml = r#"<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<sst xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" count="0" uniqueCount="0"/>"#;

        let cursor = Cursor::new(Vec::new());
        let mut zip = ZipWriter::new(cursor);
        let options =
            FileOptions::<()>::default().compression_method(zip::CompressionMethod::Deflated);

        zip.start_file("xl/workbook.xml", options).unwrap();
        zip.write_all(workbook_xml.as_bytes()).unwrap();

        zip.start_file("xl/_rels/workbook.xml.rels", options)
            .unwrap();
        zip.write_all(workbook_rels.as_bytes()).unwrap();

        zip.start_file("xl/worksheets/sheet1.xml", options).unwrap();
        zip.write_all(worksheet_xml.as_bytes()).unwrap();

        zip.start_file("xl/sharedStrings.xml", options).unwrap();
        zip.write_all(shared_strings_xml.as_bytes()).unwrap();

        let input_bytes = zip.finish().unwrap().into_inner();

        let mut patches = WorkbookCellPatches::default();
        patches.set_cell(
            "Sheet1",
            CellRef::from_a1("A1").unwrap(),
            CellPatch::set_value(CellValue::String("Hello".to_string())),
        );

        let mut output = Cursor::new(Vec::new());
        patch_xlsx_streaming_workbook_cell_patches(Cursor::new(input_bytes), &mut output, &patches)
            .expect("streaming patch should succeed");
    }
}
