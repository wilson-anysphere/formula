mod workbook_state;

pub use workbook_state::{open_memory_manager, open_storage};
pub use workbook_state::{PersistentWorkbookState, WorkbookPersistenceLocation};

use crate::atomic_write::write_file_atomic;
use crate::file_io::{
    is_xlsx_family_extension, DefinedName as AppDefinedName, Sheet as AppSheet, Table as AppTable,
    Workbook as AppWorkbook,
};
use crate::sheet_name::sheet_name_eq_case_insensitive;
use crate::state::{Cell, CellScalar};
use anyhow::Context;
use directories::ProjectDirs;
use formula_model::{
    display_formula_text, normalize_formula_text, Cell as ModelCell, CellRef,
    CellValue as ModelCellValue, DefinedNameScope, Style, Workbook as ModelWorkbook,
};
use sha2::{Digest, Sha256};
use std::collections::{HashMap, HashSet};
use std::io::{Cursor, Write};
use std::path::Path;
use std::path::PathBuf;
use std::sync::Arc;
use uuid::Uuid;
use zip::write::FileOptions;
use zip::{CompressionMethod, ZipWriter};

pub fn autosave_db_path_for_workbook(path: &str) -> Option<PathBuf> {
    let proj = ProjectDirs::from("com", "formula", "Formula")?;
    let autosave_dir = proj.data_local_dir().join("autosave");

    const PREFIX: &[u8] = b"formula-autosave-v1\0";
    let mut hasher = Sha256::new();
    hasher.update(PREFIX);
    hasher.update(path.as_bytes());
    let digest = hex::encode(hasher.finalize());
    Some(autosave_dir.join(format!("{digest}.sqlite")))
}

pub fn autosave_db_path_for_new_workbook() -> Option<PathBuf> {
    let proj = ProjectDirs::from("com", "formula", "Formula")?;
    let autosave_dir = proj.data_local_dir().join("autosave");
    Some(autosave_dir.join(format!("unsaved-{}.sqlite", Uuid::new_v4())))
}

pub fn workbook_to_model(workbook: &AppWorkbook) -> anyhow::Result<ModelWorkbook> {
    let mut model = ModelWorkbook::new();
    model.schema_version = formula_model::SCHEMA_VERSION;
    model.id = 0;
    model.date_system = workbook.date_system;

    let mut number_format_style_ids: HashMap<String, u32> = HashMap::new();

    for sheet in &workbook.sheets {
        let sheet_id = model.add_sheet(sheet.name.clone())?;
        if let Some(model_sheet) = model.sheet_mut(sheet_id) {
            model_sheet.visibility = sheet.visibility;
            model_sheet.tab_color = sheet.tab_color.clone();
            model_sheet.default_col_width = sheet.default_col_width;
            model_sheet.col_properties = sheet.col_properties.clone();
        }
        let sheet_idx = model.sheets.len().saturating_sub(1);

        for ((row, col), cell) in sheet.cells_iter() {
            let cell_ref = CellRef::new(row as u32, col as u32);

            let mut out = match (&cell.formula, &cell.input_value) {
                (Some(formula), _) => {
                    let mut c = ModelCell::new(scalar_to_model_value(&cell.computed_value));
                    c.formula = normalize_formula_text(formula);
                    c
                }
                (None, Some(value)) => ModelCell::new(scalar_to_model_value(value)),
                (None, None) => ModelCell::new(ModelCellValue::Empty),
            };

            if let Some(fmt) = cell
                .number_format
                .as_deref()
                .map(str::trim)
                .filter(|s| !s.is_empty())
            {
                if let Some(existing) = number_format_style_ids.get(fmt) {
                    out.style_id = *existing;
                } else {
                    let fmt = fmt.to_string();
                    let style_id = model.styles.intern(Style {
                        number_format: Some(fmt.clone()),
                        ..Default::default()
                    });
                    number_format_style_ids.insert(fmt, style_id);
                    out.style_id = style_id;
                }
            }

            if out.is_truly_empty() {
                continue;
            }

            if let Some(model_sheet) = model.sheets.get_mut(sheet_idx) {
                model_sheet.set_cell(cell_ref, out);
            }
        }
    }

    Ok(model)
}

pub fn workbook_from_model(model: &ModelWorkbook) -> anyhow::Result<AppWorkbook> {
    let mut workbook = AppWorkbook::new_empty(None);
    workbook.date_system = model.date_system;

    workbook.sheets = model
        .sheets
        .iter()
        .map(|sheet| sheet_from_model(sheet, &model.styles))
        .collect::<anyhow::Result<Vec<_>>>()?;

    let sheet_names_by_id: std::collections::HashMap<formula_model::WorksheetId, String> = model
        .sheets
        .iter()
        .map(|sheet| (sheet.id, sheet.name.clone()))
        .collect();

    workbook.defined_names = model
        .defined_names
        .iter()
        .map(|dn| {
            let sheet_id = match dn.scope {
                DefinedNameScope::Workbook => None,
                DefinedNameScope::Sheet(id) => sheet_names_by_id.get(&id).cloned(),
            };

            AppDefinedName {
                name: dn.name.clone(),
                refers_to: dn.refers_to.clone(),
                sheet_id,
                hidden: dn.hidden,
            }
        })
        .collect();

    workbook.tables = model
        .sheets
        .iter()
        .flat_map(|sheet| {
            let sheet_id = sheet.name.clone();
            sheet.tables.iter().map(move |table| AppTable {
                name: table.display_name.clone(),
                sheet_id: sheet_id.clone(),
                start_row: table.range.start.row as usize,
                start_col: table.range.start.col as usize,
                end_row: table.range.end.row as usize,
                end_col: table.range.end.col as usize,
                columns: table.columns.iter().map(|c| c.name.clone()).collect(),
            })
        })
        .collect();

    workbook.ensure_sheet_ids();
    for sheet in &mut workbook.sheets {
        sheet.clear_dirty_cells();
    }

    Ok(workbook)
}

fn sheet_from_model(
    sheet: &formula_model::Worksheet,
    style_table: &formula_model::StyleTable,
) -> anyhow::Result<AppSheet> {
    let mut out = AppSheet::new(sheet.name.clone(), sheet.name.clone());
    out.visibility = sheet.visibility;
    out.tab_color = sheet.tab_color.clone();
    out.default_col_width = sheet.default_col_width;
    out.col_properties = sheet.col_properties.clone();

    for (cell_ref, cell) in sheet.iter_cells() {
        let row = cell_ref.row as usize;
        let col = cell_ref.col as usize;
        let number_format = (cell.style_id != 0)
            .then(|| {
                style_table
                    .get(cell.style_id)
                    .and_then(|s| s.number_format.clone())
            })
            .flatten();

        let cached_value = model_value_to_scalar(&cell.value);
        if let Some(formula) = cell.formula.as_deref() {
            let normalized = display_formula_text(formula);
            if !normalized.trim().is_empty() {
                let mut c = Cell::from_formula(normalized);
                c.computed_value = cached_value;
                c.number_format = number_format;
                out.set_cell(row, col, c);
                continue;
            }
            // Treat empty formulas as blank/no-formula cells.
        }

        if matches!(cached_value, CellScalar::Empty) {
            if let Some(number_format) = number_format {
                let mut c = Cell::empty();
                c.number_format = Some(number_format);
                out.set_cell(row, col, c);
            }
            continue;
        }

        let mut c = Cell::from_literal(Some(cached_value));
        c.number_format = number_format;
        out.set_cell(row, col, c);
    }

    Ok(out)
}

fn scalar_to_model_value(value: &CellScalar) -> ModelCellValue {
    match value {
        CellScalar::Empty => ModelCellValue::Empty,
        CellScalar::Number(n) => ModelCellValue::Number(*n),
        CellScalar::Text(s) => ModelCellValue::String(s.clone()),
        CellScalar::Bool(b) => ModelCellValue::Boolean(*b),
        CellScalar::Error(e) => ModelCellValue::Error(
            e.parse::<formula_model::ErrorValue>()
                .unwrap_or(formula_model::ErrorValue::Unknown),
        ),
    }
}

fn model_value_to_scalar(value: &ModelCellValue) -> CellScalar {
    match value {
        ModelCellValue::Empty => CellScalar::Empty,
        ModelCellValue::Number(n) => CellScalar::Number(*n),
        ModelCellValue::String(s) => CellScalar::Text(s.clone()),
        ModelCellValue::Boolean(b) => CellScalar::Bool(*b),
        ModelCellValue::Error(e) => CellScalar::Error(e.to_string()),
        ModelCellValue::RichText(rt) => CellScalar::Text(rt.text.clone()),
        ModelCellValue::Entity(entity) => CellScalar::Text(entity.display_value.clone()),
        ModelCellValue::Record(record) => CellScalar::Text(record.to_string()),
        ModelCellValue::Image(image) => CellScalar::Text(
            image
                .alt_text
                .clone()
                .filter(|s| !s.is_empty())
                .unwrap_or_else(|| "[Image]".to_string()),
        ),
        other => match other {
            ModelCellValue::Array(arr) => CellScalar::Text(format!("{:?}", arr.data)),
            ModelCellValue::Spill(_) => CellScalar::Error("#SPILL!".to_string()),
            _ => rich_model_cell_value_to_scalar(other)
                .unwrap_or_else(|| CellScalar::Text(format!("{other:?}"))),
        },
    }
}

fn rich_model_cell_value_to_scalar(value: &ModelCellValue) -> Option<CellScalar> {
    fn json_get_str<'a>(value: &'a serde_json::Value, keys: &[&str]) -> Option<&'a str> {
        for key in keys {
            if let Some(s) = value.get(key).and_then(|v| v.as_str()) {
                return Some(s);
            }
        }
        None
    }

    fn cell_value_json_to_display_string(value: &serde_json::Value) -> Option<String> {
        let value_type = value.get("type")?.as_str()?;
        match value_type {
            "number" => Some(value.get("value")?.as_f64()?.to_string()),
            "string" => Some(value.get("value")?.as_str()?.to_string()),
            "boolean" => Some(if value.get("value")?.as_bool()? {
                "TRUE".to_string()
            } else {
                "FALSE".to_string()
            }),
            "error" => Some(value.get("value")?.as_str()?.to_string()),
            "rich_text" => Some(value.get("value")?.get("text")?.as_str()?.to_string()),
            _ => None,
        }
    }

    let serialized = serde_json::to_value(value).ok()?;
    let value_type = serialized.get("type")?.as_str()?;

    match value_type {
        "entity" => {
            let entity = serialized.get("value")?;
            let display_value =
                json_get_str(entity, &["displayValue", "display_value", "display"])?.to_string();
            Some(CellScalar::Text(display_value))
        }
        "record" => {
            let record = serialized.get("value")?;
            if let Some(display_field) = json_get_str(record, &["displayField", "display_field"]) {
                if let Some(fields) = record.get("fields").and_then(|v| v.as_object()) {
                    if let Some(display_value) = fields.get(display_field) {
                        if let Some(display) = cell_value_json_to_display_string(display_value) {
                            return Some(CellScalar::Text(display));
                        }
                    }
                }
            }

            let display_value =
                json_get_str(record, &["displayValue", "display_value", "display"])?.to_string();
            Some(CellScalar::Text(display_value))
        }
        "image" => {
            let image = serialized.get("value")?;
            let alt_text = image
                .get("altText")
                .or_else(|| image.get("alt_text"))
                .and_then(|v| v.as_str())
                .unwrap_or("");
            if alt_text.is_empty() {
                Some(CellScalar::Text("[Image]".to_string()))
            } else {
                Some(CellScalar::Text(alt_text.to_string()))
            }
        }
        _ => None,
    }
}

#[cfg(test)]
mod tests {
    use super::*;

    #[test]
    fn record_value_to_scalar_prefers_display_field_over_display_value() {
        let record = formula_model::RecordValue::default()
            .with_display_field("Name")
            .with_field("Name", "Alice");

        let scalar = model_value_to_scalar(&ModelCellValue::Record(record));
        assert_eq!(scalar, CellScalar::Text("Alice".to_string()));
    }
}

/// Export a workbook from SQLite and build an `.xlsx`/`.xlsm` package as an in-memory ZIP buffer.
///
/// We currently write a fresh XLSX ZIP from the `formula-model` export and then
/// re-apply a handful of preserved parts (VBA, drawing parts, pivot attachments)
/// plus workbook print settings.
///
/// This keeps SQLite as the source of truth for the current workbook state and
/// supports autosave/crash recovery. We intentionally do **not** use the older
/// patch-based XLSX save path here yet because emitting `WorkbookCellPatches`
/// directly from SQLite deltas is not implemented.
pub fn build_xlsx_from_storage(
    storage: &formula_storage::Storage,
    workbook_id: Uuid,
    workbook_meta: &AppWorkbook,
    path: &Path,
) -> anyhow::Result<Arc<[u8]>> {
    let xlsx_date_system = match workbook_meta.date_system {
        formula_model::DateSystem::Excel1900 => formula_xlsx::DateSystem::V1900,
        formula_model::DateSystem::Excel1904 => formula_xlsx::DateSystem::V1904,
    };

    let mut model = storage
        .export_model_workbook(workbook_id)
        .context("export workbook from storage")?;

    // Patch cached formula values using the in-memory engine results so exports
    // don't ship stale/empty `<v>` values for formula cells.
    //
    // This is particularly important for:
    // - new workbooks (no original XLSX baseline)
    // - non-XLSX imports (csv/xls/xlsb) where we generate a fresh XLSX on save
    apply_cached_formula_values(&mut model, workbook_meta);

    let mut cursor = Cursor::new(Vec::new());
    formula_xlsx::write_workbook_to_writer(&model, &mut cursor)
        .context("write workbook to bytes")?;
    let mut bytes = cursor.into_inner();

    let extension = path
        .extension()
        .and_then(|s| s.to_str())
        .map(|s| s.to_ascii_lowercase());
    let workbook_kind = extension
        .as_deref()
        .and_then(formula_xlsx::WorkbookKind::from_extension)
        .unwrap_or(formula_xlsx::WorkbookKind::Workbook);

    let wants_vba = workbook_meta.vba_project_bin.is_some() && workbook_kind.is_macro_enabled();
    let wants_vba_signature = wants_vba && workbook_meta.vba_project_signature_bin.is_some();
    let wants_preserved_drawings = workbook_meta.preserved_drawing_parts.is_some();
    let wants_preserved_pivots = workbook_meta.preserved_pivot_parts.is_some();
    let wants_power_query = workbook_meta.power_query_xml.is_some();
    let wants_macro_strip =
        workbook_kind.is_macro_free() && workbook_meta.vba_project_bin.is_some();
    let wants_content_type_enforcement = workbook_kind != formula_xlsx::WorkbookKind::Workbook;
    let needs_date_system_update = extension
        .as_deref()
        .is_some_and(|ext| is_xlsx_family_extension(ext))
        && matches!(
            workbook_meta.date_system,
            formula_model::DateSystem::Excel1904
        );

    // Repack in a streaming-friendly way to avoid `XlsxPackage::from_bytes` inflating every ZIP
    // entry (which can be prohibitively memory-intensive for large exports).
    //
    // We generate the workbook from the model (streaming ZIP writer) and then optionally apply
    // preserved parts / VBA payloads / content type tweaks via a streaming ZIP rewrite.
    let needs_repack_overrides = wants_vba
        || wants_preserved_drawings
        || wants_preserved_pivots
        || wants_power_query
        || wants_content_type_enforcement
        || needs_date_system_update;

    if needs_repack_overrides {
        let part_overrides = build_export_part_overrides_from_subset_package(
            &bytes,
            workbook_meta,
            workbook_kind,
            needs_date_system_update,
            xlsx_date_system,
            wants_vba,
            wants_vba_signature,
        )?;

        if !part_overrides.is_empty() {
            let mut cursor = Cursor::new(Vec::new());
            formula_xlsx::patch_xlsx_streaming_workbook_cell_patches_with_part_overrides_and_recalc_policy(
                Cursor::new(bytes),
                &mut cursor,
                &formula_xlsx::WorkbookCellPatches::default(),
                &part_overrides,
                formula_xlsx::RecalcPolicy::default(),
            )
            .context("apply export part overrides (streaming)")?;
            bytes = cursor.into_inner();
        }
    }

    if wants_macro_strip {
        let mut cursor = Cursor::new(Vec::new());
        formula_xlsx::strip_vba_project_streaming_with_kind(
            Cursor::new(bytes),
            &mut cursor,
            workbook_kind,
        )
        .context("strip macros for macro-free export (streaming)")?;
        bytes = cursor.into_inner();
    }

    if extension
        .as_deref()
        .is_some_and(|ext| is_xlsx_family_extension(ext))
    {
        bytes = formula_xlsx::print::write_workbook_print_settings(
            &bytes,
            &workbook_meta.print_settings,
        )
        .context("write workbook print settings")?;
    }

    Ok(Arc::<[u8]>::from(bytes))
}

pub fn write_xlsx_from_storage(
    storage: &formula_storage::Storage,
    workbook_id: Uuid,
    workbook_meta: &AppWorkbook,
    path: &Path,
) -> anyhow::Result<Arc<[u8]>> {
    let bytes = build_xlsx_from_storage(storage, workbook_id, workbook_meta, path)?;
    write_file_atomic(path, bytes.as_ref()).with_context(|| format!("write workbook {path:?}"))?;
    Ok(bytes)
}

fn build_export_part_overrides_from_subset_package(
    base_bytes: &[u8],
    workbook_meta: &AppWorkbook,
    workbook_kind: formula_xlsx::WorkbookKind,
    needs_date_system_update: bool,
    xlsx_date_system: formula_xlsx::DateSystem,
    wants_vba: bool,
    wants_vba_signature: bool,
) -> anyhow::Result<HashMap<String, formula_xlsx::PartOverride>> {
    fn zip_part_names(bytes: &[u8]) -> anyhow::Result<HashSet<String>> {
        let mut cursor = Cursor::new(bytes);
        let mut archive = zip::ZipArchive::new(&mut cursor).context("open xlsx zip archive")?;
        let mut names = HashSet::new();
        for i in 0..archive.len() {
            let file = archive.by_index(i).context("read zip entry")?;
            if file.is_dir() {
                continue;
            }
            let name = file.name();
            let canonical = name.strip_prefix('/').unwrap_or(name);
            names.insert(canonical.to_string());
        }
        Ok(names)
    }

    fn read_required_part(bytes: &[u8], name: &str) -> anyhow::Result<Vec<u8>> {
        formula_xlsx::read_part_from_reader(Cursor::new(bytes), name)
            .with_context(|| format!("read {name} from generated workbook"))?
            .with_context(|| format!("missing required {name} part"))
    }

    let base_part_names = zip_part_names(base_bytes).context("list base workbook part names")?;

    // Build a "subset" XLSX package containing only the parts we need to mutate. This keeps memory
    // proportional to the number of touched parts (and avoids inflating all worksheets).
    //
    // Note: build the subset ZIP by reading/writing each part sequentially rather than buffering
    // all part bytes at once. This reduces peak memory when multiple large sheet parts are needed
    // (e.g. several sheets with preserved drawings).
    let mut needed_sheet_parts: HashSet<String> = HashSet::new();
    let wants_preserved_drawings = workbook_meta.preserved_drawing_parts.is_some();
    let wants_preserved_pivots = workbook_meta.preserved_pivot_parts.is_some();
    if wants_preserved_drawings || wants_preserved_pivots {
        let worksheet_parts = formula_xlsx::worksheet_parts_from_reader(Cursor::new(base_bytes))
            .context("resolve worksheet parts for preserved part application")?;

        let resolve_sheet = |preserved_name: &str, preserved_index: usize| {
            worksheet_parts
                .iter()
                .find(|p| sheet_name_eq_case_insensitive(&p.name, preserved_name))
                .or_else(|| worksheet_parts.get(preserved_index))
        };

        if let Some(preserved) = workbook_meta.preserved_drawing_parts.as_ref() {
            for (sheet_name, entry) in &preserved.sheet_drawings {
                if entry.drawings.is_empty() {
                    continue;
                }
                if let Some(info) = resolve_sheet(sheet_name, entry.sheet_index) {
                    needed_sheet_parts.insert(info.worksheet_part.clone());
                }
            }
            for (sheet_name, entry) in &preserved.sheet_pictures {
                if let Some(info) = resolve_sheet(sheet_name, entry.sheet_index) {
                    needed_sheet_parts.insert(info.worksheet_part.clone());
                }
            }
            for (sheet_name, entry) in &preserved.sheet_ole_objects {
                if let Some(info) = resolve_sheet(sheet_name, entry.sheet_index) {
                    needed_sheet_parts.insert(info.worksheet_part.clone());
                }
            }
            for (sheet_name, entry) in &preserved.sheet_controls {
                if let Some(info) = resolve_sheet(sheet_name, entry.sheet_index) {
                    needed_sheet_parts.insert(info.worksheet_part.clone());
                }
            }
            for (sheet_name, entry) in &preserved.sheet_drawing_hfs {
                if let Some(info) = resolve_sheet(sheet_name, entry.sheet_index) {
                    needed_sheet_parts.insert(info.worksheet_part.clone());
                }
            }
        }

        if let Some(preserved) = workbook_meta.preserved_pivot_parts.as_ref() {
            for (sheet_name, entry) in &preserved.sheet_pivot_tables {
                if let Some(info) = resolve_sheet(sheet_name, entry.sheet_index) {
                    needed_sheet_parts.insert(info.worksheet_part.clone());
                }
            }
        }

    }

    let cursor = Cursor::new(Vec::new());
    let mut zip = ZipWriter::new(cursor);
    let options = FileOptions::<()>::default().compression_method(CompressionMethod::Deflated);

    let mut write_part = |name: &str, bytes: &[u8]| -> anyhow::Result<()> {
        zip.start_file(name, options)
            .with_context(|| format!("start subset zip entry {name}"))?;
        zip.write_all(bytes)
            .with_context(|| format!("write subset zip entry {name}"))?;
        Ok(())
    };

    write_part(
        "[Content_Types].xml",
        &read_required_part(base_bytes, "[Content_Types].xml")?,
    )?;
    write_part(
        "xl/workbook.xml",
        &read_required_part(base_bytes, "xl/workbook.xml")?,
    )?;
    write_part(
        "xl/_rels/workbook.xml.rels",
        &read_required_part(base_bytes, "xl/_rels/workbook.xml.rels")?,
    )?;

    if !needed_sheet_parts.is_empty() {
        let mut needed_sheet_parts: Vec<String> = needed_sheet_parts.into_iter().collect();
        needed_sheet_parts.sort();

        for worksheet_part in needed_sheet_parts {
            let xml = read_required_part(base_bytes, &worksheet_part)?;
            write_part(&worksheet_part, &xml)?;

            let rels_part = formula_xlsx::openxml::rels_part_name(&worksheet_part);
            let rels_xml = read_required_part(base_bytes, &rels_part)?;
            write_part(&rels_part, &rels_xml)?;
        }
    }

    let subset_bytes = zip.finish().context("finalize subset xlsx zip")?.into_inner();

    let mut pkg =
        formula_xlsx::XlsxPackage::from_bytes(&subset_bytes).context("parse subset xlsx package")?;

    if wants_vba {
        pkg.set_part(
            "xl/vbaProject.bin",
            workbook_meta
                .vba_project_bin
                .clone()
                .expect("checked is_some"),
        );
    }
    if wants_vba_signature {
        pkg.set_part(
            "xl/vbaProjectSignature.bin",
            workbook_meta
                .vba_project_signature_bin
                .clone()
                .expect("checked is_some"),
        );
    }

    if let Some(preserved) = workbook_meta.preserved_drawing_parts.as_ref() {
        pkg.apply_preserved_drawing_parts(preserved)
            .context("apply preserved drawing parts")?;
    }

    if let Some(preserved) = workbook_meta.preserved_pivot_parts.as_ref() {
        pkg.apply_preserved_pivot_parts(preserved)
            .context("apply preserved pivot parts")?;
    }

    match workbook_meta.power_query_xml.as_ref() {
        Some(bytes) => pkg.set_part("xl/formula/power-query.xml", bytes.clone()),
        None => {
            pkg.parts_map_mut().remove("xl/formula/power-query.xml");
        }
    }

    // Enforce workbook kind by patching the workbook override in `[Content_Types].xml`.
    let content_types = pkg
        .part("[Content_Types].xml")
        .ok_or_else(|| anyhow::anyhow!("subset package is missing [Content_Types].xml"))?;
    if let Some(updated) = formula_xlsx::rewrite_content_types_workbook_kind(content_types, workbook_kind)
        .context("rewrite workbook kind in [Content_Types].xml")?
    {
        pkg.set_part("[Content_Types].xml", updated);
    }

    if needs_date_system_update {
        pkg.set_workbook_date_system(xlsx_date_system)
            .context("set workbook date system")?;
    }

    let repacked_subset = pkg
        .write_to_bytes()
        .context("repack subset xlsx package")?;
    let repacked_pkg = formula_xlsx::XlsxPackage::from_bytes(&repacked_subset)
        .context("parse repacked subset package")?;

    let mut part_overrides: HashMap<String, formula_xlsx::PartOverride> = HashMap::new();
    for (name, bytes) in repacked_pkg.parts() {
        let canonical = name.strip_prefix('/').unwrap_or(name).to_string();
        let override_op = if base_part_names.contains(&canonical) {
            formula_xlsx::PartOverride::Replace(bytes.to_vec())
        } else {
            formula_xlsx::PartOverride::Add(bytes.to_vec())
        };
        part_overrides.insert(canonical, override_op);
    }

    // Ensure we handle removals deterministically (and match the previous behavior) even if the
    // part isn't present in the generated workbook.
    if workbook_meta.power_query_xml.is_none() {
        part_overrides.insert(
            "xl/formula/power-query.xml".to_string(),
            formula_xlsx::PartOverride::Remove,
        );
    }

    Ok(part_overrides)
}

fn apply_cached_formula_values(model: &mut ModelWorkbook, workbook: &AppWorkbook) {
    for sheet in &workbook.sheets {
        let Some(model_sheet) = model
            .sheets
            .iter_mut()
            .find(|s| sheet_name_eq_case_insensitive(&s.name, &sheet.name))
        else {
            continue;
        };

        for ((row, col), cell) in sheet.cells_iter() {
            if cell.formula.is_none() {
                continue;
            }
            let (row, col) = match (u32::try_from(row), u32::try_from(col)) {
                (Ok(r), Ok(c)) => (r, c),
                _ => continue,
            };
            let cell_ref = CellRef::new(row, col);

            // Only update cached values for cells that are formulas in the model workbook.
            if model_sheet.formula(cell_ref).is_none() {
                continue;
            }

            let computed = cell.computed_value.clone();
            let existing = model_sheet
                .cell(cell_ref)
                .map(|c| model_value_to_scalar(&c.value))
                .unwrap_or(CellScalar::Empty);

            // Preserve existing cached values when the engine can't evaluate the formula (commonly
            // surfaced as `#NAME?`). This keeps round-trips stable for formulas we don't support yet.
            if matches!(computed, CellScalar::Error(_)) && !matches!(existing, CellScalar::Error(_))
            {
                continue;
            }

            if computed != existing {
                model_sheet.set_value(cell_ref, scalar_to_model_value(&computed));
            }
        }
    }
}

#[cfg(test)]
mod write_xlsx_from_storage_tests {
    use super::*;
    use anyhow::Context;
    use formula_storage::ImportModelWorkbookOptions;
    use formula_xlsx::print::{CellRange, ColRange, Orientation, PrintTitles, RowRange, Scaling};
    use std::io::Cursor;
    use std::path::Path;

    fn import_app_workbook(
        storage: &formula_storage::Storage,
        workbook: &AppWorkbook,
    ) -> anyhow::Result<Uuid> {
        let model = workbook_to_model(workbook).context("convert workbook to model")?;
        let meta = storage
            .import_model_workbook(&model, ImportModelWorkbookOptions::new("test"))
            .context("import workbook into storage")?;
        Ok(meta.id)
    }

    #[test]
    fn write_xlsx_from_storage_creates_parent_dirs_and_overwrites_existing_file(
    ) -> anyhow::Result<()> {
        let storage =
            formula_storage::Storage::open_in_memory().context("open in-memory storage")?;

        let mut workbook_meta = AppWorkbook::new_empty(None);
        workbook_meta.add_sheet("Sheet1".to_string());
        workbook_meta.ensure_sheet_ids();
        workbook_meta.sheets[0].set_cell(0, 0, Cell::from_literal(Some(CellScalar::Number(123.0))));

        let workbook_id = import_app_workbook(&storage, &workbook_meta)?;

        let tmp = tempfile::tempdir().context("temp dir")?;
        let out_path = tmp.path().join("nested/dir/export.xlsx");

        // Parent directories should be created automatically.
        let first_bytes = write_xlsx_from_storage(&storage, workbook_id, &workbook_meta, &out_path)
            .context("first export")?;
        assert!(out_path.exists(), "expected output file to exist");
        assert!(
            out_path.parent().expect("path should have parent").is_dir(),
            "expected parent directories to be created"
        );
        assert_eq!(
            std::fs::read(&out_path)
                .context("read first output")?
                .as_slice(),
            first_bytes.as_ref(),
            "expected file bytes to match returned bytes"
        );

        // Pre-create the file with sentinel content to ensure overwrite/replacement semantics.
        std::fs::write(&out_path, b"old").context("write sentinel bytes")?;
        assert_eq!(std::fs::read(&out_path)?.as_slice(), b"old");

        let second_bytes =
            write_xlsx_from_storage(&storage, workbook_id, &workbook_meta, &out_path)
                .context("second export")?;
        let disk_bytes = std::fs::read(&out_path).context("read overwritten output")?;
        assert_ne!(
            disk_bytes.as_slice(),
            b"old",
            "expected export to overwrite sentinel bytes"
        );
        assert_eq!(
            disk_bytes.as_slice(),
            second_bytes.as_ref(),
            "expected on-disk bytes to match returned bytes"
        );

        Ok(())
    }

    fn assert_content_type_contains_workbook_main(
        bytes: &[u8],
        expected: &str,
    ) -> anyhow::Result<()> {
        let pkg = formula_xlsx::XlsxPackage::from_bytes(bytes).context("parse xlsx package")?;
        let ct_xml = pkg
            .part("[Content_Types].xml")
            .context("missing [Content_Types].xml")?;
        let ct_xml = std::str::from_utf8(ct_xml).context("content types is not valid utf8")?;
        assert!(
            ct_xml.contains(expected),
            "expected [Content_Types].xml to contain workbook main content type {expected}, got:\n{ct_xml}"
        );
        Ok(())
    }

    #[test]
    fn write_xlsx_from_storage_enforces_workbook_kind_and_macro_behavior_for_xlsx_family(
    ) -> anyhow::Result<()> {
        let fixture_path =
            Path::new(env!("CARGO_MANIFEST_DIR")).join("../../../fixtures/xlsx/macros/basic.xlsm");
        let mut workbook_meta =
            crate::file_io::read_xlsx_blocking(&fixture_path).context("read macro fixture")?;
        let signature_bytes = b"fake-vba-project-signature-for-storage-export-test".to_vec();
        workbook_meta.vba_project_signature_bin = Some(signature_bytes.clone());

        let expected_vba = workbook_meta
            .vba_project_bin
            .as_deref()
            .context("expected macro fixture to contain xl/vbaProject.bin")?;
        let expected_signature = signature_bytes.as_slice();

        let storage =
            formula_storage::Storage::open_in_memory().context("open in-memory storage")?;
        let workbook_id = import_app_workbook(&storage, &workbook_meta)?;

        let tmp = tempfile::tempdir().context("temp dir")?;

        let cases = [
            (
                "xlsx",
                formula_xlsx::WorkbookKind::Workbook,
                None::<&[u8]>,
                None::<&[u8]>,
            ),
            (
                "xlsm",
                formula_xlsx::WorkbookKind::MacroEnabledWorkbook,
                Some(expected_vba),
                Some(expected_signature),
            ),
            (
                "xltx",
                formula_xlsx::WorkbookKind::Template,
                None::<&[u8]>,
                None::<&[u8]>,
            ),
            (
                "xltm",
                formula_xlsx::WorkbookKind::MacroEnabledTemplate,
                Some(expected_vba),
                Some(expected_signature),
            ),
            (
                "xlam",
                formula_xlsx::WorkbookKind::MacroEnabledAddIn,
                Some(expected_vba),
                Some(expected_signature),
            ),
        ];

        for (ext, kind, expected_vba_part, expected_signature_part) in cases {
            let out_path = tmp.path().join(format!("export.{ext}"));
            let bytes = write_xlsx_from_storage(&storage, workbook_id, &workbook_meta, &out_path)
                .with_context(|| format!("export to .{ext}"))?;

            assert_content_type_contains_workbook_main(
                bytes.as_ref(),
                kind.workbook_content_type(),
            )
            .with_context(|| format!("check workbook main content type for .{ext}"))?;

            let pkg = formula_xlsx::XlsxPackage::from_bytes(bytes.as_ref())
                .with_context(|| format!("parse exported .{ext} package"))?;
            formula_xlsx::validate_opc_relationships(pkg.parts_map())
                .with_context(|| format!("validate OPC relationships for .{ext} export"))?;
            assert_eq!(
                pkg.vba_project_bin(),
                expected_vba_part,
                "unexpected VBA project presence for .{ext}"
            );
            assert_eq!(
                pkg.vba_project_signature_bin(),
                expected_signature_part,
                "unexpected VBA project signature presence for .{ext}"
            );

            if expected_vba_part.is_some() {
                // Ensure we emit a structurally valid macro-enabled package:
                // - content types contain the VBA overrides
                // - workbook relationships reference the VBA project part
                // - vbaProject.bin relationships reference the signature part (when present)
                let ct_xml = std::str::from_utf8(
                    pkg.part("[Content_Types].xml")
                        .context("missing [Content_Types].xml")?,
                )
                .context("content types is not valid utf8")?;
                assert!(
                    ct_xml.contains("application/vnd.ms-office.vbaProject"),
                    "expected macro-enabled export to contain vbaProject content type override, got:\n{ct_xml}"
                );
                assert!(
                    ct_xml.contains("application/vnd.ms-office.vbaProjectSignature"),
                    "expected macro-enabled export to contain vbaProjectSignature content type override, got:\n{ct_xml}"
                );

                let workbook_rels = std::str::from_utf8(
                    pkg.part("xl/_rels/workbook.xml.rels")
                        .context("missing xl/_rels/workbook.xml.rels")?,
                )
                .context("workbook rels is not valid utf8")?;
                assert!(
                    workbook_rels.contains("vbaProject.bin"),
                    "expected workbook.xml.rels to reference vbaProject.bin, got:\n{workbook_rels}"
                );
                assert!(
                    workbook_rels.contains("vbaProject"),
                    "expected workbook.xml.rels to contain a vbaProject relationship type, got:\n{workbook_rels}"
                );

                let vba_rels = std::str::from_utf8(
                    pkg.part("xl/_rels/vbaProject.bin.rels")
                        .context("missing xl/_rels/vbaProject.bin.rels")?,
                )
                .context("vbaProject.bin rels is not valid utf8")?;
                assert!(
                    vba_rels.contains("vbaProjectSignature.bin"),
                    "expected vbaProject.bin.rels to reference vbaProjectSignature.bin, got:\n{vba_rels}"
                );
                assert!(
                    vba_rels.contains("vbaProjectSignature"),
                    "expected vbaProject.bin.rels to contain a vbaProjectSignature relationship type, got:\n{vba_rels}"
                );
            }
        }

        Ok(())
    }

    #[test]
    fn write_xlsx_from_storage_writes_print_settings_for_xlsx_family() -> anyhow::Result<()> {
        let fixture_path = Path::new(env!("CARGO_MANIFEST_DIR"))
            .join("../../../fixtures/xlsx/basic/print-settings.xlsx");
        let workbook_meta = crate::file_io::read_xlsx_blocking(&fixture_path)
            .context("read print settings fixture")?;

        assert_eq!(
            workbook_meta.print_settings.sheets.len(),
            1,
            "expected fixture to contain one sheet worth of print settings"
        );
        let sheet = &workbook_meta.print_settings.sheets[0];
        assert_eq!(sheet.sheet_name, "Sheet1");
        assert_eq!(
            sheet.print_area.as_deref(),
            Some(
                &[CellRange {
                    start_row: 1,
                    end_row: 10,
                    start_col: 1,
                    end_col: 4,
                }][..]
            )
        );
        assert_eq!(
            sheet.print_titles,
            Some(PrintTitles {
                repeat_rows: Some(RowRange { start: 1, end: 1 }),
                repeat_cols: Some(ColRange { start: 1, end: 2 }),
            })
        );
        assert_eq!(sheet.page_setup.orientation, Orientation::Landscape);
        assert_eq!(sheet.page_setup.paper_size.code, 9);
        assert_eq!(
            sheet.page_setup.scaling,
            Scaling::FitTo {
                width: 1,
                height: 0
            }
        );
        assert!(sheet.manual_page_breaks.row_breaks_after.contains(&5));
        assert!(sheet.manual_page_breaks.col_breaks_after.contains(&2));

        let storage =
            formula_storage::Storage::open_in_memory().context("open in-memory storage")?;
        let workbook_id = import_app_workbook(&storage, &workbook_meta)?;

        let tmp = tempfile::tempdir().context("temp dir")?;
        for ext in ["xlsx", "xlsm", "xltx", "xltm", "xlam"] {
            let out_path = tmp.path().join(format!("print-settings.{ext}"));
            let bytes = write_xlsx_from_storage(&storage, workbook_id, &workbook_meta, &out_path)
                .with_context(|| format!("export to .{ext}"))?;

            let pkg = formula_xlsx::XlsxPackage::from_bytes(bytes.as_ref())
                .with_context(|| format!("parse exported .{ext} package"))?;
            formula_xlsx::validate_opc_relationships(pkg.parts_map())
                .with_context(|| format!("validate OPC relationships for .{ext} export"))?;

            let reread = formula_xlsx::print::read_workbook_print_settings(bytes.as_ref())
                .context("read print settings from exported workbook")?;
            assert_eq!(
                reread, workbook_meta.print_settings,
                "expected print settings to round-trip for .{ext}"
            );
        }

        Ok(())
    }

    #[test]
    fn write_xlsx_from_storage_reapplies_power_query_part() -> anyhow::Result<()> {
        let storage =
            formula_storage::Storage::open_in_memory().context("open in-memory storage")?;

        let mut workbook_meta = AppWorkbook::new_empty(None);
        workbook_meta.add_sheet("Sheet1".to_string());
        workbook_meta.ensure_sheet_ids();
        workbook_meta.sheets[0].set_cell(
            0,
            0,
            Cell::from_literal(Some(CellScalar::Text("hello".to_string()))),
        );
        let power_query_xml =
            b"<formula><powerQuery><![CDATA[{\"query\":\"test\"}]]></powerQuery></formula>"
                .to_vec();
        workbook_meta.power_query_xml = Some(power_query_xml.clone());

        let workbook_id = import_app_workbook(&storage, &workbook_meta)?;

        let tmp = tempfile::tempdir().context("temp dir")?;
        let out_path = tmp.path().join("power-query.xlsx");
        let bytes = write_xlsx_from_storage(&storage, workbook_id, &workbook_meta, &out_path)
            .context("export workbook with power query")?;

        let pkg = formula_xlsx::XlsxPackage::from_bytes(bytes.as_ref())
            .context("parse exported package")?;
        formula_xlsx::validate_opc_relationships(pkg.parts_map())
            .context("validate OPC relationships for exported workbook")?;
        let out_power_query = pkg
            .part("xl/formula/power-query.xml")
            .context("expected xl/formula/power-query.xml to be present")?;
        assert_eq!(
            out_power_query,
            power_query_xml.as_slice(),
            "expected power query part bytes to round-trip"
        );

        Ok(())
    }

    #[test]
    fn write_xlsx_from_storage_reapplies_preserved_drawing_parts() -> anyhow::Result<()> {
        let fixture_path =
            Path::new(env!("CARGO_MANIFEST_DIR")).join("../../../fixtures/xlsx/basic/image.xlsx");
        let workbook_meta =
            crate::file_io::read_xlsx_blocking(&fixture_path).context("read image fixture")?;
        let preserved = workbook_meta
            .preserved_drawing_parts
            .as_ref()
            .context("expected image fixture to have preserved drawing parts")?;

        let expected_image = preserved
            .parts
            .get("xl/media/image1.png")
            .context("expected preserved drawing parts to contain xl/media/image1.png")?;
        let expected_drawing = preserved
            .parts
            .get("xl/drawings/drawing1.xml")
            .context("expected preserved drawing parts to contain xl/drawings/drawing1.xml")?;

        let storage =
            formula_storage::Storage::open_in_memory().context("open in-memory storage")?;
        let workbook_id = import_app_workbook(&storage, &workbook_meta)?;

        let tmp = tempfile::tempdir().context("temp dir")?;
        let out_path = tmp.path().join("image.xlsx");
        let bytes = write_xlsx_from_storage(&storage, workbook_id, &workbook_meta, &out_path)
            .context("export workbook with preserved drawings")?;

        let pkg = formula_xlsx::XlsxPackage::from_bytes(bytes.as_ref())
            .context("parse exported package")?;
        formula_xlsx::validate_opc_relationships(pkg.parts_map())
            .context("validate OPC relationships for exported workbook")?;
        let out_image = pkg
            .part("xl/media/image1.png")
            .context("expected xl/media/image1.png to be present in output")?;
        assert_eq!(
            out_image,
            expected_image.as_slice(),
            "expected image part to be copied byte-for-byte"
        );

        let out_drawing = pkg
            .part("xl/drawings/drawing1.xml")
            .context("expected xl/drawings/drawing1.xml to be present in output")?;
        assert_eq!(
            out_drawing,
            expected_drawing.as_slice(),
            "expected drawing part to be copied byte-for-byte"
        );

        let sheet_xml = pkg
            .part("xl/worksheets/sheet1.xml")
            .context("expected worksheet xml to be present")?;
        let sheet_xml_str =
            std::str::from_utf8(sheet_xml).context("worksheet xml is not valid utf8")?;
        assert!(
            sheet_xml_str.contains("<drawing"),
            "expected output worksheet xml to contain a <drawing> tag after preserved drawing application, got:\n{sheet_xml_str}"
        );

        Ok(())
    }

    #[test]
    fn write_xlsx_from_storage_reapplies_preserved_pivot_parts() -> anyhow::Result<()> {
        let fixture_path = Path::new(env!("CARGO_MANIFEST_DIR"))
            .join("../../../fixtures/xlsx/pivots/pivot-fixture.xlsx");
        let workbook_meta =
            crate::file_io::read_xlsx_blocking(&fixture_path).context("read pivot fixture")?;
        let preserved = workbook_meta
            .preserved_pivot_parts
            .as_ref()
            .context("expected pivot fixture to have preserved pivot parts")?;

        let expected_pivot_table = preserved
            .parts
            .get("xl/pivotTables/pivotTable1.xml")
            .context("expected preserved pivot parts to contain pivotTable1.xml")?;
        let expected_cache_def = preserved
            .parts
            .get("xl/pivotCache/pivotCacheDefinition1.xml")
            .context("expected preserved pivot parts to contain pivotCacheDefinition1.xml")?;

        let storage =
            formula_storage::Storage::open_in_memory().context("open in-memory storage")?;
        let workbook_id = import_app_workbook(&storage, &workbook_meta)?;

        let tmp = tempfile::tempdir().context("temp dir")?;
        let out_path = tmp.path().join("pivot.xlsx");
        let bytes = write_xlsx_from_storage(&storage, workbook_id, &workbook_meta, &out_path)
            .context("export workbook with preserved pivots")?;

        let pkg = formula_xlsx::XlsxPackage::from_bytes(bytes.as_ref())
            .context("parse exported package")?;
        formula_xlsx::validate_opc_relationships(pkg.parts_map())
            .context("validate OPC relationships for exported workbook")?;
        let out_pivot_table = pkg
            .part("xl/pivotTables/pivotTable1.xml")
            .context("expected xl/pivotTables/pivotTable1.xml to be present in output")?;
        assert_eq!(
            out_pivot_table,
            expected_pivot_table.as_slice(),
            "expected pivot table part to be copied byte-for-byte"
        );

        let out_cache_def = pkg
            .part("xl/pivotCache/pivotCacheDefinition1.xml")
            .context("expected xl/pivotCache/pivotCacheDefinition1.xml to be present in output")?;
        assert_eq!(
            out_cache_def,
            expected_cache_def.as_slice(),
            "expected pivot cache definition part to be copied byte-for-byte"
        );

        Ok(())
    }

    #[test]
    fn write_xlsx_from_storage_sets_date1904_for_xlsx_family_outputs() -> anyhow::Result<()> {
        let storage =
            formula_storage::Storage::open_in_memory().context("open in-memory storage")?;

        let mut workbook_meta = AppWorkbook::new_empty(None);
        workbook_meta.date_system = formula_model::DateSystem::Excel1904;
        workbook_meta.add_sheet("Sheet1".to_string());
        workbook_meta.ensure_sheet_ids();
        workbook_meta.sheets[0].set_cell(0, 0, Cell::from_literal(Some(CellScalar::Number(1.0))));

        let workbook_id = import_app_workbook(&storage, &workbook_meta)?;
        let tmp = tempfile::tempdir().context("temp dir")?;

        for ext in ["xlsx", "xlsm", "xltx", "xltm", "xlam"] {
            let out_path = tmp.path().join(format!("date1904.{ext}"));
            let bytes = write_xlsx_from_storage(&storage, workbook_id, &workbook_meta, &out_path)
                .with_context(|| format!("export .{ext}"))?;

            let pkg = formula_xlsx::XlsxPackage::from_bytes(bytes.as_ref())
                .with_context(|| format!("parse exported .{ext} package"))?;
            formula_xlsx::validate_opc_relationships(pkg.parts_map())
                .with_context(|| format!("validate OPC relationships for .{ext} export"))?;

            let reread = formula_xlsx::read_workbook_from_reader(Cursor::new(bytes.as_ref()))
                .with_context(|| format!("read exported workbook model for .{ext}"))?;
            assert_eq!(
                reread.date_system,
                formula_model::DateSystem::Excel1904,
                "expected exported workbook to use the 1904 date system for .{ext}"
            );
        }

        Ok(())
    }

    #[test]
    fn write_xlsx_from_storage_preserves_cached_values_when_engine_returns_error(
    ) -> anyhow::Result<()> {
        let storage =
            formula_storage::Storage::open_in_memory().context("open in-memory storage")?;

        // Seed the storage workbook with a formula cell that has a non-error cached value.
        let mut workbook_for_storage = AppWorkbook::new_empty(None);
        workbook_for_storage.add_sheet("Sheet1".to_string());
        workbook_for_storage.ensure_sheet_ids();
        workbook_for_storage.sheets[0].set_cell(
            0,
            0,
            Cell::from_literal(Some(CellScalar::Number(1.0))),
        );
        let mut formula_cell = Cell::from_formula("=UNSUPPORTED(A1)".to_string());
        formula_cell.computed_value = CellScalar::Number(99.0);
        workbook_for_storage.sheets[0].set_cell(0, 1, formula_cell);

        let workbook_id = import_app_workbook(&storage, &workbook_for_storage)?;

        // Simulate the in-memory engine failing to evaluate the formula by reporting an error.
        // The export should preserve the last known cached value from storage.
        let mut workbook_meta = workbook_for_storage.clone();
        let mut error_cell = Cell::from_formula("=UNSUPPORTED(A1)".to_string());
        error_cell.computed_value = CellScalar::Error("#NAME?".to_string());
        workbook_meta.sheets[0].set_cell(0, 1, error_cell);

        let tmp = tempfile::tempdir().context("temp dir")?;
        let out_path = tmp.path().join("cached-values.xlsx");
        let bytes = write_xlsx_from_storage(&storage, workbook_id, &workbook_meta, &out_path)?;

        let pkg = formula_xlsx::XlsxPackage::from_bytes(bytes.as_ref())
            .context("parse exported package")?;
        formula_xlsx::validate_opc_relationships(pkg.parts_map())
            .context("validate OPC relationships for exported workbook")?;

        let reread = formula_xlsx::read_workbook_from_reader(Cursor::new(bytes.as_ref()))
            .context("read exported workbook model")?;
        let sheet = reread.sheets.first().context("missing first sheet")?;
        let cell = sheet.cell(CellRef::new(0, 1)).context("missing B1 cell")?;
        assert_eq!(
            cell.value,
            formula_model::CellValue::Number(99.0),
            "expected export to preserve non-error cached value when the engine reports an error"
        );

        Ok(())
    }

    #[test]
    fn write_xlsx_from_storage_strips_activex_parts_for_macro_free_exports() -> anyhow::Result<()> {
        let fixture_path = Path::new(env!("CARGO_MANIFEST_DIR"))
            .join("../../../fixtures/xlsx/basic/activex-control.xlsx");
        let mut workbook_meta =
            crate::file_io::read_xlsx_blocking(&fixture_path).context("read activeX fixture")?;

        let preserved = workbook_meta
            .preserved_drawing_parts
            .as_ref()
            .context("expected activeX fixture to have preserved drawing parts")?;
        assert!(
            preserved.parts.contains_key("xl/activeX/activeX1.bin"),
            "expected preserved parts to include activeX binary"
        );
        assert!(
            preserved.parts.contains_key("xl/ctrlProps/ctrlProp1.xml"),
            "expected preserved parts to include control properties"
        );

        // Force the macro strip path (`WorkbookKind::Workbook` is macro-free) by indicating this
        // workbook contains VBA, even though the fixture itself is an `.xlsx`.
        workbook_meta.vba_project_bin = Some(b"dummy-vba".to_vec());

        let storage =
            formula_storage::Storage::open_in_memory().context("open in-memory storage")?;
        let workbook_id = import_app_workbook(&storage, &workbook_meta)?;

        let tmp = tempfile::tempdir().context("temp dir")?;
        let out_path = tmp.path().join("activex-stripped.xlsx");
        let bytes = write_xlsx_from_storage(&storage, workbook_id, &workbook_meta, &out_path)
            .context("export activeX workbook as macro-free xlsx")?;

        let pkg = formula_xlsx::XlsxPackage::from_bytes(bytes.as_ref()).context("parse export")?;
        formula_xlsx::validate_opc_relationships(pkg.parts_map())
            .context("validate OPC relationships for exported workbook")?;

        // Macro strip should remove ActiveX/controls parts so macro-free exports don't preserve
        // macro-capable artifacts like ActiveX binaries.
        assert!(
            pkg.part("xl/activeX/activeX1.bin").is_none(),
            "expected activeX binary to be stripped"
        );
        assert!(
            pkg.part("xl/activeX/activeX1.xml").is_none(),
            "expected activeX xml to be stripped"
        );
        assert!(
            pkg.part("xl/ctrlProps/ctrlProp1.xml").is_none(),
            "expected ctrlProps xml to be stripped"
        );

        let sheet_xml = pkg
            .part("xl/worksheets/sheet1.xml")
            .context("missing worksheet xml")?;
        let sheet_xml =
            std::str::from_utf8(sheet_xml).context("worksheet xml is not valid utf8")?;
        assert!(
            !sheet_xml.contains("<control ")
                && !sheet_xml.contains("<control>")
                && !sheet_xml.contains("<control/>")
                && !sheet_xml.contains("</control>"),
            "expected worksheet to have no <control> elements after macro strip, got:\n{sheet_xml}"
        );

        if let Some(sheet_rels) = pkg.part("xl/worksheets/_rels/sheet1.xml.rels") {
            let sheet_rels =
                std::str::from_utf8(sheet_rels).context("sheet rels is not valid utf8")?;
            assert!(
                !sheet_rels.contains("ctrlProps/"),
                "expected sheet relationships to stop referencing ctrlProps after macro strip, got:\n{sheet_rels}"
            );
            assert!(
                !sheet_rels.contains("activeX/"),
                "expected sheet relationships to stop referencing activeX after macro strip, got:\n{sheet_rels}"
            );
        }

        Ok(())
    }
}
