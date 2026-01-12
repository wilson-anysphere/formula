mod workbook_state;

pub use workbook_state::{PersistentWorkbookState, WorkbookPersistenceLocation};
pub use workbook_state::{open_memory_manager, open_storage};

use crate::file_io::{
    DefinedName as AppDefinedName, Sheet as AppSheet, Table as AppTable, Workbook as AppWorkbook,
};
use crate::state::{Cell, CellScalar};
use anyhow::Context;
use directories::ProjectDirs;
use formula_model::{
    display_formula_text, normalize_formula_text, Cell as ModelCell, CellRef,
    CellValue as ModelCellValue, DefinedNameScope, Workbook as ModelWorkbook,
};
use sha2::{Digest, Sha256};
use std::io::Cursor;
use std::path::Path;
use std::path::PathBuf;
use std::sync::Arc;
use uuid::Uuid;

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

    for sheet in &workbook.sheets {
        let sheet_id = model.add_sheet(sheet.name.clone())?;
        let Some(model_sheet) = model.sheet_mut(sheet_id) else {
            continue;
        };

        for ((row, col), cell) in sheet.cells_iter() {
            let cell_ref = CellRef::new(row as u32, col as u32);

            let out = match (&cell.formula, &cell.input_value) {
                (Some(formula), _) => {
                    let mut c = ModelCell::new(scalar_to_model_value(&cell.computed_value));
                    c.formula = normalize_formula_text(formula);
                    c
                }
                (None, Some(value)) => ModelCell::new(scalar_to_model_value(value)),
                (None, None) => ModelCell::new(ModelCellValue::Empty),
            };

            if out.is_truly_empty() {
                continue;
            }

            model_sheet.set_cell(cell_ref, out);
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

    for (cell_ref, cell) in sheet.iter_cells() {
        let row = cell_ref.row as usize;
        let col = cell_ref.col as usize;
        let number_format = style_table
            .get(cell.style_id)
            .and_then(|s| s.number_format.clone());

        let cached_value = model_value_to_scalar(&cell.value);
        if let Some(formula) = cell.formula.as_deref() {
            if formula.trim().is_empty() {
                continue;
            }
            let normalized = display_formula_text(formula);
            let mut c = Cell::from_formula(normalized);
            c.computed_value = cached_value;
            c.number_format = number_format;
            out.set_cell(row, col, c);
            continue;
        }

        if matches!(cached_value, CellScalar::Empty) {
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
        ModelCellValue::Array(arr) => CellScalar::Text(format!("{:?}", arr.data)),
        ModelCellValue::Spill(_) => CellScalar::Error("#SPILL!".to_string()),
    }
}

/// Export a workbook from SQLite and write it as an `.xlsx`/`.xlsm` file.
///
/// We currently write a fresh XLSX ZIP from the `formula-model` export and then
/// re-apply a handful of preserved parts (VBA, drawing parts, pivot attachments)
/// plus workbook print settings.
///
/// This keeps SQLite as the source of truth for the current workbook state and
/// supports autosave/crash recovery. We intentionally do **not** use the older
/// patch-based XLSX save path here yet because emitting `WorkbookCellPatches`
/// directly from SQLite deltas is not implemented.
pub fn write_xlsx_from_storage(
    storage: &formula_storage::Storage,
    workbook_id: Uuid,
    workbook_meta: &AppWorkbook,
    path: &Path,
) -> anyhow::Result<Arc<[u8]>> {
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
    formula_xlsx::write_workbook_to_writer(&model, &mut cursor).context("write workbook to bytes")?;
    let mut bytes = cursor.into_inner();

    let extension = path
        .extension()
        .and_then(|s| s.to_str())
        .map(|s| s.to_ascii_lowercase());
    let wants_vba =
        workbook_meta.vba_project_bin.is_some() && matches!(extension.as_deref(), Some("xlsm"));
    let wants_preserved_drawings = workbook_meta.preserved_drawing_parts.is_some();
    let wants_preserved_pivots = workbook_meta.preserved_pivot_parts.is_some();
    let wants_power_query = workbook_meta.power_query_xml.is_some();

    if wants_vba || wants_preserved_drawings || wants_preserved_pivots || wants_power_query {
        let mut pkg = formula_xlsx::XlsxPackage::from_bytes(&bytes).context("parse generated xlsx")?;

        if wants_vba {
            pkg.set_part(
                "xl/vbaProject.bin",
                workbook_meta
                    .vba_project_bin
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

        bytes = pkg.write_to_bytes().context("repack xlsx package")?;
    }

    if matches!(extension.as_deref(), Some("xlsx") | Some("xlsm")) {
        bytes = formula_xlsx::print::write_workbook_print_settings(&bytes, &workbook_meta.print_settings)
            .context("write workbook print settings")?;
    }

    let bytes = Arc::<[u8]>::from(bytes);
    std::fs::write(path, bytes.as_ref()).with_context(|| format!("write workbook {path:?}"))?;
    Ok(bytes)
}

fn apply_cached_formula_values(model: &mut ModelWorkbook, workbook: &AppWorkbook) {
    for sheet in &workbook.sheets {
        let Some(model_sheet) = model
            .sheets
            .iter_mut()
            .find(|s| s.name.eq_ignore_ascii_case(&sheet.name))
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
            if matches!(computed, CellScalar::Error(_)) && !matches!(existing, CellScalar::Error(_)) {
                continue;
            }

            if computed != existing {
                model_sheet.set_value(cell_ref, scalar_to_model_value(&computed));
            }
        }
    }
}
