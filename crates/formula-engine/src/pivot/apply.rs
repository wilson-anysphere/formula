use std::collections::HashMap;

use chrono::Datelike;
use thiserror::Error;

use formula_model::{CellRef as ModelCellRef, CellValue, DateSystem, Style, Workbook, WorksheetId};

use crate::date::{ymd_to_serial, ExcelDate, ExcelDateSystem};

use super::{CellRef, CellWrite, Layout, PivotConfig, PivotResult, PivotValue, ShowAsType};

/// Options controlling how pivot results are rendered into worksheet cells.
#[derive(Debug, Clone, PartialEq, Eq)]
pub struct PivotApplyOptions {
    /// Whether to apply explicit number formats from [`super::ValueField::number_format`].
    ///
    /// When importing pivots from XLSX, this typically corresponds to
    /// `pivotTableDefinition@applyNumberFormats` (defaulting to `true` when missing).
    pub apply_number_formats: bool,
    /// Default number format code used for date values (`PivotValue::Date`) when no more specific
    /// format is available.
    pub default_date_number_format: String,
    /// Default number format code used for percent "show as" values when a value field does not
    /// have an explicit number format.
    pub default_percent_number_format: String,
}

impl Default for PivotApplyOptions {
    fn default() -> Self {
        Self {
            apply_number_formats: true,
            // Excel built-in 14 in en-US.
            default_date_number_format: "m/d/yyyy".to_string(),
            default_percent_number_format: "0.00%".to_string(),
        }
    }
}

#[derive(Debug, Error)]
pub enum PivotApplyError {
    #[error("worksheet not found: {0}")]
    SheetNotFound(WorksheetId),
}

impl PivotResult {
    /// Converts the computed pivot into a list of worksheet cell writes, including basic number
    /// format hints derived from `config`.
    ///
    /// These writes are suitable for application via [`apply_pivot_cell_writes_to_worksheet`].
    pub fn to_cell_writes_with_formats(
        &self,
        destination: CellRef,
        config: &PivotConfig,
        options: &PivotApplyOptions,
    ) -> Vec<CellWrite> {
        let value_field_count = config.value_fields.len();
        let row_label_width = match config.layout {
            Layout::Compact => 1,
            Layout::Outline | Layout::Tabular => config.row_fields.len(),
        };

        let mut out = Vec::new();
        for (r, row) in self.data.iter().enumerate() {
            for (c, value) in row.iter().enumerate() {
                let mut number_format = None;

                // Always apply a date format when the pivot output is a typed date.
                if matches!(value, PivotValue::Date(_)) {
                    number_format = Some(options.default_date_number_format.clone());
                } else if options.apply_number_formats
                    && r > 0
                    && value_field_count > 0
                    && c >= row_label_width
                {
                    let vf_idx = (c - row_label_width) % value_field_count;
                    let vf = &config.value_fields[vf_idx];

                    if let Some(fmt) = vf.number_format.clone() {
                        number_format = Some(fmt);
                    } else if is_percent_show_as(vf.show_as) {
                        number_format = Some(options.default_percent_number_format.clone());
                    }
                }

                out.push(CellWrite {
                    row: destination.row + r as u32,
                    col: destination.col + c as u32,
                    value: value.clone(),
                    number_format,
                });
            }
        }
        out
    }
}

/// Apply a list of pivot cell writes into a workbook worksheet, converting pivot values into
/// worksheet [`CellValue`]s and interning styles for any `number_format` hints.
pub fn apply_pivot_cell_writes_to_worksheet(
    workbook: &mut Workbook,
    sheet_id: WorksheetId,
    writes: &[CellWrite],
) -> Result<(), PivotApplyError> {
    let date_system = workbook_excel_date_system(workbook.date_system);
    let mut style_cache: HashMap<String, u32> = HashMap::new();

    for write in writes {
        let cell_ref = ModelCellRef::new(write.row, write.col);
        let cell_value = pivot_value_to_cell_value(&write.value, date_system);
        let style_id = write.number_format.as_ref().map(|fmt| {
            *style_cache.entry(fmt.clone()).or_insert_with(|| {
                workbook.styles.intern(Style {
                    number_format: Some(fmt.clone()),
                    ..Style::default()
                })
            })
        });

        let sheet = workbook
            .sheet_mut(sheet_id)
            .ok_or(PivotApplyError::SheetNotFound(sheet_id))?;
        sheet.set_value(cell_ref, cell_value);
        if let Some(style_id) = style_id {
            sheet.set_style_id(cell_ref, style_id);
        }
    }

    Ok(())
}

/// Convenience wrapper: generate formatted cell writes from a pivot result and apply them to the
/// worksheet.
pub fn apply_pivot_result_to_worksheet(
    workbook: &mut Workbook,
    sheet_id: WorksheetId,
    destination: ModelCellRef,
    result: &PivotResult,
    config: &PivotConfig,
    options: PivotApplyOptions,
) -> Result<(), PivotApplyError> {
    let destination = CellRef {
        row: destination.row,
        col: destination.col,
    };
    let writes = result.to_cell_writes_with_formats(destination, config, &options);
    apply_pivot_cell_writes_to_worksheet(workbook, sheet_id, &writes)
}

fn is_percent_show_as(show_as: Option<ShowAsType>) -> bool {
    matches!(
        show_as.unwrap_or(ShowAsType::Normal),
        ShowAsType::PercentOfGrandTotal
            | ShowAsType::PercentOfRowTotal
            | ShowAsType::PercentOfColumnTotal
            | ShowAsType::PercentOf
            | ShowAsType::PercentDifferenceFrom
    )
}

fn workbook_excel_date_system(system: DateSystem) -> ExcelDateSystem {
    match system {
        DateSystem::Excel1900 => ExcelDateSystem::Excel1900 { lotus_compat: true },
        DateSystem::Excel1904 => ExcelDateSystem::Excel1904,
    }
}

fn pivot_value_to_cell_value(value: &PivotValue, date_system: ExcelDateSystem) -> CellValue {
    match value {
        PivotValue::Blank => CellValue::Empty,
        PivotValue::Number(n) => CellValue::Number(*n),
        PivotValue::Text(s) => CellValue::String(s.clone()),
        PivotValue::Bool(b) => CellValue::Boolean(*b),
        PivotValue::Date(d) => {
            let excel_date = ExcelDate::new(d.year(), d.month() as u8, d.day() as u8);
            let serial = ymd_to_serial(excel_date, date_system).unwrap_or(0) as f64;
            CellValue::Number(serial)
        }
    }
}
