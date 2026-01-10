use crate::sort_filter::sort::{compute_header_rows, compute_row_permutation};
use crate::sort_filter::{apply_autofilter, AutoFilter, CellValue, FilterResult, RowPermutation, SortSpec};
use crate::{parse_formula, CellAddr, LocaleConfig, ParseOptions, SerializeOptions};
use formula_model::{CellRef, CellValue as ModelCellValue, Outline, Range, RowProperties, Worksheet};

pub fn sort_worksheet_range(sheet: &mut Worksheet, range: Range, spec: &SortSpec) -> RowPermutation {
    let row_count = range.height() as usize;
    if row_count <= 1 || spec.keys.is_empty() {
        return RowPermutation {
            new_to_old: (0..row_count).collect(),
            old_to_new: (0..row_count).collect(),
        };
    }

    let start_row = range.start.row;
    let start_col = range.start.col;

    let width = range.width() as usize;
    let cell_at = |sheet: &Worksheet, local_row: usize, local_col: usize| -> CellValue {
        if local_col >= width {
            return CellValue::Blank;
        }
        let row = start_row + local_row as u32;
        let col = start_col + local_col as u32;
        model_cell_value_to_sort_value(&sheet.value(CellRef::new(row, col)))
    };

    let header_rows = compute_header_rows(row_count, spec.header, &spec.keys, |r, c| cell_at(sheet, r, c));
    let perm = compute_row_permutation(row_count, header_rows, &spec.keys, |r, c| cell_at(sheet, r, c));

    // Nothing to permute (e.g. header-only range).
    if header_rows >= row_count {
        return perm;
    }

    let data_start_row = start_row + header_rows as u32;
    if data_start_row > range.end.row {
        return perm;
    }

    let data_range = Range::new(
        CellRef::new(data_start_row, range.start.col),
        CellRef::new(range.end.row, range.end.col),
    );

    let moved_cells = sheet
        .iter_cells_in_range(data_range)
        .map(|(cell_ref, cell)| (cell_ref, cell.clone()))
        .collect::<Vec<_>>();

    sheet.clear_range(data_range);

    for (cell_ref, mut cell) in moved_cells {
        let local_old_row = (cell_ref.row - start_row) as usize;
        let local_new_row = perm.old_to_new[local_old_row];
        let new_row = start_row + local_new_row as u32;

        if let Some(formula) = cell.formula.as_deref() {
            if let Some(rewritten) = rewrite_formula_for_move(
                formula,
                CellAddr::new(cell_ref.row, cell_ref.col),
                CellAddr::new(new_row, cell_ref.col),
            ) {
                cell.formula = Some(rewritten);
            }
        }

        sheet.set_cell(CellRef::new(new_row, cell_ref.col), cell);
    }

    permute_row_properties(sheet, start_row, header_rows, range.end.row, &perm);

    perm
}

pub fn apply_autofilter_to_outline(
    sheet: &Worksheet,
    outline: &mut Outline,
    range: Range,
    filter: Option<&AutoFilter>,
) -> FilterResult {
    let row_count = range.height() as usize;
    let col_count = range.width() as usize;

    if row_count == 0 || col_count == 0 {
        return FilterResult {
            visible_rows: Vec::new(),
            hidden_sheet_rows: Vec::new(),
        };
    }

    // Always clear any existing filter-hidden flags for the data rows within the range.
    // AutoFilter treats the first row as the header row, so we never set filter hidden on it.
    let header_row_1based = range.start.row + 1;
    let data_start_row_1based = header_row_1based + 1;
    let end_row_1based = range.end.row + 1;
    if data_start_row_1based <= end_row_1based {
        outline
            .rows
            .clear_filter_hidden_range(data_start_row_1based, end_row_1based);
    }

    let Some(filter) = filter else {
        return FilterResult {
            visible_rows: vec![true; row_count],
            hidden_sheet_rows: Vec::new(),
        };
    };

    let range_ref = crate::sort_filter::RangeRef {
        start_row: range.start.row as usize,
        start_col: range.start.col as usize,
        end_row: range.end.row as usize,
        end_col: range.end.col as usize,
    };

    let mut rows = Vec::with_capacity(row_count);
    for local_row in 0..row_count {
        let row_idx = range.start.row + local_row as u32;
        let mut row = Vec::with_capacity(col_count);
        for local_col in 0..col_count {
            let col_idx = range.start.col + local_col as u32;
            let value = sheet.value(CellRef::new(row_idx, col_idx));
            row.push(model_cell_value_to_sort_value(&value));
        }
        rows.push(row);
    }

    let range_data = crate::sort_filter::RangeData::new(range_ref, rows)
        .expect("worksheet range should always produce rectangular RangeData");

    let result = apply_autofilter(&range_data, filter);

    for hidden_row_0based in &result.hidden_sheet_rows {
        outline.rows.set_filter_hidden((*hidden_row_0based as u32) + 1, true);
    }

    result
}

fn permute_row_properties(
    sheet: &mut Worksheet,
    start_row: u32,
    header_rows: usize,
    end_row: u32,
    perm: &RowPermutation,
) {
    let data_start = start_row + header_rows as u32;
    if data_start > end_row {
        return;
    }

    let mut extracted: Vec<(u32, RowProperties)> = Vec::new();
    for row in data_start..=end_row {
        if let Some(props) = sheet.row_properties.remove(&row) {
            extracted.push((row, props));
        }
    }

    for (old_row, props) in extracted {
        let local_old = (old_row - start_row) as usize;
        let local_new = perm.old_to_new[local_old];
        let new_row = start_row + local_new as u32;
        sheet.row_properties.insert(new_row, props);
    }
}

fn rewrite_formula_for_move(formula: &str, from: CellAddr, to: CellAddr) -> Option<String> {
    let opts = ParseOptions {
        locale: LocaleConfig::en_us(),
        normalize_relative_to: Some(from),
    };
    let ast = parse_formula(formula, opts).ok()?;

    let out_opts = SerializeOptions {
        locale: LocaleConfig::en_us(),
        include_xlfn_prefix: true,
        origin: Some(to),
        omit_equals: false,
    };

    ast.to_string(out_opts).ok()
}

fn model_cell_value_to_sort_value(value: &ModelCellValue) -> CellValue {
    match value {
        ModelCellValue::Empty => CellValue::Blank,
        ModelCellValue::Number(n) => CellValue::Number(*n),
        ModelCellValue::String(s) => CellValue::Text(s.clone()),
        ModelCellValue::Boolean(b) => CellValue::Bool(*b),
        ModelCellValue::Error(err) => CellValue::Text(err.to_string()),
        ModelCellValue::RichText(rt) => CellValue::Text(rt.plain_text().to_string()),
        ModelCellValue::Array(_) => CellValue::Blank,
        ModelCellValue::Spill(_) => CellValue::Blank,
    }
}
