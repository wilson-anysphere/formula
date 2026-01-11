use formula_model::{CellRef, Range, Worksheet};

pub(crate) fn worksheet_used_range(sheet: &Worksheet) -> Option<Range> {
    if let Some(range) = sheet.used_range() {
        return Some(range);
    }

    let mut min_cell: Option<CellRef> = None;
    let mut max_cell: Option<CellRef> = None;
    for (cell_ref, _) in sheet.iter_cells() {
        min_cell = Some(match min_cell {
            Some(min) => CellRef::new(min.row.min(cell_ref.row), min.col.min(cell_ref.col)),
            None => cell_ref,
        });
        max_cell = Some(match max_cell {
            Some(max) => CellRef::new(max.row.max(cell_ref.row), max.col.max(cell_ref.col)),
            None => cell_ref,
        });
    }

    match (min_cell, max_cell) {
        (Some(start), Some(end)) => Some(Range::new(start, end)),
        _ => None,
    }
}

pub(crate) fn worksheet_dimension_range(sheet: &Worksheet) -> Range {
    worksheet_used_range(sheet).unwrap_or_else(|| Range::new(CellRef::new(0, 0), CellRef::new(0, 0)))
}

pub(crate) fn parse_dimension_ref(value: &str) -> Option<Range> {
    let trimmed = value.trim();
    if trimmed.is_empty() {
        return None;
    }

    let (start, end) = match trimmed.split_once(':') {
        Some((a, b)) => (a.trim(), b.trim()),
        None => (trimmed, trimmed),
    };

    let start = CellRef::from_a1(start).ok()?;
    let end = CellRef::from_a1(end).ok()?;
    Some(Range::new(start, end))
}

