use super::PrintError;
use core::fmt::Write as _;
use formula_model::sheet_name_eq_case_insensitive;

#[derive(Debug, Clone, Copy, PartialEq, Eq)]
pub struct CellRange {
    /// 1-based row number.
    pub start_row: u32,
    /// 1-based row number.
    pub end_row: u32,
    /// 1-based column number.
    pub start_col: u32,
    /// 1-based column number.
    pub end_col: u32,
}

impl CellRange {
    pub fn normalized(self) -> Self {
        Self {
            start_row: self.start_row.min(self.end_row),
            end_row: self.start_row.max(self.end_row),
            start_col: self.start_col.min(self.end_col),
            end_col: self.start_col.max(self.end_col),
        }
    }
}

#[derive(Debug, Clone, Copy, PartialEq, Eq)]
pub struct RowRange {
    /// 1-based row number.
    pub start: u32,
    /// 1-based row number.
    pub end: u32,
}

impl RowRange {
    pub fn normalized(self) -> Self {
        Self {
            start: self.start.min(self.end),
            end: self.start.max(self.end),
        }
    }
}

#[derive(Debug, Clone, Copy, PartialEq, Eq)]
pub struct ColRange {
    /// 1-based column number.
    pub start: u32,
    /// 1-based column number.
    pub end: u32,
}

impl ColRange {
    pub fn normalized(self) -> Self {
        Self {
            start: self.start.min(self.end),
            end: self.start.max(self.end),
        }
    }
}

#[derive(Debug, Clone, Copy, PartialEq, Eq, Default)]
pub struct PrintTitles {
    pub repeat_rows: Option<RowRange>,
    pub repeat_cols: Option<ColRange>,
}

#[derive(Debug, Clone, PartialEq, Eq)]
struct AreaRef {
    sheet_name: String,
    range: A1Range,
}

#[derive(Debug, Clone, PartialEq, Eq)]
enum A1Range {
    Cell(CellRange),
    Row(RowRange),
    Col(ColRange),
}

pub fn parse_print_area_defined_name(
    expected_sheet_name: &str,
    formula: &str,
) -> Result<Vec<CellRange>, PrintError> {
    let areas = split_areas(formula)?
        .into_iter()
        .map(parse_area_ref)
        .collect::<Result<Vec<_>, _>>()?;

    let mut ranges = Vec::new();
    for area in areas {
        if !sheet_name_eq_case_insensitive(&area.sheet_name, expected_sheet_name) {
            return Err(PrintError::InvalidA1(format!(
                "expected sheet {expected_sheet_name:?}, found {found:?}",
                found = area.sheet_name
            )));
        }

        match area.range {
            A1Range::Cell(r) => ranges.push(r.normalized()),
            A1Range::Row(_) | A1Range::Col(_) => {
                return Err(PrintError::InvalidA1(
                    "print area must be a cell range".to_string(),
                ))
            }
        }
    }

    Ok(ranges)
}

pub fn parse_print_titles_defined_name(
    expected_sheet_name: &str,
    formula: &str,
) -> Result<PrintTitles, PrintError> {
    let areas = split_areas(formula)?
        .into_iter()
        .map(parse_area_ref)
        .collect::<Result<Vec<_>, _>>()?;

    let mut titles = PrintTitles::default();
    for area in areas {
        if !sheet_name_eq_case_insensitive(&area.sheet_name, expected_sheet_name) {
            return Err(PrintError::InvalidA1(format!(
                "expected sheet {expected_sheet_name:?}, found {found:?}",
                found = area.sheet_name
            )));
        }

        match area.range {
            A1Range::Row(r) => titles.repeat_rows = Some(r.normalized()),
            A1Range::Col(r) => titles.repeat_cols = Some(r.normalized()),
            A1Range::Cell(r) => {
                // Some producers represent whole-row/whole-column print titles as explicit
                // cell ranges (e.g. `$A$1:$IV$1`, `$A$1:$A$65536`) instead of row/col-only
                // references (`$1:$1`, `$A:$A`).
                //
                // Best-effort: interpret single-row / single-column cell ranges as print titles.
                let r = r.normalized();
                if r.start_row == r.end_row && r.start_col != r.end_col {
                    titles.repeat_rows = Some(RowRange {
                        start: r.start_row,
                        end: r.end_row,
                    });
                } else if r.start_col == r.end_col && r.start_row != r.end_row {
                    titles.repeat_cols = Some(ColRange {
                        start: r.start_col,
                        end: r.end_col,
                    });
                } else {
                    return Err(PrintError::InvalidA1(
                        "print titles must be a row or column range".to_string(),
                    ));
                }
            }
        }
    }

    Ok(titles)
}

pub fn format_print_area_defined_name(sheet_name: &str, ranges: &[CellRange]) -> String {
    let sheet = format_sheet_name(sheet_name);
    let mut out = String::new();
    for (idx, range) in ranges.iter().copied().enumerate() {
        if idx > 0 {
            out.push(',');
        }
        out.push_str(&sheet);
        out.push('!');
        push_cell_range(&mut out, range);
    }
    out
}

pub fn format_print_titles_defined_name(sheet_name: &str, titles: &PrintTitles) -> String {
    let sheet = format_sheet_name(sheet_name);
    let mut out = String::new();
    if let Some(rows) = titles.repeat_rows {
        out.push_str(&sheet);
        out.push('!');
        push_row_range(&mut out, rows);
    }
    if let Some(cols) = titles.repeat_cols {
        if !out.is_empty() {
            out.push(',');
        }
        out.push_str(&sheet);
        out.push('!');
        push_col_range(&mut out, cols);
    }
    out
}

fn split_areas(formula: &str) -> Result<Vec<&str>, PrintError> {
    let mut parts = Vec::new();
    let mut start = 0usize;
    let mut in_quotes = false;
    let bytes = formula.as_bytes();
    let mut i = 0usize;

    while i < bytes.len() {
        match bytes[i] {
            b'\'' => {
                if in_quotes {
                    if bytes.get(i..).and_then(|s| s.get(1)) == Some(&b'\'') {
                        // Escaped quote in a sheet name.
                        i += 1;
                    } else {
                        in_quotes = false;
                    }
                } else {
                    in_quotes = true;
                }
            }
            b',' if !in_quotes => {
                let part = formula[start..i].trim();
                if !part.is_empty() {
                    parts.push(part);
                }
                start = i + 1;
            }
            _ => {}
        }

        i += 1;
    }

    let part = formula[start..].trim();
    if !part.is_empty() {
        parts.push(part);
    }

    Ok(parts)
}

fn parse_area_ref(area: &str) -> Result<AreaRef, PrintError> {
    let (sheet_name, rest) = split_sheet_name(area)?;
    let range = parse_range(rest)?;

    Ok(AreaRef { sheet_name, range })
}

fn split_sheet_name(input: &str) -> Result<(String, &str), PrintError> {
    let bytes = input.as_bytes();
    if bytes.is_empty() {
        return Err(PrintError::InvalidA1("empty reference".to_string()));
    }

    if bytes[0] == b'\'' {
        let mut sheet = String::new();
        let mut i = 1usize;
        while i < bytes.len() {
            match bytes[i] {
                b'\'' => {
                    if bytes.get(i..).and_then(|s| s.get(1)) == Some(&b'\'') {
                        sheet.push('\'');
                        i += 2;
                        continue;
                    }

                    // End of quoted sheet name.
                    let Some(bang_idx) = i.checked_add(1) else {
                        return Err(PrintError::InvalidA1(format!(
                            "expected ! after quoted sheet name in {input:?}"
                        )));
                    };
                    if bytes.get(bang_idx) != Some(&b'!') {
                        return Err(PrintError::InvalidA1(format!(
                            "expected ! after quoted sheet name in {input:?}"
                        )));
                    }

                    let Some(rest_start) = i.checked_add(2) else {
                        return Err(PrintError::InvalidA1(format!(
                            "expected ! after quoted sheet name in {input:?}"
                        )));
                    };
                    let rest = input.get(rest_start..).ok_or_else(|| {
                        PrintError::InvalidA1(format!(
                            "invalid utf-8 boundary after quoted sheet name in {input:?}"
                        ))
                    })?;
                    return Ok((sheet, rest));
                }
                _ => {
                    let ch = input
                        .get(i..)
                        .and_then(|s| s.chars().next())
                        .ok_or_else(|| {
                            PrintError::InvalidA1(format!(
                                "invalid utf-8 boundary while parsing quoted sheet name in {input:?}"
                            ))
                        })?;
                    sheet.push(ch);
                    i += ch.len_utf8();
                    continue;
                }
            }
        }

        return Err(PrintError::InvalidA1(format!(
            "unterminated quoted sheet name in {input:?}"
        )));
    }

    let Some(idx) = input.find('!') else {
        return Err(PrintError::InvalidA1(format!(
            "expected ! in area reference {input:?}"
        )));
    };

    Ok((input[..idx].to_string(), &input[(idx + 1)..]))
}

fn parse_range(ref_str: &str) -> Result<A1Range, PrintError> {
    let (start, end) = if let Some(idx) = ref_str.find(':') {
        (&ref_str[..idx], &ref_str[(idx + 1)..])
    } else {
        (ref_str, ref_str)
    };

    let start = parse_endpoint(start)?;
    let end = parse_endpoint(end)?;

    match (start, end) {
        (Endpoint::Cell(a), Endpoint::Cell(b)) => Ok(A1Range::Cell(CellRange {
            start_row: a.row,
            end_row: b.row,
            start_col: a.col,
            end_col: b.col,
        })),
        (Endpoint::Row(a), Endpoint::Row(b)) => Ok(A1Range::Row(RowRange { start: a, end: b })),
        (Endpoint::Col(a), Endpoint::Col(b)) => Ok(A1Range::Col(ColRange { start: a, end: b })),
        _ => Err(PrintError::InvalidA1(format!(
            "mismatched range endpoints in {ref_str:?}"
        ))),
    }
}

#[derive(Debug, Clone, Copy, PartialEq, Eq)]
struct CellRef {
    row: u32,
    col: u32,
}

#[derive(Debug, Clone, Copy, PartialEq, Eq)]
enum Endpoint {
    Cell(CellRef),
    Row(u32),
    Col(u32),
}

fn parse_endpoint(s: &str) -> Result<Endpoint, PrintError> {
    let endpoint = formula_model::parse_a1_endpoint(s).map_err(|err| {
        PrintError::InvalidA1(format!("invalid endpoint {s:?}: {err}"))
    })?;

    match endpoint {
        formula_model::A1Endpoint::Cell(cell) => {
            let row = cell.row.checked_add(1).ok_or_else(|| {
                PrintError::InvalidA1(format!("row out of range in endpoint {s:?}"))
            })?;
            let col0 = cell.col.checked_add(1).ok_or_else(|| {
                PrintError::InvalidA1(format!("col out of range in endpoint {s:?}"))
            })?;
            Ok(Endpoint::Cell(CellRef { row, col: col0 }))
        }
        formula_model::A1Endpoint::Row(row0) => {
            let row = row0.checked_add(1).ok_or_else(|| {
                PrintError::InvalidA1(format!("row out of range in endpoint {s:?}"))
            })?;
            Ok(Endpoint::Row(row))
        }
        formula_model::A1Endpoint::Col(col0) => {
            let col = col0.checked_add(1).ok_or_else(|| {
                PrintError::InvalidA1(format!("col out of range in endpoint {s:?}"))
            })?;
            Ok(Endpoint::Col(col))
        }
    }
}

fn format_sheet_name(sheet_name: &str) -> String {
    let mut out = String::new();
    formula_model::push_sheet_name_a1(&mut out, sheet_name);
    out
}

fn push_cell_ref(out: &mut String, row: u32, col: u32) {
    if row == 0 || col == 0 {
        // Best-effort fallback for malformed inputs.
        out.push('$');
        if col > 0 {
            formula_model::push_column_label(col - 1, out);
        }
        out.push('$');
        let _ = write!(out, "{row}");
        return;
    }

    formula_model::push_a1_cell_ref_row1(u64::from(row), col - 1, true, true, out);
}

fn push_cell_range(out: &mut String, range: CellRange) {
    let range = range.normalized();
    if range.start_row == 0
        || range.end_row == 0
        || range.start_col == 0
        || range.end_col == 0
    {
        // Best-effort fallback for malformed inputs.
        push_cell_ref(out, range.start_row, range.start_col);
        out.push(':');
        push_cell_ref(out, range.end_row, range.end_col);
        return;
    }

    formula_model::push_a1_cell_range_row1(
        u64::from(range.start_row),
        range.start_col - 1,
        u64::from(range.end_row),
        range.end_col - 1,
        true,
        true,
        out,
    );
}

fn push_row_range(out: &mut String, range: RowRange) {
    let range = range.normalized();
    formula_model::push_a1_row_range_row1(u64::from(range.start), u64::from(range.end), true, out);
}

fn push_col_range(out: &mut String, range: ColRange) {
    let range = range.normalized();
    if range.start == 0 || range.end == 0 {
        // Best-effort fallback for malformed inputs.
        out.push('$');
        if range.start > 0 {
            formula_model::push_column_label(range.start - 1, out);
        }
        out.push(':');
        out.push('$');
        if range.end > 0 {
            formula_model::push_column_label(range.end - 1, out);
        }
        return;
    }

    formula_model::push_a1_col_range(range.start - 1, range.end - 1, true, out);
}

#[cfg(test)]
mod tests {
    use super::*;

    #[test]
    fn parse_print_area_defined_name_matches_unicode_sheet_names_case_insensitive_like_excel() {
        let ranges = parse_print_area_defined_name("Straße", "'STRASSE'!$A$1:$A$1")
            .expect("should parse print area with Unicode-aware sheet matching");
        assert_eq!(
            ranges,
            vec![CellRange {
                start_row: 1,
                end_row: 1,
                start_col: 1,
                end_col: 1
            }]
        );
    }

    #[test]
    fn parse_print_titles_defined_name_matches_unicode_sheet_names_case_insensitive_like_excel() {
        let titles = parse_print_titles_defined_name("Straße", "'STRASSE'!$1:$1,'STRASSE'!$A:$A")
            .expect("should parse print titles with Unicode-aware sheet matching");
        assert_eq!(titles.repeat_rows, Some(RowRange { start: 1, end: 1 }));
        assert_eq!(titles.repeat_cols, Some(ColRange { start: 1, end: 1 }));
    }

    #[test]
    fn format_print_defined_names_quote_sheet_names_that_look_like_tokens() {
        let range = CellRange {
            start_row: 1,
            end_row: 1,
            start_col: 1,
            end_col: 1,
        };

        assert_eq!(
            format_print_area_defined_name("A1", &[range]),
            "'A1'!$A$1"
        );
        assert_eq!(
            format_print_area_defined_name("TRUE", &[range]),
            "'TRUE'!$A$1"
        );
        assert_eq!(
            format_print_titles_defined_name("R1C1", &PrintTitles { repeat_rows: Some(RowRange { start: 1, end: 1 }), repeat_cols: None }),
            "'R1C1'!$1:$1"
        );
    }
}
