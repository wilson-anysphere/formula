use super::PrintError;
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
    ranges
        .iter()
        .map(|r| {
            format!(
                "{sheet}!{ref_str}",
                sheet = format_sheet_name(sheet_name),
                ref_str = format_cell_range(*r)
            )
        })
        .collect::<Vec<_>>()
        .join(",")
}

pub fn format_print_titles_defined_name(sheet_name: &str, titles: &PrintTitles) -> String {
    let sheet = format_sheet_name(sheet_name);
    let mut parts = Vec::new();

    if let Some(rows) = titles.repeat_rows {
        parts.push(format!(
            "{sheet}!{ref_str}",
            ref_str = format_row_range(rows)
        ));
    }
    if let Some(cols) = titles.repeat_cols {
        parts.push(format!(
            "{sheet}!{ref_str}",
            ref_str = format_col_range(cols)
        ));
    }

    parts.join(",")
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
                    if i + 1 < bytes.len() && bytes[i + 1] == b'\'' {
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
                    if i + 1 < bytes.len() && bytes[i + 1] == b'\'' {
                        sheet.push('\'');
                        i += 2;
                        continue;
                    }

                    // End of quoted sheet name.
                    if i + 1 >= bytes.len() || bytes[i + 1] != b'!' {
                        return Err(PrintError::InvalidA1(format!(
                            "expected ! after quoted sheet name in {input:?}"
                        )));
                    }

                    let rest = &input[(i + 2)..];
                    return Ok((sheet, rest));
                }
                _ => sheet.push(bytes[i] as char),
            }
            i += 1;
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
    let trimmed = s.trim().trim_matches('$');
    if trimmed.is_empty() {
        return Err(PrintError::InvalidA1("empty endpoint".to_string()));
    }

    let mut letters = String::new();
    let mut digits = String::new();

    for ch in trimmed.chars() {
        if ch == '$' {
            continue;
        }
        if ch.is_ascii_alphabetic() && digits.is_empty() {
            letters.push(ch);
        } else if ch.is_ascii_digit() {
            digits.push(ch);
        } else {
            return Err(PrintError::InvalidA1(format!(
                "invalid character {ch:?} in endpoint {s:?}"
            )));
        }
    }

    match (letters.is_empty(), digits.is_empty()) {
        (false, false) => Ok(Endpoint::Cell(CellRef {
            col: letters_to_col(&letters)?,
            row: digits.parse::<u32>().map_err(|_| {
                PrintError::InvalidA1(format!("invalid row number in endpoint {s:?}"))
            })?,
        })),
        (false, true) => Ok(Endpoint::Col(letters_to_col(&letters)?)),
        (true, false) => Ok(Endpoint::Row(digits.parse::<u32>().map_err(|_| {
            PrintError::InvalidA1(format!("invalid row number in endpoint {s:?}"))
        })?)),
        (true, true) => Err(PrintError::InvalidA1(format!("invalid endpoint {s:?}"))),
    }
}

fn letters_to_col(letters: &str) -> Result<u32, PrintError> {
    let mut col = 0u32;
    for ch in letters.chars() {
        if !ch.is_ascii_alphabetic() {
            return Err(PrintError::InvalidA1(format!(
                "invalid column letters {letters:?}"
            )));
        }
        let digit = (ch.to_ascii_uppercase() as u8 - b'A' + 1) as u32;
        col = col
            .checked_mul(26)
            .and_then(|c| c.checked_add(digit))
            .ok_or_else(|| {
                PrintError::InvalidA1(format!("invalid column letters {letters:?}"))
            })?;
    }

    if col == 0 {
        return Err(PrintError::InvalidA1(format!(
            "invalid column letters {letters:?}"
        )));
    }
    Ok(col)
}

fn col_to_letters(mut col: u32) -> String {
    let mut chars = Vec::new();
    while col > 0 {
        let rem = ((col - 1) % 26) as u8;
        chars.push((b'A' + rem) as char);
        col = (col - 1) / 26;
    }
    chars.iter().rev().collect()
}

fn format_sheet_name(sheet_name: &str) -> String {
    if sheet_name
        .chars()
        .all(|c| c.is_ascii_alphanumeric() || c == '_')
    {
        return sheet_name.to_string();
    }

    let escaped = sheet_name.replace('\'', "''");
    format!("'{escaped}'")
}

fn format_cell_ref(row: u32, col: u32) -> String {
    format!("${col}${row}", col = col_to_letters(col), row = row)
}

fn format_cell_range(range: CellRange) -> String {
    let range = range.normalized();
    if range.start_row == range.end_row && range.start_col == range.end_col {
        return format_cell_ref(range.start_row, range.start_col);
    }

    format!(
        "{start}:{end}",
        start = format_cell_ref(range.start_row, range.start_col),
        end = format_cell_ref(range.end_row, range.end_col)
    )
}

fn format_row_range(range: RowRange) -> String {
    let range = range.normalized();
    if range.start == range.end {
        return format!("${row}:${row}", row = range.start);
    }
    format!("${start}:${end}", start = range.start, end = range.end)
}

fn format_col_range(range: ColRange) -> String {
    let range = range.normalized();
    let start = col_to_letters(range.start);
    let end = col_to_letters(range.end);
    if range.start == range.end {
        return format!("${col}:${col}", col = start);
    }
    format!("${start}:${end}")
}
