use std::cmp::{max, min};

#[derive(Clone, Debug, PartialEq, Eq)]
pub(crate) struct SheetPrefix {
    /// Exact prefix text as it appears in the formula, including the trailing `!`.
    raw: String,
    /// Normalized (unescaped) sheet name without workbook qualifier.
    sheet_name: String,
    external_workbook: bool,
}

#[derive(Clone, Copy, Debug, PartialEq, Eq)]
pub(crate) struct GridRange {
    pub start_row: u32,
    pub start_col: u32,
    pub end_row: u32,
    pub end_col: u32,
}

impl GridRange {
    pub(crate) fn new(start_row: u32, start_col: u32, end_row: u32, end_col: u32) -> Self {
        let sr = min(start_row, end_row);
        let er = max(start_row, end_row);
        let sc = min(start_col, end_col);
        let ec = max(start_col, end_col);
        Self {
            start_row: sr,
            start_col: sc,
            end_row: er,
            end_col: ec,
        }
    }

    pub(crate) fn contains(&self, row: u32, col: u32) -> bool {
        row >= self.start_row && row <= self.end_row && col >= self.start_col && col <= self.end_col
    }
}

#[derive(Clone, Copy, Debug, PartialEq, Eq)]
pub(crate) struct A1CellRef {
    pub row: u32,
    pub col: u32,
    pub row_abs: bool,
    pub col_abs: bool,
}

#[derive(Clone, Copy, Debug, PartialEq, Eq)]
pub(crate) struct A1ColRef {
    pub col: u32,
    pub abs: bool,
}

#[derive(Clone, Copy, Debug, PartialEq, Eq)]
pub(crate) struct A1RowRef {
    pub row: u32,
    pub abs: bool,
}

#[derive(Clone, Debug, PartialEq, Eq)]
pub(crate) enum RefKind {
    Cell(A1CellRef),
    CellRange { start: A1CellRef, end: A1CellRef },
    ColRange { start: A1ColRef, end: A1ColRef },
    RowRange { start: A1RowRef, end: A1RowRef },
    RefError,
}

#[derive(Clone, Debug, PartialEq, Eq)]
pub(crate) struct ParsedRef {
    prefix: Option<SheetPrefix>,
    kind: RefKind,
}

#[derive(Clone, Copy, Debug)]
struct RefTarget<'a> {
    sheet: &'a str,
    external_workbook: bool,
}

impl ParsedRef {
    fn target_sheet<'a>(&'a self, ctx_sheet: &'a str) -> RefTarget<'a> {
        match &self.prefix {
            Some(prefix) => RefTarget {
                sheet: &prefix.sheet_name,
                external_workbook: prefix.external_workbook,
            },
            None => RefTarget {
                sheet: ctx_sheet,
                external_workbook: false,
            },
        }
    }

    fn to_formula_string(&self) -> String {
        let mut out = String::new();
        if let Some(prefix) = &self.prefix {
            out.push_str(&prefix.raw);
        }
        match &self.kind {
            RefKind::Cell(cell) => out.push_str(&fmt_cell(*cell)),
            RefKind::CellRange { start, end } => {
                out.push_str(&fmt_cell(*start));
                out.push(':');
                out.push_str(&fmt_cell(*end));
            }
            RefKind::ColRange { start, end } => {
                out.push_str(&fmt_col(*start));
                out.push(':');
                out.push_str(&fmt_col(*end));
            }
            RefKind::RowRange { start, end } => {
                out.push_str(&fmt_row(*start));
                out.push(':');
                out.push_str(&fmt_row(*end));
            }
            RefKind::RefError => out.push_str("#REF!"),
        }
        out
    }
}

#[derive(Clone, Debug, PartialEq, Eq)]
enum FormulaPart {
    Raw(String),
    Ref(ParsedRef),
}

#[derive(Clone, Debug, PartialEq, Eq)]
enum RefRewrite {
    Kind(RefKind),
    Text(String),
}

pub(crate) enum StructuralEdit {
    InsertRows { sheet: String, row: u32, count: u32 },
    DeleteRows { sheet: String, row: u32, count: u32 },
    InsertCols { sheet: String, col: u32, count: u32 },
    DeleteCols { sheet: String, col: u32, count: u32 },
}

pub(crate) struct RangeMapEdit {
    pub sheet: String,
    pub moved_region: GridRange,
    pub delta_row: i32,
    pub delta_col: i32,
    pub deleted_region: Option<GridRange>,
}

pub(crate) fn rewrite_formula_for_structural_edit(
    formula: &str,
    ctx_sheet: &str,
    edit: &StructuralEdit,
) -> (String, bool) {
    rewrite_formula(formula, ctx_sheet, |parsed, target| {
        if target.external_workbook {
            return None;
        }
        let matches = match edit {
            StructuralEdit::InsertRows { sheet, .. }
            | StructuralEdit::DeleteRows { sheet, .. }
            | StructuralEdit::InsertCols { sheet, .. }
            | StructuralEdit::DeleteCols { sheet, .. } => target.sheet.eq_ignore_ascii_case(sheet),
        };
        if !matches {
            return None;
        }

        let new_kind = match edit {
            StructuralEdit::InsertRows { row, count, .. } => {
                adjust_ref_kind_insert_rows(&parsed.kind, *row, *count)
            }
            StructuralEdit::DeleteRows { row, count, .. } => {
                adjust_ref_kind_delete_rows(&parsed.kind, *row, *count)
            }
            StructuralEdit::InsertCols { col, count, .. } => {
                adjust_ref_kind_insert_cols(&parsed.kind, *col, *count)
            }
            StructuralEdit::DeleteCols { col, count, .. } => {
                adjust_ref_kind_delete_cols(&parsed.kind, *col, *count)
            }
        };
        Some(RefRewrite::Kind(new_kind))
    })
}

pub(crate) fn rewrite_formula_for_copy_delta(
    formula: &str,
    ctx_sheet: &str,
    delta_row: i32,
    delta_col: i32,
) -> (String, bool) {
    rewrite_formula(formula, ctx_sheet, |parsed, target| {
        if target.external_workbook {
            return None;
        }
        Some(RefRewrite::Kind(adjust_ref_kind_copy_delta(
            &parsed.kind,
            delta_row,
            delta_col,
        )))
    })
}

pub(crate) fn rewrite_formula_for_range_map(
    formula: &str,
    ctx_sheet: &str,
    edit: &RangeMapEdit,
) -> (String, bool) {
    rewrite_formula(formula, ctx_sheet, |parsed, target| {
        if target.external_workbook || !target.sheet.eq_ignore_ascii_case(&edit.sheet) {
            return None;
        }
        Some(map_reference_via_range_map(parsed, edit))
    })
}

fn rewrite_formula<F>(formula: &str, ctx_sheet: &str, mut f: F) -> (String, bool)
where
    F: FnMut(&ParsedRef, RefTarget<'_>) -> Option<RefRewrite>,
{
    let parts = tokenize_formula(formula);
    let mut changed = false;
    let mut out = String::new();
    for part in parts {
        match part {
            FormulaPart::Raw(text) => out.push_str(&text),
            FormulaPart::Ref(mut parsed) => {
                let original = parsed.to_formula_string();
                let target = parsed.target_sheet(ctx_sheet);
                if let Some(rewrite) = f(&parsed, target) {
                    match rewrite {
                        RefRewrite::Kind(new_kind) => {
                            if parsed.kind != new_kind {
                                changed = true;
                                parsed.kind = new_kind;
                            }
                            out.push_str(&parsed.to_formula_string());
                        }
                        RefRewrite::Text(text) => {
                            if text != original {
                                changed = true;
                            }
                            out.push_str(&text);
                        }
                    }
                } else {
                    out.push_str(&original);
                }
            }
        }
    }
    (out, changed)
}

fn tokenize_formula(formula: &str) -> Vec<FormulaPart> {
    let mut parts = Vec::new();
    let mut raw_buf = String::new();
    let mut i = 0usize;
    while i < formula.len() {
        let ch = formula[i..].chars().next().unwrap();
        if ch == '"' {
            // String literal - preserve exactly and skip reference parsing inside.
            let start = i;
            i += 1;
            while i < formula.len() {
                let c = formula[i..].chars().next().unwrap();
                i += c.len_utf8();
                if c == '"' {
                    // Escaped quote is doubled.
                    if i < formula.len() && formula[i..].starts_with('"') {
                        i += 1;
                        continue;
                    }
                    break;
                }
            }
            raw_buf.push_str(&formula[start..i]);
            continue;
        }

        if ch == '[' {
            if let Some((parsed, consumed)) = parse_reference_at(formula, i) {
                if !raw_buf.is_empty() {
                    parts.push(FormulaPart::Raw(std::mem::take(&mut raw_buf)));
                }
                parts.push(FormulaPart::Ref(parsed));
                i += consumed;
                continue;
            }
            // Assume structured reference like Table1[Column] or [@Column].
            if let Some(end) = formula[i..].find(']') {
                raw_buf.push_str(&formula[i..i + end + 1]);
                i += end + 1;
                continue;
            }
        }

        if let Some((parsed, consumed)) = parse_reference_at(formula, i) {
            if !raw_buf.is_empty() {
                parts.push(FormulaPart::Raw(std::mem::take(&mut raw_buf)));
            }
            parts.push(FormulaPart::Ref(parsed));
            i += consumed;
            continue;
        }

        raw_buf.push(ch);
        i += ch.len_utf8();
    }

    if !raw_buf.is_empty() {
        parts.push(FormulaPart::Raw(raw_buf));
    }

    parts
}

fn parse_reference_at(formula: &str, start: usize) -> Option<(ParsedRef, usize)> {
    if !is_boundary_before(formula, start) {
        return None;
    }

    let (prefix, idx) = if let Some((prefix, next)) = parse_sheet_prefix(formula, start) {
        (Some(prefix), next)
    } else {
        (None, start)
    };

    let (kind, end_idx) = parse_ref_body(formula, idx)?;

    if !is_boundary_after(formula, end_idx) {
        return None;
    }

    // Avoid mis-parsing functions like LOG10( ... ) as a cell reference.
    if matches!(kind, RefKind::Cell(_) | RefKind::CellRange { .. })
        && formula[end_idx..].starts_with('(')
    {
        return None;
    }

    Some((ParsedRef { prefix, kind }, end_idx - start))
}

fn is_ident_char(ch: char) -> bool {
    ch.is_ascii_alphanumeric() || ch == '_' || ch == '.'
}

fn is_boundary_before(formula: &str, idx: usize) -> bool {
    if idx == 0 {
        return true;
    }
    let prev = formula[..idx].chars().next_back().unwrap_or('\0');
    !is_ident_char(prev)
}

fn is_boundary_after(formula: &str, idx: usize) -> bool {
    if idx >= formula.len() {
        return true;
    }
    let next = formula[idx..].chars().next().unwrap_or('\0');
    !is_ident_char(next)
}

fn parse_sheet_prefix(formula: &str, start: usize) -> Option<(SheetPrefix, usize)> {
    let bytes = formula.as_bytes();
    if start >= bytes.len() {
        return None;
    }

    match bytes[start] as char {
        '\'' => parse_quoted_sheet_prefix(formula, start),
        '[' => parse_bracketed_workbook_prefix(formula, start),
        _ => parse_unquoted_sheet_prefix(formula, start),
    }
}

fn parse_quoted_sheet_prefix(formula: &str, start: usize) -> Option<(SheetPrefix, usize)> {
    let mut i = start + 1;
    while i < formula.len() {
        let ch = formula[i..].chars().next().unwrap();
        if ch == '\'' {
            // Escaped quote inside quoted sheet name.
            if formula[i + 1..].starts_with('\'') {
                i += 2;
                continue;
            }
            // Closing quote.
            let after_quote = i + 1;
            if after_quote < formula.len() && formula[after_quote..].starts_with('!') {
                let raw = formula[start..after_quote + 1].to_string();
                let inner = formula[start + 1..i].replace("''", "'");
                let (sheet_name, external_workbook) = if inner.starts_with('[') {
                    if let Some(end) = inner.find(']') {
                        let sheet = inner[end + 1..].to_string();
                        (sheet, true)
                    } else {
                        (inner.clone(), false)
                    }
                } else {
                    (inner.clone(), false)
                };
                return Some((
                    SheetPrefix {
                        raw,
                        sheet_name,
                        external_workbook,
                    },
                    after_quote + 1,
                ));
            }
            return None;
        }
        i += ch.len_utf8();
    }
    None
}

fn parse_bracketed_workbook_prefix(formula: &str, start: usize) -> Option<(SheetPrefix, usize)> {
    let rest = &formula[start..];
    let end_bracket = rest.find(']')?;
    let after_bracket = start + end_bracket + 1;
    let bang = formula[after_bracket..].find('!')?;
    if bang == 0 {
        return None;
    }
    let bang_idx = after_bracket + bang;
    let raw = formula[start..bang_idx + 1].to_string();
    let sheet_name = formula[after_bracket..bang_idx].to_string();
    Some((
        SheetPrefix {
            raw,
            sheet_name,
            external_workbook: true,
        },
        bang_idx + 1,
    ))
}

fn parse_unquoted_sheet_prefix(formula: &str, start: usize) -> Option<(SheetPrefix, usize)> {
    let rest = &formula[start..];
    let bang = rest.find('!')?;
    if bang == 0 {
        return None;
    }
    let name = &rest[..bang];
    if name.chars().any(|ch| !is_ident_char(ch)) {
        return None;
    }
    Some((
        SheetPrefix {
            raw: rest[..bang + 1].to_string(),
            sheet_name: name.to_string(),
            external_workbook: false,
        },
        start + bang + 1,
    ))
}

fn parse_ref_body(formula: &str, start: usize) -> Option<(RefKind, usize)> {
    let rest = &formula[start..];
    if rest.len() >= 5 && rest[..5].eq_ignore_ascii_case("#REF!") {
        return Some((RefKind::RefError, start + 5));
    }

    let mut i = start;
    let ch = formula[i..].chars().next()?;

    if ch == '$' || ch.is_ascii_alphabetic() {
        let (col1, col1_abs, after_col1) = parse_col(formula, i)?;
        i = after_col1;

        let next = formula[i..].chars().next();
        match next {
            Some('$') => {
                let (row1, row1_abs, after_row1) = parse_row(formula, i)?;
                i = after_row1;
                let cell1 = A1CellRef {
                    col: col1,
                    row: row1,
                    col_abs: col1_abs,
                    row_abs: row1_abs,
                };
                if formula[i..].starts_with(':') {
                    let (cell2, after_cell2) = parse_cell(formula, i + 1)?;
                    return Some((
                        RefKind::CellRange {
                            start: cell1,
                            end: cell2,
                        },
                        after_cell2,
                    ));
                }
                return Some((RefKind::Cell(cell1), i));
            }
            Some(d) if d.is_ascii_digit() => {
                let (row1, row1_abs, after_row1) = parse_row(formula, i)?;
                i = after_row1;
                let cell1 = A1CellRef {
                    col: col1,
                    row: row1,
                    col_abs: col1_abs,
                    row_abs: row1_abs,
                };
                if formula[i..].starts_with(':') {
                    let (cell2, after_cell2) = parse_cell(formula, i + 1)?;
                    return Some((
                        RefKind::CellRange {
                            start: cell1,
                            end: cell2,
                        },
                        after_cell2,
                    ));
                }
                return Some((RefKind::Cell(cell1), i));
            }
            Some(':') => {
                let (col2, col2_abs, after_col2) = parse_col(formula, i + 1)?;
                if let Some(next) = formula[after_col2..].chars().next() {
                    if next == '$' || next.is_ascii_digit() {
                        return None;
                    }
                }
                return Some((
                    RefKind::ColRange {
                        start: A1ColRef {
                            col: col1,
                            abs: col1_abs,
                        },
                        end: A1ColRef {
                            col: col2,
                            abs: col2_abs,
                        },
                    },
                    after_col2,
                ));
            }
            _ => return None,
        }
    }

    if ch == '$' || ch.is_ascii_digit() {
        if let Some((row1, row1_abs, after_row1)) = parse_row(formula, start) {
            if formula[after_row1..].starts_with(':') {
                let (row2, row2_abs, after_row2) = parse_row(formula, after_row1 + 1)?;
                return Some((
                    RefKind::RowRange {
                        start: A1RowRef { row: row1, abs: row1_abs },
                        end: A1RowRef { row: row2, abs: row2_abs },
                    },
                    after_row2,
                ));
            }
        }
    }

    None
}

fn parse_col(formula: &str, start: usize) -> Option<(u32, bool, usize)> {
    let mut i = start;
    let mut abs = false;
    if formula[i..].starts_with('$') {
        abs = true;
        i += 1;
    }
    let mut end = i;
    while end < formula.len() {
        let ch = formula[end..].chars().next().unwrap();
        if !ch.is_ascii_alphabetic() {
            break;
        }
        end += ch.len_utf8();
        if end - i > 3 {
            return None;
        }
    }
    if end == i {
        return None;
    }
    let col = col_from_name(&formula[i..end])?;
    Some((col, abs, end))
}

fn parse_row(formula: &str, start: usize) -> Option<(u32, bool, usize)> {
    let mut i = start;
    let mut abs = false;
    if formula[i..].starts_with('$') {
        abs = true;
        i += 1;
    }
    let mut end = i;
    while end < formula.len() {
        let ch = formula[end..].chars().next().unwrap();
        if !ch.is_ascii_digit() {
            break;
        }
        end += ch.len_utf8();
    }
    if end == i {
        return None;
    }
    let row_1_based: u32 = formula[i..end].parse().ok()?;
    if row_1_based == 0 {
        return None;
    }
    Some((row_1_based - 1, abs, end))
}

fn parse_cell(formula: &str, start: usize) -> Option<(A1CellRef, usize)> {
    let (col, col_abs, after_col) = parse_col(formula, start)?;
    let (row, row_abs, after_row) = parse_row(formula, after_col)?;
    Some((
        A1CellRef {
            col,
            row,
            col_abs,
            row_abs,
        },
        after_row,
    ))
}

fn fmt_cell(cell: A1CellRef) -> String {
    let mut out = String::new();
    if cell.col_abs {
        out.push('$');
    }
    out.push_str(&col_to_name(cell.col));
    if cell.row_abs {
        out.push('$');
    }
    out.push_str(&(cell.row + 1).to_string());
    out
}

fn fmt_col(col: A1ColRef) -> String {
    let mut out = String::new();
    if col.abs {
        out.push('$');
    }
    out.push_str(&col_to_name(col.col));
    out
}

fn fmt_row(row: A1RowRef) -> String {
    let mut out = String::new();
    if row.abs {
        out.push('$');
    }
    out.push_str(&(row.row + 1).to_string());
    out
}

fn col_to_name(col: u32) -> String {
    // 0-indexed -> A1 letters.
    let mut n = col + 1;
    let mut out = Vec::<u8>::new();
    while n > 0 {
        let rem = (n - 1) % 26;
        out.push(b'A' + rem as u8);
        n = (n - 1) / 26;
    }
    out.reverse();
    String::from_utf8(out).expect("column letters are ASCII")
}

fn col_from_name(name: &str) -> Option<u32> {
    let mut col: u32 = 0;
    let mut len = 0usize;
    for ch in name.chars() {
        if !ch.is_ascii_alphabetic() {
            return None;
        }
        let v = ch.to_ascii_uppercase() as u32 - 'A' as u32 + 1;
        col = col.checked_mul(26)?.checked_add(v)?;
        len += 1;
        if len > 3 {
            return None;
        }
    }
    if len == 0 {
        None
    } else {
        Some(col - 1)
    }
}

fn adjust_ref_kind_insert_rows(kind: &RefKind, at: u32, count: u32) -> RefKind {
    match kind {
        RefKind::Cell(cell) => RefKind::Cell(A1CellRef {
            row: adjust_row_insert(cell.row, at, count),
            ..*cell
        }),
        RefKind::CellRange { start, end } => RefKind::CellRange {
            start: A1CellRef {
                row: adjust_row_insert(start.row, at, count),
                ..*start
            },
            end: A1CellRef {
                row: adjust_row_insert_range_end(end.row, start.row, at, count),
                ..*end
            },
        },
        RefKind::RowRange { start, end } => RefKind::RowRange {
            start: A1RowRef {
                row: adjust_row_insert(start.row, at, count),
                ..*start
            },
            end: A1RowRef {
                row: adjust_row_insert_range_end(end.row, start.row, at, count),
                ..*end
            },
        },
        _ => kind.clone(),
    }
}

fn adjust_ref_kind_insert_cols(kind: &RefKind, at: u32, count: u32) -> RefKind {
    match kind {
        RefKind::Cell(cell) => RefKind::Cell(A1CellRef {
            col: adjust_col_insert(cell.col, at, count),
            ..*cell
        }),
        RefKind::CellRange { start, end } => RefKind::CellRange {
            start: A1CellRef {
                col: adjust_col_insert(start.col, at, count),
                ..*start
            },
            end: A1CellRef {
                col: adjust_col_insert_range_end(end.col, start.col, at, count),
                ..*end
            },
        },
        RefKind::ColRange { start, end } => RefKind::ColRange {
            start: A1ColRef {
                col: adjust_col_insert(start.col, at, count),
                ..*start
            },
            end: A1ColRef {
                col: adjust_col_insert_range_end(end.col, start.col, at, count),
                ..*end
            },
        },
        _ => kind.clone(),
    }
}

fn adjust_ref_kind_delete_rows(kind: &RefKind, at: u32, count: u32) -> RefKind {
    let del_end = at.saturating_add(count.saturating_sub(1));
    match kind {
        RefKind::Cell(cell) => match adjust_row_delete(cell.row, at, del_end, count) {
            Some(new_row) => RefKind::Cell(A1CellRef { row: new_row, ..*cell }),
            None => RefKind::RefError,
        },
        RefKind::CellRange { start, end } => {
            let Some((new_start, new_end)) =
                adjust_row_range_delete(start.row, end.row, at, del_end, count)
            else {
                return RefKind::RefError;
            };
            RefKind::CellRange {
                start: A1CellRef { row: new_start, ..*start },
                end: A1CellRef { row: new_end, ..*end },
            }
        }
        RefKind::RowRange { start, end } => {
            let Some((new_start, new_end)) =
                adjust_row_range_delete(start.row, end.row, at, del_end, count)
            else {
                return RefKind::RefError;
            };
            RefKind::RowRange {
                start: A1RowRef { row: new_start, ..*start },
                end: A1RowRef { row: new_end, ..*end },
            }
        }
        _ => kind.clone(),
    }
}

fn adjust_ref_kind_delete_cols(kind: &RefKind, at: u32, count: u32) -> RefKind {
    let del_end = at.saturating_add(count.saturating_sub(1));
    match kind {
        RefKind::Cell(cell) => match adjust_col_delete(cell.col, at, del_end, count) {
            Some(new_col) => RefKind::Cell(A1CellRef { col: new_col, ..*cell }),
            None => RefKind::RefError,
        },
        RefKind::CellRange { start, end } => {
            let Some((new_start, new_end)) =
                adjust_col_range_delete(start.col, end.col, at, del_end, count)
            else {
                return RefKind::RefError;
            };
            RefKind::CellRange {
                start: A1CellRef { col: new_start, ..*start },
                end: A1CellRef { col: new_end, ..*end },
            }
        }
        RefKind::ColRange { start, end } => {
            let Some((new_start, new_end)) =
                adjust_col_range_delete(start.col, end.col, at, del_end, count)
            else {
                return RefKind::RefError;
            };
            RefKind::ColRange {
                start: A1ColRef { col: new_start, ..*start },
                end: A1ColRef { col: new_end, ..*end },
            }
        }
        _ => kind.clone(),
    }
}

fn adjust_ref_kind_copy_delta(kind: &RefKind, delta_row: i32, delta_col: i32) -> RefKind {
    match kind {
        RefKind::Cell(cell) => apply_delta_cell(*cell, delta_row, delta_col)
            .map(RefKind::Cell)
            .unwrap_or(RefKind::RefError),
        RefKind::CellRange { start, end } => {
            let Some(start) = apply_delta_cell(*start, delta_row, delta_col) else {
                return RefKind::RefError;
            };
            let Some(end) = apply_delta_cell(*end, delta_row, delta_col) else {
                return RefKind::RefError;
            };
            RefKind::CellRange { start, end }
        }
        RefKind::ColRange { start, end } => {
            let start_col = apply_delta_col(*start, delta_col);
            let end_col = apply_delta_col(*end, delta_col);
            match (start_col, end_col) {
                (Some(sc), Some(ec)) => RefKind::ColRange {
                    start: A1ColRef { col: sc, ..*start },
                    end: A1ColRef { col: ec, ..*end },
                },
                _ => RefKind::RefError,
            }
        }
        RefKind::RowRange { start, end } => {
            let start_row = apply_delta_row(*start, delta_row);
            let end_row = apply_delta_row(*end, delta_row);
            match (start_row, end_row) {
                (Some(sr), Some(er)) => RefKind::RowRange {
                    start: A1RowRef { row: sr, ..*start },
                    end: A1RowRef { row: er, ..*end },
                },
                _ => RefKind::RefError,
            }
        }
        RefKind::RefError => RefKind::RefError,
    }
}

fn apply_delta_cell(cell: A1CellRef, delta_row: i32, delta_col: i32) -> Option<A1CellRef> {
    let new_row = if cell.row_abs {
        cell.row as i32
    } else {
        cell.row as i32 + delta_row
    };
    let new_col = if cell.col_abs {
        cell.col as i32
    } else {
        cell.col as i32 + delta_col
    };
    if new_row < 0 || new_col < 0 {
        return None;
    }
    Some(A1CellRef {
        row: new_row as u32,
        col: new_col as u32,
        ..cell
    })
}

fn apply_delta_col(col: A1ColRef, delta_col: i32) -> Option<u32> {
    if col.abs {
        Some(col.col)
    } else {
        let new = col.col as i32 + delta_col;
        if new < 0 {
            None
        } else {
            Some(new as u32)
        }
    }
}

fn apply_delta_row(row: A1RowRef, delta_row: i32) -> Option<u32> {
    if row.abs {
        Some(row.row)
    } else {
        let new = row.row as i32 + delta_row;
        if new < 0 {
            None
        } else {
            Some(new as u32)
        }
    }
}

fn adjust_row_insert(row: u32, at: u32, count: u32) -> u32 {
    if row >= at {
        row + count
    } else {
        row
    }
}

fn adjust_row_insert_range_end(end_row: u32, start_row: u32, at: u32, count: u32) -> u32 {
    if start_row < at && end_row >= at {
        end_row + count
    } else {
        adjust_row_insert(end_row, at, count)
    }
}

fn adjust_col_insert(col: u32, at: u32, count: u32) -> u32 {
    if col >= at {
        col + count
    } else {
        col
    }
}

fn adjust_col_insert_range_end(end_col: u32, start_col: u32, at: u32, count: u32) -> u32 {
    if start_col < at && end_col >= at {
        end_col + count
    } else {
        adjust_col_insert(end_col, at, count)
    }
}

fn adjust_row_delete(row: u32, del_start: u32, del_end: u32, count: u32) -> Option<u32> {
    if row < del_start {
        Some(row)
    } else if row > del_end {
        Some(row - count)
    } else {
        None
    }
}

fn adjust_col_delete(col: u32, del_start: u32, del_end: u32, count: u32) -> Option<u32> {
    if col < del_start {
        Some(col)
    } else if col > del_end {
        Some(col - count)
    } else {
        None
    }
}

fn adjust_row_range_delete(
    start: u32,
    end: u32,
    del_start: u32,
    del_end: u32,
    count: u32,
) -> Option<(u32, u32)> {
    if end < del_start {
        return Some((start, end));
    }
    if start > del_end {
        return Some((start - count, end - count));
    }
    if start >= del_start && end <= del_end {
        return None;
    }

    let mut new_start = start;
    let mut new_end = end;

    if start >= del_start && start <= del_end {
        new_start = del_start;
    }

    if end >= del_start && end <= del_end {
        if del_start == 0 {
            return None;
        }
        new_end = del_start - 1;
    } else if end > del_end {
        new_end = end - count;
    }

    if new_start > new_end {
        None
    } else {
        Some((new_start, new_end))
    }
}

fn adjust_col_range_delete(
    start: u32,
    end: u32,
    del_start: u32,
    del_end: u32,
    count: u32,
) -> Option<(u32, u32)> {
    if end < del_start {
        return Some((start, end));
    }
    if start > del_end {
        return Some((start - count, end - count));
    }
    if start >= del_start && end <= del_end {
        return None;
    }

    let mut new_start = start;
    let mut new_end = end;

    if start >= del_start && start <= del_end {
        new_start = del_start;
    }

    if end >= del_start && end <= del_end {
        if del_start == 0 {
            return None;
        }
        new_end = del_start - 1;
    } else if end > del_end {
        new_end = end - count;
    }

    if new_start > new_end {
        None
    } else {
        Some((new_start, new_end))
    }
}

fn map_reference_via_range_map(parsed: &ParsedRef, edit: &RangeMapEdit) -> RefRewrite {
    match &parsed.kind {
        RefKind::Cell(cell) => {
            if let Some(deleted) = edit.deleted_region {
                if deleted.contains(cell.row, cell.col) {
                    return RefRewrite::Kind(RefKind::RefError);
                }
            }
            if !edit.moved_region.contains(cell.row, cell.col) {
                return RefRewrite::Kind(parsed.kind.clone());
            }
            let new_row = cell.row as i32 + edit.delta_row;
            let new_col = cell.col as i32 + edit.delta_col;
            if new_row < 0 || new_col < 0 {
                return RefRewrite::Kind(RefKind::RefError);
            }
            RefRewrite::Kind(RefKind::Cell(A1CellRef {
                row: new_row as u32,
                col: new_col as u32,
                ..*cell
            }))
        }
        RefKind::CellRange { start, end } => {
            let original = GridRange::new(start.row, start.col, end.row, end.col);
            let mut areas = vec![original];

            if let Some(deleted) = edit.deleted_region {
                areas = subtract_region(&areas, deleted);
                if areas.is_empty() {
                    return RefRewrite::Kind(RefKind::RefError);
                }
            }

            areas = apply_move_region(&areas, edit.moved_region, edit.delta_row, edit.delta_col);
            if areas.is_empty() {
                return RefRewrite::Kind(RefKind::RefError);
            }

            let mut area_strings = Vec::new();
            for area in areas {
                let kind = if area.start_row == area.end_row && area.start_col == area.end_col {
                    RefKind::Cell(A1CellRef {
                        row: area.start_row,
                        col: area.start_col,
                        row_abs: start.row_abs,
                        col_abs: start.col_abs,
                    })
                } else {
                    RefKind::CellRange {
                        start: A1CellRef {
                            row: area.start_row,
                            col: area.start_col,
                            row_abs: start.row_abs,
                            col_abs: start.col_abs,
                        },
                        end: A1CellRef {
                            row: area.end_row,
                            col: area.end_col,
                            row_abs: end.row_abs,
                            col_abs: end.col_abs,
                        },
                    }
                };
                let r = ParsedRef {
                    prefix: parsed.prefix.clone(),
                    kind,
                };
                area_strings.push(r.to_formula_string());
            }

            if area_strings.len() == 1 {
                RefRewrite::Text(area_strings[0].clone())
            } else {
                RefRewrite::Text(format!("({})", area_strings.join(",")))
            }
        }
        _ => RefRewrite::Kind(parsed.kind.clone()),
    }
}

fn subtract_region(areas: &[GridRange], deleted: GridRange) -> Vec<GridRange> {
    let mut out = Vec::new();
    for area in areas {
        let overlap = rect_intersection(*area, deleted);
        match overlap {
            None => out.push(*area),
            Some(over) if over == *area => {}
            Some(over) => out.extend(rect_difference(*area, over)),
        }
    }
    out
}

fn apply_move_region(
    areas: &[GridRange],
    moved_region: GridRange,
    delta_row: i32,
    delta_col: i32,
) -> Vec<GridRange> {
    let mut out = Vec::new();
    for area in areas {
        let overlap = rect_intersection(*area, moved_region);
        match overlap {
            None => out.push(*area),
            Some(over) if over == *area => {
                if let Some(shifted) = shift_range(*area, delta_row, delta_col) {
                    out.push(shifted);
                }
            }
            Some(over) => {
                out.extend(rect_difference(*area, over));
                if let Some(shifted) = shift_range(over, delta_row, delta_col) {
                    out.push(shifted);
                }
            }
        }
    }
    out
}

fn rect_intersection(a: GridRange, b: GridRange) -> Option<GridRange> {
    let start_row = max(a.start_row, b.start_row);
    let start_col = max(a.start_col, b.start_col);
    let end_row = min(a.end_row, b.end_row);
    let end_col = min(a.end_col, b.end_col);
    if start_row > end_row || start_col > end_col {
        None
    } else {
        Some(GridRange::new(start_row, start_col, end_row, end_col))
    }
}

fn rect_difference(range: GridRange, overlap: GridRange) -> Vec<GridRange> {
    let mut out = Vec::new();
    if range.start_row < overlap.start_row {
        out.push(GridRange::new(
            range.start_row,
            range.start_col,
            overlap.start_row - 1,
            range.end_col,
        ));
    }
    if overlap.end_row < range.end_row {
        out.push(GridRange::new(
            overlap.end_row + 1,
            range.start_col,
            range.end_row,
            range.end_col,
        ));
    }
    let mid_start_row = overlap.start_row;
    let mid_end_row = overlap.end_row;
    if range.start_col < overlap.start_col {
        out.push(GridRange::new(
            mid_start_row,
            range.start_col,
            mid_end_row,
            overlap.start_col - 1,
        ));
    }
    if overlap.end_col < range.end_col {
        out.push(GridRange::new(
            mid_start_row,
            overlap.end_col + 1,
            mid_end_row,
            range.end_col,
        ));
    }
    out
}

fn shift_range(area: GridRange, delta_row: i32, delta_col: i32) -> Option<GridRange> {
    let start_row = area.start_row as i32 + delta_row;
    let start_col = area.start_col as i32 + delta_col;
    let end_row = area.end_row as i32 + delta_row;
    let end_col = area.end_col as i32 + delta_col;
    if start_row < 0 || start_col < 0 || end_row < 0 || end_col < 0 {
        return None;
    }
    Some(GridRange::new(
        start_row as u32,
        start_col as u32,
        end_row as u32,
        end_col as u32,
    ))
}
