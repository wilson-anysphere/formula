use serde::{Deserialize, Serialize};
use serde_json::Value as JsonValue;
use std::collections::{HashMap, HashSet};
use std::fmt;

pub const DEFAULT_SHEET: &str = "Sheet1";

#[derive(Clone, Copy, Debug, PartialEq, Eq, Hash)]
pub struct CellCoord {
    pub row: u32,
    pub col: u32,
}

impl CellCoord {
    pub fn new(row: u32, col: u32) -> Self {
        Self { row, col }
    }
}

#[derive(Clone, Debug)]
pub struct Cell {
    pub input: JsonValue,
    pub value: JsonValue,
}

impl Cell {
    fn new(input: JsonValue) -> Self {
        let value = if is_formula_input(&input) {
            JsonValue::Null
        } else {
            input.clone()
        };
        Self { input, value }
    }
}

#[derive(Clone, Debug, Serialize, Deserialize, PartialEq)]
pub struct CellData {
    pub sheet: String,
    pub address: String,
    pub input: JsonValue,
    pub value: JsonValue,
}

#[derive(Clone, Debug, Serialize, Deserialize, PartialEq)]
pub struct CellChange {
    pub sheet: String,
    pub address: String,
    pub value: JsonValue,
}

#[derive(Clone, Debug, Default)]
pub struct Workbook {
    sheets: HashMap<String, Sheet>,
}

#[derive(Clone, Debug, Default)]
struct Sheet {
    cells: HashMap<CellCoord, Cell>,
}

#[derive(Debug)]
pub enum WorkbookError {
    InvalidAddress(String),
    InvalidRange(String),
    InvalidCellValue(String),
    InvalidFormula(String),
    MissingSheet(String),
    InvalidJson(String),
}

impl fmt::Display for WorkbookError {
    fn fmt(&self, f: &mut fmt::Formatter<'_>) -> fmt::Result {
        match self {
            WorkbookError::InvalidAddress(addr) => write!(f, "invalid cell address: {addr}"),
            WorkbookError::InvalidRange(r) => write!(f, "invalid range: {r}"),
            WorkbookError::InvalidCellValue(v) => write!(f, "invalid cell value: {v}"),
            WorkbookError::InvalidFormula(expr) => write!(f, "invalid formula: {expr}"),
            WorkbookError::MissingSheet(name) => write!(f, "missing sheet: {name}"),
            WorkbookError::InvalidJson(err) => write!(f, "invalid workbook json: {err}"),
        }
    }
}

impl std::error::Error for WorkbookError {}

impl Workbook {
    pub fn new() -> Self {
        let mut wb = Self::default();
        wb.sheets
            .insert(DEFAULT_SHEET.to_string(), Sheet::default());
        wb
    }

    pub fn from_json_str(json: &str) -> Result<Self, WorkbookError> {
        #[derive(Debug, Serialize, Deserialize)]
        struct WorkbookJson {
            sheets: HashMap<String, SheetJson>,
        }

        #[derive(Debug, Serialize, Deserialize)]
        struct SheetJson {
            cells: HashMap<String, JsonValue>,
        }

        let parsed: WorkbookJson = serde_json::from_str(json)
            .map_err(|err| WorkbookError::InvalidJson(err.to_string()))?;

        let mut wb = Workbook::default();
        for (sheet_name, sheet_json) in parsed.sheets {
            let mut sheet = Sheet::default();
            for (address, input) in sheet_json.cells {
                if !is_scalar_json(&input) {
                    return Err(WorkbookError::InvalidCellValue(address));
                }
                let coord =
                    parse_a1(&address).map_err(|_| WorkbookError::InvalidAddress(address))?;
                sheet.cells.insert(coord, Cell::new(input));
            }
            wb.sheets.insert(sheet_name, sheet);
        }

        if wb.sheets.is_empty() {
            wb.sheets
                .insert(DEFAULT_SHEET.to_string(), Sheet::default());
        }

        Ok(wb)
    }

    pub fn to_json_str(&self) -> Result<String, WorkbookError> {
        #[derive(Debug, Serialize)]
        struct WorkbookJson<'a> {
            sheets: HashMap<&'a str, SheetJson<'a>>,
        }

        #[derive(Debug, Serialize)]
        struct SheetJson<'a> {
            cells: HashMap<String, &'a JsonValue>,
        }

        let mut sheets = HashMap::new();
        for (sheet_name, sheet) in &self.sheets {
            let mut cells = HashMap::new();
            for (coord, cell) in &sheet.cells {
                cells.insert(format_a1(*coord), &cell.input);
            }
            sheets.insert(sheet_name.as_str(), SheetJson { cells });
        }

        serde_json::to_string(&WorkbookJson { sheets })
            .map_err(|err| WorkbookError::InvalidJson(err.to_string()))
    }

    pub fn get_cell(&self, address: &str, sheet: Option<&str>) -> Result<CellData, WorkbookError> {
        let sheet_name = sheet.unwrap_or(DEFAULT_SHEET);
        let coord = parse_a1(address).map_err(|_| WorkbookError::InvalidAddress(address.into()))?;
        let sheet = self
            .sheets
            .get(sheet_name)
            .ok_or_else(|| WorkbookError::MissingSheet(sheet_name.to_string()))?;

        let (input, value) = match sheet.cells.get(&coord) {
            Some(cell) => (cell.input.clone(), cell.value.clone()),
            None => (JsonValue::Null, JsonValue::Null),
        };

        Ok(CellData {
            sheet: sheet_name.to_string(),
            address: format_a1(coord),
            input,
            value,
        })
    }

    pub fn set_cell(
        &mut self,
        address: &str,
        input: JsonValue,
        sheet: Option<&str>,
    ) -> Result<(), WorkbookError> {
        if !is_scalar_json(&input) {
            return Err(WorkbookError::InvalidCellValue(address.to_string()));
        }
        let sheet_name = sheet.unwrap_or(DEFAULT_SHEET);
        let coord = parse_a1(address).map_err(|_| WorkbookError::InvalidAddress(address.into()))?;
        let sheet = self
            .sheets
            .entry(sheet_name.to_string())
            .or_insert_with(Sheet::default);
        sheet.cells.insert(coord, Cell::new(input));
        Ok(())
    }

    pub fn get_range(
        &self,
        range: &str,
        sheet: Option<&str>,
    ) -> Result<Vec<Vec<CellData>>, WorkbookError> {
        let sheet_name = sheet.unwrap_or(DEFAULT_SHEET);
        let (start, end) =
            parse_range(range).map_err(|_| WorkbookError::InvalidRange(range.into()))?;

        let mut rows = Vec::new();
        for row in start.row..=end.row {
            let mut cols = Vec::new();
            for col in start.col..=end.col {
                let address = format_a1(CellCoord::new(row, col));
                cols.push(self.get_cell(&address, Some(sheet_name))?);
            }
            rows.push(cols);
        }
        Ok(rows)
    }

    pub fn set_range(
        &mut self,
        range: &str,
        values: Vec<Vec<JsonValue>>,
        sheet: Option<&str>,
    ) -> Result<(), WorkbookError> {
        let sheet_name = sheet.unwrap_or(DEFAULT_SHEET);
        let (start, end) =
            parse_range(range).map_err(|_| WorkbookError::InvalidRange(range.into()))?;

        let expected_rows = (end.row - start.row + 1) as usize;
        let expected_cols = (end.col - start.col + 1) as usize;
        if values.len() != expected_rows || values.iter().any(|row| row.len() != expected_cols) {
            return Err(WorkbookError::InvalidRange(format!(
                "range {range} expects {expected_rows}x{expected_cols} values"
            )));
        }

        for (row_idx, row_values) in values.into_iter().enumerate() {
            for (col_idx, input) in row_values.into_iter().enumerate() {
                if !is_scalar_json(&input) {
                    let coord =
                        CellCoord::new(start.row + row_idx as u32, start.col + col_idx as u32);
                    return Err(WorkbookError::InvalidCellValue(format_a1(coord)));
                }
                let coord = CellCoord::new(start.row + row_idx as u32, start.col + col_idx as u32);
                let address = format_a1(coord);
                self.set_cell(&address, input, Some(sheet_name))?;
            }
        }
        Ok(())
    }

    pub fn recalculate(&mut self, sheet: Option<&str>) -> Result<Vec<CellChange>, WorkbookError> {
        let sheet_name = sheet.unwrap_or(DEFAULT_SHEET);
        let sheet = self
            .sheets
            .get(sheet_name)
            .ok_or_else(|| WorkbookError::MissingSheet(sheet_name.to_string()))?
            .clone();

        let formula_coords: Vec<CellCoord> = sheet
            .cells
            .iter()
            .filter_map(|(coord, cell)| is_formula_input(&cell.input).then_some(*coord))
            .collect();

        let mut ctx = EvalContext::default();
        let mut changes = Vec::new();

        for coord in &formula_coords {
            let address = format_a1(*coord);
            let new_value = self.eval_cell_value(sheet_name, *coord, &mut ctx)?;
            let old_value = self
                .sheets
                .get(sheet_name)
                .and_then(|s| s.cells.get(coord))
                .map(|c| c.value.clone())
                .unwrap_or(JsonValue::Null);

            if old_value != new_value {
                if let Some(sheet_mut) = self.sheets.get_mut(sheet_name) {
                    if let Some(cell_mut) = sheet_mut.cells.get_mut(coord) {
                        cell_mut.value = new_value.clone();
                    }
                }
                changes.push(CellChange {
                    sheet: sheet_name.to_string(),
                    address,
                    value: new_value,
                });
            }
        }

        Ok(changes)
    }

    fn eval_cell_value(
        &self,
        sheet_name: &str,
        coord: CellCoord,
        ctx: &mut EvalContext,
    ) -> Result<JsonValue, WorkbookError> {
        if let Some(value) = ctx.cache.get(&coord) {
            return Ok(value.clone());
        }

        if !ctx.visiting.insert(coord) {
            return Ok(JsonValue::String("#CYCLE!".to_string()));
        }

        let cell = self
            .sheets
            .get(sheet_name)
            .and_then(|sheet| sheet.cells.get(&coord));

        let value = match cell {
            None => JsonValue::Null,
            Some(cell) => {
                if let Some(formula) = cell.input.as_str().and_then(|s| s.strip_prefix('=')) {
                    let num = eval_numeric_expr(formula, |addr| {
                        let dep = parse_a1(addr)
                            .map_err(|_| WorkbookError::InvalidAddress(addr.to_string()))?;
                        self.eval_cell_as_number(sheet_name, dep, ctx)
                    })?;
                    JsonValue::Number(
                        serde_json::Number::from_f64(num)
                            .ok_or_else(|| WorkbookError::InvalidFormula(formula.to_string()))?,
                    )
                } else {
                    cell.value.clone()
                }
            }
        };

        ctx.visiting.remove(&coord);
        ctx.cache.insert(coord, value.clone());
        Ok(value)
    }

    fn eval_cell_as_number(
        &self,
        sheet_name: &str,
        coord: CellCoord,
        ctx: &mut EvalContext,
    ) -> Result<f64, WorkbookError> {
        let value = self.eval_cell_value(sheet_name, coord, ctx)?;
        json_as_number(&value).ok_or_else(|| {
            WorkbookError::InvalidFormula(format!("{} is not a number", format_a1(coord)))
        })
    }
}

#[derive(Default)]
struct EvalContext {
    visiting: HashSet<CellCoord>,
    cache: HashMap<CellCoord, JsonValue>,
}

fn is_formula_input(value: &JsonValue) -> bool {
    value
        .as_str()
        .is_some_and(|s| s.starts_with('=') && s.len() > 1)
}

fn json_as_number(value: &JsonValue) -> Option<f64> {
    match value {
        JsonValue::Null => Some(0.0),
        JsonValue::Bool(b) => Some(if *b { 1.0 } else { 0.0 }),
        JsonValue::Number(num) => num.as_f64(),
        JsonValue::String(s) => s.trim().parse::<f64>().ok(),
        _ => None,
    }
}

fn is_scalar_json(value: &JsonValue) -> bool {
    matches!(
        value,
        JsonValue::Null | JsonValue::Bool(_) | JsonValue::Number(_) | JsonValue::String(_)
    )
}

pub fn parse_a1(address: &str) -> Result<CellCoord, ()> {
    let mut chars = address.trim().chars().peekable();
    let mut col: u32 = 0;
    let mut saw_letter = false;
    while let Some(ch) = chars.peek().copied() {
        if ch.is_ascii_alphabetic() {
            saw_letter = true;
            let upper = ch.to_ascii_uppercase();
            col = col * 26 + (upper as u32 - 'A' as u32 + 1);
            chars.next();
        } else {
            break;
        }
    }

    if !saw_letter {
        return Err(());
    }

    let mut row_str = String::new();
    while let Some(ch) = chars.peek().copied() {
        if ch.is_ascii_digit() {
            row_str.push(ch);
            chars.next();
        } else {
            break;
        }
    }

    if row_str.is_empty() || chars.next().is_some() {
        return Err(());
    }

    let row: u32 = row_str.parse().map_err(|_| ())?;
    if row == 0 || col == 0 {
        return Err(());
    }
    Ok(CellCoord::new(row, col))
}

pub fn format_a1(coord: CellCoord) -> String {
    let mut col = coord.col;
    let mut letters = Vec::new();
    while col > 0 {
        let rem = (col - 1) % 26;
        letters.push((b'A' + rem as u8) as char);
        col = (col - 1) / 26;
    }
    letters.reverse();
    format!("{}{}", letters.into_iter().collect::<String>(), coord.row)
}

pub fn parse_range(range: &str) -> Result<(CellCoord, CellCoord), ()> {
    let parts: Vec<&str> = range.trim().split(':').collect();
    if parts.is_empty() || parts.len() > 2 {
        return Err(());
    }
    let start = parse_a1(parts[0])?;
    let end = if parts.len() == 2 {
        parse_a1(parts[1])?
    } else {
        start
    };

    let top = start.row.min(end.row);
    let left = start.col.min(end.col);
    let bottom = start.row.max(end.row);
    let right = start.col.max(end.col);
    Ok((CellCoord::new(top, left), CellCoord::new(bottom, right)))
}

#[derive(Clone, Copy, Debug, PartialEq, Eq)]
enum Op {
    Add,
    Sub,
    Mul,
    Div,
    Neg,
}

impl Op {
    fn precedence(self) -> u8 {
        match self {
            Op::Neg => 3,
            Op::Mul | Op::Div => 2,
            Op::Add | Op::Sub => 1,
        }
    }

    fn right_assoc(self) -> bool {
        matches!(self, Op::Neg)
    }

    fn arity(self) -> usize {
        match self {
            Op::Neg => 1,
            _ => 2,
        }
    }
}

#[derive(Clone, Debug, PartialEq)]
enum Token {
    Number(f64),
    CellRef(String),
    Op(Op),
    LParen,
    RParen,
}

fn eval_numeric_expr<F>(expr: &str, mut resolve_cell: F) -> Result<f64, WorkbookError>
where
    F: FnMut(&str) -> Result<f64, WorkbookError>,
{
    let tokens = tokenize(expr)?;
    let rpn = to_rpn(&tokens)?;
    let mut stack: Vec<f64> = Vec::new();

    for token in rpn {
        match token {
            Token::Number(num) => stack.push(num),
            Token::CellRef(addr) => stack.push(resolve_cell(&addr)?),
            Token::Op(op) => match op.arity() {
                1 => {
                    let a = stack.pop().ok_or_else(|| {
                        WorkbookError::InvalidFormula("missing operand".to_string())
                    })?;
                    stack.push(-a);
                }
                2 => {
                    let b = stack.pop().ok_or_else(|| {
                        WorkbookError::InvalidFormula("missing operand".to_string())
                    })?;
                    let a = stack.pop().ok_or_else(|| {
                        WorkbookError::InvalidFormula("missing operand".to_string())
                    })?;
                    let result = match op {
                        Op::Add => a + b,
                        Op::Sub => a - b,
                        Op::Mul => a * b,
                        Op::Div => a / b,
                        Op::Neg => unreachable!("handled by arity"),
                    };
                    stack.push(result);
                }
                _ => unreachable!("op arity is 1 or 2"),
            },
            Token::LParen | Token::RParen => {
                return Err(WorkbookError::InvalidFormula(
                    "unexpected paren in rpn".to_string(),
                ))
            }
        }
    }

    if stack.len() != 1 {
        return Err(WorkbookError::InvalidFormula(
            "invalid expression".to_string(),
        ));
    }

    Ok(stack[0])
}

fn tokenize(expr: &str) -> Result<Vec<Token>, WorkbookError> {
    let mut tokens = Vec::new();
    let mut chars = expr.chars().peekable();
    let mut prev_was_value = false;

    while let Some(ch) = chars.peek().copied() {
        if ch.is_whitespace() {
            chars.next();
            continue;
        }

        if ch.is_ascii_digit() || ch == '.' {
            let mut buf = String::new();
            while let Some(ch2) = chars.peek().copied() {
                if ch2.is_ascii_digit() || ch2 == '.' {
                    buf.push(ch2);
                    chars.next();
                } else {
                    break;
                }
            }
            let number: f64 = buf.parse().map_err(|_| {
                WorkbookError::InvalidFormula(format!("invalid number literal: {buf}"))
            })?;
            tokens.push(Token::Number(number));
            prev_was_value = true;
            continue;
        }

        if ch.is_ascii_alphabetic() {
            let mut buf = String::new();
            while let Some(ch2) = chars.peek().copied() {
                if ch2.is_ascii_alphanumeric() {
                    buf.push(ch2.to_ascii_uppercase());
                    chars.next();
                } else {
                    break;
                }
            }
            if parse_a1(&buf).is_err() {
                return Err(WorkbookError::InvalidFormula(format!(
                    "invalid cell reference: {buf}"
                )));
            }
            tokens.push(Token::CellRef(buf));
            prev_was_value = true;
            continue;
        }

        match ch {
            '+' => {
                chars.next();
                tokens.push(Token::Op(Op::Add));
                prev_was_value = false;
            }
            '-' => {
                chars.next();
                let op = if prev_was_value { Op::Sub } else { Op::Neg };
                tokens.push(Token::Op(op));
                prev_was_value = false;
            }
            '*' => {
                chars.next();
                tokens.push(Token::Op(Op::Mul));
                prev_was_value = false;
            }
            '/' => {
                chars.next();
                tokens.push(Token::Op(Op::Div));
                prev_was_value = false;
            }
            '(' => {
                chars.next();
                tokens.push(Token::LParen);
                prev_was_value = false;
            }
            ')' => {
                chars.next();
                tokens.push(Token::RParen);
                prev_was_value = true;
            }
            _ => {
                return Err(WorkbookError::InvalidFormula(format!(
                    "unexpected character: {ch}"
                )))
            }
        }
    }

    Ok(tokens)
}

fn to_rpn(tokens: &[Token]) -> Result<Vec<Token>, WorkbookError> {
    let mut output = Vec::new();
    let mut ops: Vec<Token> = Vec::new();

    for token in tokens {
        match token {
            Token::Number(_) | Token::CellRef(_) => output.push(token.clone()),
            Token::Op(op) => {
                while let Some(top) = ops.last() {
                    match top {
                        Token::Op(top_op) => {
                            let should_pop = if op.right_assoc() {
                                op.precedence() < top_op.precedence()
                            } else {
                                op.precedence() <= top_op.precedence()
                            };
                            if should_pop {
                                output.push(ops.pop().expect("exists"));
                                continue;
                            }
                        }
                        Token::LParen => {}
                        _ => {}
                    }
                    break;
                }
                ops.push(Token::Op(*op));
            }
            Token::LParen => ops.push(Token::LParen),
            Token::RParen => {
                let mut found = false;
                while let Some(top) = ops.pop() {
                    if matches!(top, Token::LParen) {
                        found = true;
                        break;
                    }
                    output.push(top);
                }
                if !found {
                    return Err(WorkbookError::InvalidFormula(
                        "mismatched parentheses".to_string(),
                    ));
                }
            }
        }
    }

    while let Some(top) = ops.pop() {
        if matches!(top, Token::LParen | Token::RParen) {
            return Err(WorkbookError::InvalidFormula(
                "mismatched parentheses".to_string(),
            ));
        }
        output.push(top);
    }

    Ok(output)
}

#[cfg(test)]
mod tests {
    use super::*;
    use pretty_assertions::assert_eq;
    use serde_json::json;

    #[test]
    fn parse_and_format_a1() {
        assert_eq!(parse_a1("A1"), Ok(CellCoord::new(1, 1)));
        assert_eq!(parse_a1("AA12"), Ok(CellCoord::new(12, 27)));
        assert_eq!(format_a1(CellCoord::new(12, 27)), "AA12");
    }

    #[test]
    fn recalculate_updates_formula_cells() {
        let mut wb = Workbook::new();
        wb.set_cell("A1", json!(1), None).unwrap();
        wb.set_cell("A2", json!("=A1*2"), None).unwrap();

        let changes = wb.recalculate(None).unwrap();
        assert_eq!(
            changes,
            vec![CellChange {
                sheet: DEFAULT_SHEET.to_string(),
                address: "A2".to_string(),
                value: json!(2.0)
            }]
        );

        let cell = wb.get_cell("A2", None).unwrap();
        assert_eq!(cell.value, json!(2.0));
    }
}
