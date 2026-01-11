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
                    eval_formula(formula, |dep| self.eval_cell_value(sheet_name, dep, ctx))?
                } else {
                    cell.value.clone()
                }
            }
        };

        ctx.visiting.remove(&coord);
        ctx.cache.insert(coord, value.clone());
        Ok(value)
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

fn is_scalar_json(value: &JsonValue) -> bool {
    matches!(
        value,
        JsonValue::Null | JsonValue::Bool(_) | JsonValue::Number(_) | JsonValue::String(_)
    )
}

pub fn parse_a1(address: &str) -> Result<CellCoord, ()> {
    let mut chars = address.trim().chars().peekable();
    if chars.peek().is_some_and(|ch| *ch == '$') {
        chars.next();
    }
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

    if chars.peek().is_some_and(|ch| *ch == '$') {
        chars.next();
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

const ERROR_DIV0: &str = "#DIV/0!";
const ERROR_NAME: &str = "#NAME?";
const ERROR_NA: &str = "#N/A";
const ERROR_REF: &str = "#REF!";
const ERROR_VALUE: &str = "#VALUE!";
const ERROR_NUM: &str = "#NUM!";

#[derive(Clone, Debug, PartialEq)]
enum Expr {
    Empty,
    Number(f64),
    Scalar(JsonValue),
    Reference(CellCoord),
    Range(CellCoord, CellCoord),
    Unary {
        op: UnaryOp,
        rhs: Box<Expr>,
    },
    Binary {
        op: BinaryOp,
        lhs: Box<Expr>,
        rhs: Box<Expr>,
    },
    FunctionCall {
        name: String,
        args: Vec<Expr>,
    },
    Error(&'static str),
}

#[derive(Clone, Copy, Debug, PartialEq, Eq)]
enum UnaryOp {
    Plus,
    Minus,
}

#[derive(Clone, Copy, Debug, PartialEq, Eq)]
enum BinaryOp {
    Add,
    Sub,
    Mul,
    Div,
}

#[derive(Clone, Debug, PartialEq)]
enum FormulaToken {
    Number(f64),
    String(String),
    ErrorCode(String),
    Ident(String),
    Plus,
    Minus,
    Star,
    Slash,
    OtherOp(String),
    LParen,
    RParen,
    Comma,
    Colon,
}

struct FormulaParser {
    tokens: Vec<FormulaToken>,
    pos: usize,
}

impl FormulaParser {
    fn new(tokens: Vec<FormulaToken>) -> Self {
        Self { tokens, pos: 0 }
    }

    fn peek(&self) -> Option<&FormulaToken> {
        self.tokens.get(self.pos)
    }

    fn bump(&mut self) -> Option<FormulaToken> {
        let tok = self.tokens.get(self.pos).cloned();
        if tok.is_some() {
            self.pos += 1;
        }
        tok
    }

    fn parse_expression(&mut self) -> Expr {
        let mut expr = self.parse_term();
        loop {
            let op = match self.peek() {
                Some(FormulaToken::Plus) => BinaryOp::Add,
                Some(FormulaToken::Minus) => BinaryOp::Sub,
                _ => break,
            };
            self.bump();
            let rhs = self.parse_term();
            expr = Expr::Binary {
                op,
                lhs: Box::new(expr),
                rhs: Box::new(rhs),
            };
        }
        expr
    }

    fn parse_term(&mut self) -> Expr {
        let mut expr = self.parse_unary();
        loop {
            let op = match self.peek() {
                Some(FormulaToken::Star) => BinaryOp::Mul,
                Some(FormulaToken::Slash) => BinaryOp::Div,
                _ => break,
            };
            self.bump();
            let rhs = self.parse_unary();
            expr = Expr::Binary {
                op,
                lhs: Box::new(expr),
                rhs: Box::new(rhs),
            };
        }
        expr
    }

    fn parse_unary(&mut self) -> Expr {
        match self.peek() {
            Some(FormulaToken::Plus) => {
                self.bump();
                Expr::Unary {
                    op: UnaryOp::Plus,
                    rhs: Box::new(self.parse_unary()),
                }
            }
            Some(FormulaToken::Minus) => {
                self.bump();
                Expr::Unary {
                    op: UnaryOp::Minus,
                    rhs: Box::new(self.parse_unary()),
                }
            }
            _ => self.parse_primary(),
        }
    }

    fn parse_primary(&mut self) -> Expr {
        match self.bump() {
            None => Expr::Empty,
            Some(FormulaToken::Number(n)) => Expr::Number(n),
            Some(FormulaToken::String(s)) => Expr::Scalar(JsonValue::String(s)),
            Some(FormulaToken::ErrorCode(code)) => Expr::Scalar(JsonValue::String(code)),
            Some(FormulaToken::Ident(name)) => self.parse_ident_primary(name),
            Some(FormulaToken::LParen) => {
                let inner = self.parse_expression();
                if !matches!(self.bump(), Some(FormulaToken::RParen)) {
                    return Expr::Error(ERROR_VALUE);
                }
                inner
            }
            _ => Expr::Error(ERROR_VALUE),
        }
    }

    fn parse_ident_primary(&mut self, ident: String) -> Expr {
        if ident == "TRUE" {
            return Expr::Scalar(JsonValue::Bool(true));
        }
        if ident == "FALSE" {
            return Expr::Scalar(JsonValue::Bool(false));
        }

        if matches!(self.peek(), Some(FormulaToken::LParen)) {
            self.bump();
            let mut args = Vec::new();
            if matches!(self.peek(), Some(FormulaToken::RParen)) {
                self.bump();
                return Expr::FunctionCall { name: ident, args };
            }

            loop {
                args.push(self.parse_expression());
                match self.peek() {
                    Some(FormulaToken::Comma) => {
                        self.bump();
                        continue;
                    }
                    Some(FormulaToken::RParen) => {
                        self.bump();
                        break;
                    }
                    _ => return Expr::Error(ERROR_VALUE),
                }
            }

            return Expr::FunctionCall { name: ident, args };
        }

        let start = match parse_reference_token(&ident) {
            Ok(coord) => coord,
            Err(code) => return Expr::Error(code),
        };

        if matches!(self.peek(), Some(FormulaToken::Colon)) {
            self.bump();
            let end_ident = match self.bump() {
                Some(FormulaToken::Ident(name)) => name,
                _ => return Expr::Error(ERROR_VALUE),
            };
            let end = match parse_reference_token(&end_ident) {
                Ok(coord) => coord,
                Err(_) => return Expr::Error(ERROR_REF),
            };
            Expr::Range(start, end)
        } else {
            Expr::Reference(start)
        }
    }
}

fn parse_reference_token(token: &str) -> Result<CellCoord, &'static str> {
    parse_a1(token).map_err(|_| {
        let stripped = token.replace('$', "");
        if stripped.chars().any(|ch| ch.is_ascii_digit()) {
            ERROR_REF
        } else {
            ERROR_NAME
        }
    })
}

fn tokenize_formula(expr: &str) -> Result<Vec<FormulaToken>, &'static str> {
    let mut tokens = Vec::new();
    let mut chars = expr.chars().peekable();

    while let Some(ch) = chars.peek().copied() {
        if ch.is_whitespace() {
            chars.next();
            continue;
        }

        match ch {
            // Sheet prefixes are not supported in this lightweight evaluator yet.
            // Mirror the JS fallback behavior by treating them as invalid references.
            '!' => return Err(ERROR_REF),
            '\'' => {
                // Only treat `'Sheet Name'!A1`-style prefixes as #REF!. Any other
                // stray `'` should behave like an unknown token and map to
                // #VALUE! (matching the current JS evaluator).
                let mut lookahead = chars.clone();
                lookahead.next();
                let mut saw_closing = false;
                while let Some(next) = lookahead.next() {
                    if next == '\'' {
                        saw_closing = true;
                        break;
                    }
                }
                if !saw_closing {
                    return Err(ERROR_VALUE);
                }
                if lookahead.next() != Some('!') {
                    return Err(ERROR_VALUE);
                }
                match lookahead.peek().copied() {
                    Some('$') => return Err(ERROR_REF),
                    Some(next) if next.is_ascii_alphabetic() => return Err(ERROR_REF),
                    _ => return Err(ERROR_VALUE),
                }
            }
            '"' => {
                chars.next();
                let mut buf = String::new();
                while let Some(next) = chars.next() {
                    if next == '"' {
                        if matches!(chars.peek(), Some('"')) {
                            // Match the JS evaluator behavior: keep doubled quotes as-is.
                            buf.push('"');
                            buf.push('"');
                            chars.next();
                            continue;
                        }
                        break;
                    }
                    buf.push(next);
                }
                tokens.push(FormulaToken::String(buf));
            }
            '#' => {
                chars.next();
                let mut buf = String::from("#");
                let mut saw_char = false;
                while let Some(next) = chars.peek().copied() {
                    if next.is_whitespace()
                        || matches!(next, ',' | ')' | '(' | '+' | '-' | '*' | '/')
                    {
                        break;
                    }
                    saw_char = true;
                    buf.push(next);
                    chars.next();
                }
                if !saw_char {
                    return Err(ERROR_VALUE);
                }
                tokens.push(FormulaToken::ErrorCode(buf));
            }
            '0'..='9' | '.' => {
                let mut buf = String::new();
                while let Some(ch2) = chars.peek().copied() {
                    if ch2.is_ascii_digit() || ch2 == '.' {
                        buf.push(ch2);
                        chars.next();
                        continue;
                    }

                    if ch2 == 'e' || ch2 == 'E' {
                        // Match the JS tokenizer behavior: only treat `e`/`E` as
                        // exponent if it's followed by an optional sign and at
                        // least one digit. Otherwise, stop the number token
                        // before the `e` and let the parser handle the
                        // remaining tokens.
                        let mut lookahead = chars.clone();
                        lookahead.next(); // consume the `e`
                        if matches!(lookahead.peek(), Some('+') | Some('-')) {
                            lookahead.next();
                        }

                        if !lookahead.peek().is_some_and(|next| next.is_ascii_digit()) {
                            break;
                        }

                        buf.push(ch2);
                        chars.next();

                        if let Some(sign) = chars.peek().copied() {
                            if sign == '+' || sign == '-' {
                                buf.push(sign);
                                chars.next();
                            }
                        }

                        while let Some(exp_ch) = chars.peek().copied() {
                            if exp_ch.is_ascii_digit() {
                                buf.push(exp_ch);
                                chars.next();
                            } else {
                                break;
                            }
                        }

                        continue;
                    }

                    break;
                }

                let number: f64 = buf.parse().map_err(|_| ERROR_VALUE)?;
                tokens.push(FormulaToken::Number(number));
            }
            '+' => {
                chars.next();
                tokens.push(FormulaToken::Plus);
            }
            '-' => {
                chars.next();
                tokens.push(FormulaToken::Minus);
            }
            '*' => {
                chars.next();
                tokens.push(FormulaToken::Star);
            }
            '/' => {
                chars.next();
                tokens.push(FormulaToken::Slash);
            }
            '>' => {
                chars.next();
                if chars.peek().is_some_and(|next| *next == '=') {
                    chars.next();
                    tokens.push(FormulaToken::OtherOp(">=".to_string()));
                } else {
                    tokens.push(FormulaToken::OtherOp(">".to_string()));
                }
            }
            '<' => {
                chars.next();
                if chars.peek().is_some_and(|next| *next == '=') {
                    chars.next();
                    tokens.push(FormulaToken::OtherOp("<=".to_string()));
                } else if chars.peek().is_some_and(|next| *next == '>') {
                    chars.next();
                    tokens.push(FormulaToken::OtherOp("<>".to_string()));
                } else {
                    tokens.push(FormulaToken::OtherOp("<".to_string()));
                }
            }
            '=' | '^' | '&' => {
                chars.next();
                tokens.push(FormulaToken::OtherOp(ch.to_string()));
            }
            '(' => {
                chars.next();
                tokens.push(FormulaToken::LParen);
            }
            ')' => {
                chars.next();
                tokens.push(FormulaToken::RParen);
            }
            ',' => {
                chars.next();
                tokens.push(FormulaToken::Comma);
            }
            ':' => {
                chars.next();
                tokens.push(FormulaToken::Colon);
            }
            _ if ch.is_ascii_alphabetic() || ch == '$' || ch == '_' => {
                let mut buf = String::new();
                while let Some(ch2) = chars.peek().copied() {
                    if ch2.is_ascii_alphanumeric() || ch2 == '$' || ch2 == '_' || ch2 == '.' {
                        buf.push(ch2.to_ascii_uppercase());
                        chars.next();
                    } else {
                        break;
                    }
                }
                if buf.is_empty() {
                    return Err(ERROR_VALUE);
                }
                tokens.push(FormulaToken::Ident(buf));
            }
            _ => return Err(ERROR_VALUE),
        }
    }

    Ok(tokens)
}

#[derive(Clone, Debug, PartialEq)]
enum EvalValue {
    Scalar(JsonValue),
    Array(Vec<JsonValue>),
}

fn eval_formula<F>(expr: &str, mut get_cell: F) -> Result<JsonValue, WorkbookError>
where
    F: FnMut(CellCoord) -> Result<JsonValue, WorkbookError>,
{
    let tokens = match tokenize_formula(expr) {
        Ok(tokens) => tokens,
        Err(code) => return Ok(JsonValue::String(code.to_string())),
    };
    if tokens.is_empty() {
        return Ok(JsonValue::Null);
    }

    let mut parser = FormulaParser::new(tokens);
    let parsed = parser.parse_expression();

    let value = eval_expr(&parsed, &mut get_cell)?;
    Ok(match value {
        EvalValue::Scalar(v) => v,
        EvalValue::Array(arr) => arr.into_iter().next().unwrap_or(JsonValue::Null),
    })
}

fn eval_expr<F>(expr: &Expr, get_cell: &mut F) -> Result<EvalValue, WorkbookError>
where
    F: FnMut(CellCoord) -> Result<JsonValue, WorkbookError>,
{
    match expr {
        Expr::Empty => Ok(EvalValue::Scalar(JsonValue::Null)),
        Expr::Error(code) => Ok(EvalValue::Scalar(JsonValue::String((*code).to_string()))),
        Expr::Number(n) => Ok(EvalValue::Scalar(number_to_json(*n))),
        Expr::Scalar(value) => Ok(EvalValue::Scalar(value.clone())),
        Expr::Reference(coord) => Ok(EvalValue::Scalar(get_cell(*coord)?)),
        Expr::Range(start, end) => {
            let (top_left, bottom_right) = normalize_range(*start, *end);
            let mut values = Vec::new();
            for row in top_left.row..=bottom_right.row {
                for col in top_left.col..=bottom_right.col {
                    values.push(get_cell(CellCoord::new(row, col))?);
                }
            }
            Ok(EvalValue::Array(values))
        }
        Expr::Unary { op, rhs } => {
            let rhs_value = eval_expr(rhs, get_cell)?;
            Ok(apply_unary(*op, rhs_value))
        }
        Expr::Binary { op, lhs, rhs } => {
            let lhs_value = eval_expr(lhs, get_cell)?;
            let rhs_value = eval_expr(rhs, get_cell)?;
            Ok(apply_binary(*op, lhs_value, rhs_value))
        }
        Expr::FunctionCall { name, args } => {
            let mut evaluated = Vec::with_capacity(args.len());
            for arg in args {
                evaluated.push(eval_expr(arg, get_cell)?);
            }
            Ok(eval_function(name, &evaluated))
        }
    }
}

fn normalize_range(start: CellCoord, end: CellCoord) -> (CellCoord, CellCoord) {
    let top = start.row.min(end.row);
    let left = start.col.min(end.col);
    let bottom = start.row.max(end.row);
    let right = start.col.max(end.col);
    (CellCoord::new(top, left), CellCoord::new(bottom, right))
}

fn is_error_code(value: &JsonValue) -> Option<&str> {
    match value {
        JsonValue::String(s) if s.starts_with('#') => Some(s),
        _ => None,
    }
}

fn to_number(value: &JsonValue) -> Option<f64> {
    match value {
        JsonValue::Null => Some(0.0),
        JsonValue::Bool(b) => Some(if *b { 1.0 } else { 0.0 }),
        JsonValue::Number(num) => num.as_f64().filter(|n| n.is_finite()),
        JsonValue::String(s) => {
            let trimmed = s.trim();
            if trimmed.is_empty() {
                return Some(0.0);
            }
            parse_js_number(trimmed)
        }
        _ => None,
    }
}

fn parse_js_number(text: &str) -> Option<f64> {
    // Mirror the JS evaluator's `Number(trimmed)` behavior for a small set of
    // string forms that Rust's `f64::from_str` doesn't accept (notably
    // hex/binary/octal prefixes).
    //
    // We only care about finite values; non-finite values should be treated as
    // invalid (`null` in the JS evaluator).
    if let Ok(num) = text.parse::<f64>() {
        return num.is_finite().then_some(num);
    }

    // JS Number() accepts 0x/0b/0o prefixes, but *not* with a leading sign.
    // (e.g. Number("0x10") === 16, Number("-0x10") === NaN)
    let (radix, digits) = if let Some(rest) = text.strip_prefix("0x").or_else(|| text.strip_prefix("0X")) {
        (16, rest)
    } else if let Some(rest) = text.strip_prefix("0b").or_else(|| text.strip_prefix("0B")) {
        (2, rest)
    } else if let Some(rest) = text.strip_prefix("0o").or_else(|| text.strip_prefix("0O")) {
        (8, rest)
    } else {
        return None;
    };

    if digits.is_empty() {
        return None;
    }

    let value = u64::from_str_radix(digits, radix).ok()? as f64;
    value.is_finite().then_some(value)
}

fn number_to_json(value: f64) -> JsonValue {
    if !value.is_finite() {
        return JsonValue::String(ERROR_NUM.to_string());
    }
    match serde_json::Number::from_f64(value) {
        Some(num) => JsonValue::Number(num),
        None => JsonValue::String(ERROR_NUM.to_string()),
    }
}

fn apply_unary(op: UnaryOp, rhs: EvalValue) -> EvalValue {
    let rhs_scalar = match rhs {
        EvalValue::Scalar(v) => v,
        EvalValue::Array(_) => return EvalValue::Scalar(JsonValue::String(ERROR_VALUE.to_string())),
    };

    if let Some(err) = is_error_code(&rhs_scalar) {
        return EvalValue::Scalar(JsonValue::String(err.to_string()));
    }

    let num = match to_number(&rhs_scalar) {
        Some(n) => n,
        None => return EvalValue::Scalar(JsonValue::String(ERROR_VALUE.to_string())),
    };

    match op {
        UnaryOp::Plus => EvalValue::Scalar(number_to_json(num)),
        UnaryOp::Minus => EvalValue::Scalar(number_to_json(-num)),
    }
}

fn apply_binary(op: BinaryOp, lhs: EvalValue, rhs: EvalValue) -> EvalValue {
    let lhs_scalar = match lhs {
        EvalValue::Scalar(v) => v,
        EvalValue::Array(_) => return EvalValue::Scalar(JsonValue::String(ERROR_VALUE.to_string())),
    };
    let rhs_scalar = match rhs {
        EvalValue::Scalar(v) => v,
        EvalValue::Array(_) => return EvalValue::Scalar(JsonValue::String(ERROR_VALUE.to_string())),
    };

    if let Some(err) = is_error_code(&lhs_scalar) {
        return EvalValue::Scalar(JsonValue::String(err.to_string()));
    }
    if let Some(err) = is_error_code(&rhs_scalar) {
        return EvalValue::Scalar(JsonValue::String(err.to_string()));
    }

    let lhs_num = match to_number(&lhs_scalar) {
        Some(n) => n,
        None => return EvalValue::Scalar(JsonValue::String(ERROR_VALUE.to_string())),
    };
    let rhs_num = match to_number(&rhs_scalar) {
        Some(n) => n,
        None => return EvalValue::Scalar(JsonValue::String(ERROR_VALUE.to_string())),
    };

    let result = match op {
        BinaryOp::Add => lhs_num + rhs_num,
        BinaryOp::Sub => lhs_num - rhs_num,
        BinaryOp::Mul => lhs_num * rhs_num,
        BinaryOp::Div => {
            if rhs_num == 0.0 {
                return EvalValue::Scalar(JsonValue::String(ERROR_DIV0.to_string()));
            }
            lhs_num / rhs_num
        }
    };

    EvalValue::Scalar(number_to_json(result))
}

fn eval_function(name: &str, args: &[EvalValue]) -> EvalValue {
    let upper = name.to_ascii_uppercase();
    if upper == "SUM" {
        let mut nums = Vec::new();
        if let Some(err) = flatten_numbers(args, &mut nums) {
            return EvalValue::Scalar(JsonValue::String(err));
        }
        let sum: f64 = nums.into_iter().sum();
        return EvalValue::Scalar(number_to_json(sum));
    }

    if upper == "AVERAGE" {
        let mut nums = Vec::new();
        if let Some(err) = flatten_numbers(args, &mut nums) {
            return EvalValue::Scalar(JsonValue::String(err));
        }
        if nums.is_empty() {
            return EvalValue::Scalar(number_to_json(0.0));
        }
        let sum: f64 = nums.iter().sum();
        return EvalValue::Scalar(number_to_json(sum / nums.len() as f64));
    }

    if upper == "IF" {
        let cond = args
            .get(0)
            .cloned()
            .unwrap_or(EvalValue::Scalar(JsonValue::Null));
        if let EvalValue::Scalar(ref scalar) = cond {
            if let Some(err) = is_error_code(scalar) {
                return EvalValue::Scalar(JsonValue::String(err.to_string()));
            }
        }

        let cond_num = match &cond {
            EvalValue::Scalar(value) => to_number(value),
            EvalValue::Array(_) => None,
        };
        let truthy = match cond_num {
            Some(num) => num != 0.0,
            None => js_truthy(&cond),
        };

        let chosen = if truthy { args.get(1) } else { args.get(2) }
            .cloned()
            .unwrap_or(EvalValue::Scalar(JsonValue::Null));

        if let EvalValue::Scalar(ref scalar) = chosen {
            if let Some(err) = is_error_code(scalar) {
                return EvalValue::Scalar(JsonValue::String(err.to_string()));
            }
        }

        return match chosen {
            EvalValue::Scalar(v) => EvalValue::Scalar(v),
            EvalValue::Array(arr) => {
                EvalValue::Scalar(arr.into_iter().next().unwrap_or(JsonValue::Null))
            }
        };
    }

    if upper == "VLOOKUP" {
        return EvalValue::Scalar(JsonValue::String(ERROR_NA.to_string()));
    }

    EvalValue::Scalar(JsonValue::String(ERROR_NAME.to_string()))
}

fn js_truthy(value: &EvalValue) -> bool {
    match value {
        // JS arrays are objects, so Boolean([]) is always true.
        EvalValue::Array(_) => true,
        EvalValue::Scalar(scalar) => match scalar {
            JsonValue::Null => false,
            JsonValue::Bool(b) => *b,
            JsonValue::Number(num) => num.as_f64().is_some_and(|v| v != 0.0),
            JsonValue::String(s) => !s.is_empty(),
            _ => true,
        },
    }
}

fn flatten_numbers(values: &[EvalValue], out: &mut Vec<f64>) -> Option<String> {
    for val in values {
        match val {
            EvalValue::Scalar(scalar) => {
                if let Some(err) = is_error_code(scalar) {
                    return Some(err.to_string());
                }
                if let Some(num) = to_number(scalar) {
                    out.push(num);
                }
            }
            EvalValue::Array(arr) => {
                for scalar in arr {
                    if let Some(err) = is_error_code(scalar) {
                        return Some(err.to_string());
                    }
                    if let Some(num) = to_number(scalar) {
                        out.push(num);
                    }
                }
            }
        }
    }
    None
}

#[cfg(test)]
mod tests {
    use super::*;
    use pretty_assertions::assert_eq;
    use serde_json::json;

    #[test]
    fn parse_and_format_a1() {
        assert_eq!(parse_a1("A1"), Ok(CellCoord::new(1, 1)));
        assert_eq!(parse_a1("$A$1"), Ok(CellCoord::new(1, 1)));
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

    #[test]
    fn sum_over_range_evaluates() {
        let mut wb = Workbook::new();
        wb.set_cell("A1", json!(1), None).unwrap();
        wb.set_cell("A2", json!(2), None).unwrap();
        wb.set_cell("A3", json!("=SUM(A1:A2)"), None).unwrap();

        wb.recalculate(None).unwrap();
        let cell = wb.get_cell("A3", None).unwrap();
        assert_eq!(cell.value, json!(3.0));
    }

    #[test]
    fn division_by_zero_returns_excel_error_code() {
        let mut wb = Workbook::new();
        wb.set_cell("A1", json!(0), None).unwrap();
        wb.set_cell("B1", json!("=1/A1"), None).unwrap();

        wb.recalculate(None).unwrap();
        let cell = wb.get_cell("B1", None).unwrap();
        assert_eq!(cell.value, json!(ERROR_DIV0));
    }

    #[test]
    fn average_over_values_evaluates() {
        let mut wb = Workbook::new();
        wb.set_cell("A1", json!("=AVERAGE(1,2)"), None).unwrap();

        wb.recalculate(None).unwrap();
        let cell = wb.get_cell("A1", None).unwrap();
        assert_eq!(cell.value, json!(1.5));
    }

    #[test]
    fn if_function_selects_branches() {
        let mut wb = Workbook::new();
        wb.set_cell("A1", json!("=IF(1,2,3)"), None).unwrap();
        wb.set_cell("A2", json!("=IF(0,2,3)"), None).unwrap();

        wb.recalculate(None).unwrap();
        assert_eq!(wb.get_cell("A1", None).unwrap().value, json!(2.0));
        assert_eq!(wb.get_cell("A2", None).unwrap().value, json!(3.0));
    }

    #[test]
    fn comparison_operators_terminate_expression_like_js_evaluator() {
        let mut wb = Workbook::new();
        // The current JS evaluator tokenizes comparison operators but doesn't
        // evaluate them, so parsing stops at the operator and the left-hand
        // side is returned.
        wb.set_cell("A1", json!("=1>0"), None).unwrap();
        wb.set_cell("A2", json!("=SUM(1,2)>0"), None).unwrap();

        wb.recalculate(None).unwrap();
        assert_eq!(wb.get_cell("A1", None).unwrap().value, json!(1.0));
        assert_eq!(wb.get_cell("A2", None).unwrap().value, json!(3.0));
    }

    #[test]
    fn vlookup_is_stubbed_to_na() {
        let mut wb = Workbook::new();
        wb.set_range(
            "A1:B2",
            vec![vec![json!(1), json!(2)], vec![json!(3), json!(4)]],
            None,
        )
        .unwrap();
        wb.set_cell("C1", json!("=VLOOKUP(1,A1:B2,2,FALSE)"), None)
            .unwrap();

        wb.recalculate(None).unwrap();
        assert_eq!(wb.get_cell("C1", None).unwrap().value, json!(ERROR_NA));
    }

    #[test]
    fn sheet_references_return_ref_error() {
        let mut wb = Workbook::new();
        wb.set_cell("A1", json!("=Sheet2!A1"), None).unwrap();

        wb.recalculate(None).unwrap();
        assert_eq!(wb.get_cell("A1", None).unwrap().value, json!(ERROR_REF));
    }

    #[test]
    fn stray_apostrophe_returns_value_error() {
        let mut wb = Workbook::new();
        wb.set_cell("A1", json!("='A1"), None).unwrap();

        wb.recalculate(None).unwrap();
        assert_eq!(wb.get_cell("A1", None).unwrap().value, json!(ERROR_VALUE));
    }

    #[test]
    fn dotted_function_names_fall_back_to_name_error() {
        let mut wb = Workbook::new();
        wb.set_cell("A1", json!("=_xlfn.SEQUENCE(1)"), None).unwrap();

        wb.recalculate(None).unwrap();
        assert_eq!(wb.get_cell("A1", None).unwrap().value, json!(ERROR_NAME));
    }

    #[test]
    fn to_number_matches_js_number_parsing_for_hex_strings() {
        let mut wb = Workbook::new();
        wb.set_cell("A1", json!("=SUM(\"0x10\",1)"), None).unwrap();
        wb.set_cell("A2", json!("=\"0x10\"+1"), None).unwrap();
        wb.set_cell("A3", json!("=SUM(\"-0x10\",1)"), None).unwrap();

        wb.recalculate(None).unwrap();

        assert_eq!(wb.get_cell("A1", None).unwrap().value, json!(17.0));
        assert_eq!(wb.get_cell("A2", None).unwrap().value, json!(17.0));
        // JS Number("-0x10") is NaN -> ignored by SUM.
        assert_eq!(wb.get_cell("A3", None).unwrap().value, json!(1.0));
    }

    #[test]
    fn trailing_tokens_are_ignored_like_js_evaluator() {
        let mut wb = Workbook::new();
        wb.set_cell("A1", json!("=1,2"), None).unwrap();
        wb.set_cell("A2", json!("=1e"), None).unwrap();

        wb.recalculate(None).unwrap();

        assert_eq!(wb.get_cell("A1", None).unwrap().value, json!(1.0));
        assert_eq!(wb.get_cell("A2", None).unwrap().value, json!(1.0));
    }

    #[test]
    fn load_from_json_then_recalculate_updates_formula_cells() {
        let json_str = r#"{
            "sheets": {
                "Sheet1": {
                    "cells": {
                        "A1": 1,
                        "A2": "=A1*2"
                    }
                }
            }
        }"#;

        let mut wb = Workbook::from_json_str(json_str).unwrap();
        wb.recalculate(None).unwrap();

        let cell = wb.get_cell("A2", None).unwrap();
        assert_eq!(cell.value, json!(2.0));
    }
}
