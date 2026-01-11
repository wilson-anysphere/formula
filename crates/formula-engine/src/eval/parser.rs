use crate::eval::address::{parse_a1, AddressParseError, CellAddr};
use crate::eval::ast::{
    BinaryOp, CellRef, CompareOp, Expr, NameRef, ParsedExpr, RangeRef, SheetReference, UnaryOp,
};
use crate::value::ErrorKind;
use formula_model::{EXCEL_MAX_COLS, EXCEL_MAX_ROWS};
use thiserror::Error;

#[derive(Debug, Error, Clone, PartialEq, Eq)]
pub enum FormulaParseError {
    #[error("unexpected end of input")]
    UnexpectedEof,
    #[error("unexpected token: {0}")]
    UnexpectedToken(String),
    #[error("invalid address: {0}")]
    InvalidAddress(#[from] AddressParseError),
    #[error("expected {expected}, got {got}")]
    Expected { expected: String, got: String },
}

#[derive(Debug, Clone, PartialEq)]
enum Token {
    Number(f64),
    String(String),
    Ident(String),
    StructuredRef(String),
    SheetName(String),
    Error(ErrorKind),
    LParen,
    RParen,
    Comma,
    Colon,
    Bang,
    At,
    Hash,
    Plus,
    Minus,
    Star,
    Slash,
    Caret,
    Eq,
    Ne,
    Lt,
    Le,
    Gt,
    Ge,
    End,
}

pub struct Parser;

impl Parser {
    pub fn parse(formula: &str) -> Result<ParsedExpr, FormulaParseError> {
        let mut lexer = Lexer::new(formula);
        let tokens = lexer.tokenize()?;
        let mut p = ParserImpl::new(tokens);
        let expr = p.parse_formula()?;
        p.expect(Token::End)?;
        Ok(expr)
    }
}

struct ParserImpl {
    tokens: Vec<Token>,
    pos: usize,
}

#[derive(Debug, Clone, Copy, PartialEq, Eq)]
enum RefStart {
    Cell(CellAddr),
    Col(u32),
    Row(u32),
}

impl ParserImpl {
    fn new(tokens: Vec<Token>) -> Self {
        Self { tokens, pos: 0 }
    }

    fn parse_formula(&mut self) -> Result<ParsedExpr, FormulaParseError> {
        self.parse_compare()
    }

    fn parse_compare(&mut self) -> Result<ParsedExpr, FormulaParseError> {
        let mut left = self.parse_add_sub()?;
        loop {
            let op = match self.peek() {
                Token::Eq => CompareOp::Eq,
                Token::Ne => CompareOp::Ne,
                Token::Lt => CompareOp::Lt,
                Token::Le => CompareOp::Le,
                Token::Gt => CompareOp::Gt,
                Token::Ge => CompareOp::Ge,
                _ => break,
            };
            self.next();
            let right = self.parse_add_sub()?;
            left = Expr::Compare {
                op,
                left: Box::new(left),
                right: Box::new(right),
            };
        }
        Ok(left)
    }

    fn parse_add_sub(&mut self) -> Result<ParsedExpr, FormulaParseError> {
        let mut left = self.parse_mul_div()?;
        loop {
            let op = match self.peek() {
                Token::Plus => BinaryOp::Add,
                Token::Minus => BinaryOp::Sub,
                _ => break,
            };
            self.next();
            let right = self.parse_mul_div()?;
            left = Expr::Binary {
                op,
                left: Box::new(left),
                right: Box::new(right),
            };
        }
        Ok(left)
    }

    fn parse_mul_div(&mut self) -> Result<ParsedExpr, FormulaParseError> {
        let mut left = self.parse_unary()?;
        loop {
            let op = match self.peek() {
                Token::Star => BinaryOp::Mul,
                Token::Slash => BinaryOp::Div,
                _ => break,
            };
            self.next();
            let right = self.parse_unary()?;
            left = Expr::Binary {
                op,
                left: Box::new(left),
                right: Box::new(right),
            };
        }
        Ok(left)
    }

    fn parse_power(&mut self) -> Result<ParsedExpr, FormulaParseError> {
        let left = self.parse_primary()?;
        if *self.peek() == Token::Caret {
            self.next();
            // Excel's exponentiation is right-associative and binds tighter than unary
            // operators (e.g. `-2^2` == `-(2^2)`).
            let right = self.parse_unary()?;
            return Ok(Expr::Binary {
                op: BinaryOp::Pow,
                left: Box::new(left),
                right: Box::new(right),
            });
        }
        Ok(left)
    }

    fn parse_unary(&mut self) -> Result<ParsedExpr, FormulaParseError> {
        match self.peek() {
            Token::Plus => {
                self.next();
                Ok(Expr::Unary {
                    op: UnaryOp::Plus,
                    expr: Box::new(self.parse_unary()?),
                })
            }
            Token::Minus => {
                self.next();
                Ok(Expr::Unary {
                    op: UnaryOp::Minus,
                    expr: Box::new(self.parse_unary()?),
                })
            }
            Token::At => {
                self.next();
                Ok(Expr::ImplicitIntersection(Box::new(self.parse_unary()?)))
            }
            _ => self.parse_power(),
        }
    }

    fn parse_primary(&mut self) -> Result<ParsedExpr, FormulaParseError> {
        let mut expr = match self.peek().clone() {
            Token::Number(n) => {
                self.next();
                if *self.peek() == Token::Colon {
                    let row = row_index_from_number(n)?;
                    self.parse_cell_or_range(SheetReference::Current, RefStart::Row(row))
                } else {
                    Ok(Expr::Number(n))
                }
            }
            Token::String(s) => {
                self.next();
                Ok(Expr::Text(s))
            }
            Token::Error(e) => {
                self.next();
                Ok(Expr::Error(e))
            }
            Token::StructuredRef(text) => {
                self.next();
                let (sref, end) = crate::structured_refs::parse_structured_ref(&text, 0)
                    .ok_or_else(|| {
                        FormulaParseError::UnexpectedToken(format!(
                            "invalid structured reference: {text}"
                        ))
                    })?;
                if end != text.len() {
                    return Err(FormulaParseError::UnexpectedToken(format!(
                        "invalid structured reference: {text}"
                    )));
                }
                Ok(Expr::StructuredRef(sref))
            }
            Token::Ident(id) => {
                // Function call or reference/name.
                if self.peek_n(1) == Token::LParen {
                    self.parse_function_call()
                } else if self.peek_n(1) == Token::Bang {
                    self.parse_sheet_ref()
                } else {
                    self.next();
                    match id.to_ascii_uppercase().as_str() {
                        "TRUE" => Ok(Expr::Bool(true)),
                        "FALSE" => Ok(Expr::Bool(false)),
                        _ => {
                            if *self.peek() == Token::Colon {
                                // Range start: allow column references like `A:A`.
                                if let Ok(addr) = try_parse_cell_addr(&id) {
                                    self.parse_cell_or_range(
                                        SheetReference::Current,
                                        RefStart::Cell(addr),
                                    )
                                } else if let Some(col) = try_parse_col_ref(&id) {
                                    self.parse_cell_or_range(
                                        SheetReference::Current,
                                        RefStart::Col(col),
                                    )
                                } else {
                                    Err(FormulaParseError::InvalidAddress(
                                        AddressParseError::InvalidA1(id),
                                    ))
                                }
                            } else {
                                match try_parse_cell_addr(&id) {
                                    Ok(addr) => self.parse_cell_or_range(
                                        SheetReference::Current,
                                        RefStart::Cell(addr),
                                    ),
                                    Err(_) => Ok(Expr::NameRef(NameRef {
                                        sheet: SheetReference::Current,
                                        name: id,
                                    })),
                                }
                            }
                        }
                    }
                }
            }
            Token::SheetName(_name) => {
                if self.peek_n(1) == Token::Bang {
                    self.parse_sheet_ref()
                } else {
                    self.next();
                    Ok(Expr::Error(ErrorKind::Name))
                }
            }
            Token::LParen => {
                self.next();
                let expr = self.parse_compare()?;
                self.expect(Token::RParen)?;
                Ok(expr)
            }
            other => Err(FormulaParseError::UnexpectedToken(format!("{other:?}"))),
        }?;

        // Postfix spill-range operator (`#`).
        while *self.peek() == Token::Hash {
            self.next();
            expr = Expr::SpillRange(Box::new(expr));
        }

        Ok(expr)
    }

    fn parse_function_call(&mut self) -> Result<ParsedExpr, FormulaParseError> {
        let (name, original_name) = match self.next() {
            Token::Ident(s) => {
                let upper = s.to_ascii_uppercase();
                let base = upper.strip_prefix("_XLFN.").unwrap_or(&upper).to_string();
                (base, s)
            }
            other => {
                return Err(FormulaParseError::Expected {
                    expected: "identifier".to_string(),
                    got: format!("{other:?}"),
                })
            }
        };
        self.expect(Token::LParen)?;
        let mut args = Vec::new();
        if *self.peek() != Token::RParen {
            loop {
                args.push(self.parse_compare()?);
                if *self.peek() == Token::Comma {
                    self.next();
                    continue;
                }
                break;
            }
        }
        self.expect(Token::RParen)?;
        Ok(Expr::FunctionCall {
            name,
            original_name,
            args,
        })
    }

    fn parse_sheet_ref(&mut self) -> Result<ParsedExpr, FormulaParseError> {
        let sheet = match self.next() {
            Token::Ident(s) | Token::SheetName(s) => SheetReference::Sheet(s),
            other => {
                return Err(FormulaParseError::Expected {
                    expected: "sheet name".to_string(),
                    got: format!("{other:?}"),
                })
            }
        };
        self.expect(Token::Bang)?;

        // External workbook references are not supported yet, but we accept the syntax
        // and let evaluation return `#REF!`.
        if matches!(self.peek(), Token::Ident(id) if id.starts_with('[')) {
            self.next();
            return Ok(Expr::Error(ErrorKind::Ref));
        }

        let start = match self.next() {
            Token::Ident(s) => {
                if let Ok(addr) = parse_a1(&s) {
                    RefStart::Cell(addr)
                } else if *self.peek() == Token::Colon {
                    let col = try_parse_col_ref(&s).ok_or_else(|| {
                        FormulaParseError::InvalidAddress(AddressParseError::InvalidA1(s))
                    })?;
                    RefStart::Col(col)
                } else {
                    return Ok(Expr::NameRef(NameRef { sheet, name: s }));
                }
            }
            Token::Number(n) => {
                if *self.peek() != Token::Colon {
                    return Err(FormulaParseError::Expected {
                        expected: "cell address".to_string(),
                        got: format!("{n}"),
                    });
                }
                RefStart::Row(row_index_from_number(n)?)
            }
            other => {
                return Err(FormulaParseError::Expected {
                    expected: "cell address".to_string(),
                    got: format!("{other:?}"),
                })
            }
        };
        self.parse_cell_or_range(sheet, start)
    }

    fn parse_cell_or_range(
        &mut self,
        sheet: SheetReference<String>,
        start: RefStart,
    ) -> Result<ParsedExpr, FormulaParseError> {
        if *self.peek() == Token::Colon {
            self.next();
            let end_token = self.next();
            let end = match (start, end_token) {
                (RefStart::Cell(_), Token::Ident(s)) => RefStart::Cell(parse_a1(&s)?),
                (RefStart::Col(_), Token::Ident(s)) => {
                    RefStart::Col(try_parse_col_ref(&s).ok_or_else(|| {
                        FormulaParseError::Expected {
                            expected: "column reference".to_string(),
                            got: format!("{s:?}"),
                        }
                    })?)
                }
                (RefStart::Row(_), Token::Number(n)) => RefStart::Row(row_index_from_number(n)?),
                (RefStart::Cell(_), other) => {
                    return Err(FormulaParseError::Expected {
                        expected: "cell address".to_string(),
                        got: format!("{other:?}"),
                    })
                }
                (RefStart::Col(_), other) => {
                    return Err(FormulaParseError::Expected {
                        expected: "column reference".to_string(),
                        got: format!("{other:?}"),
                    })
                }
                (RefStart::Row(_), other) => {
                    return Err(FormulaParseError::Expected {
                        expected: "row reference".to_string(),
                        got: format!("{other:?}"),
                    })
                }
            };

            let (start, end) = match (start, end) {
                (RefStart::Cell(a), RefStart::Cell(b)) => (a, b),
                (RefStart::Col(a), RefStart::Col(b)) => {
                    let max_row = EXCEL_MAX_ROWS.saturating_sub(1);
                    (
                        CellAddr { row: 0, col: a },
                        CellAddr {
                            row: max_row,
                            col: b,
                        },
                    )
                }
                (RefStart::Row(a), RefStart::Row(b)) => {
                    let max_col = EXCEL_MAX_COLS.saturating_sub(1);
                    (
                        CellAddr { row: a, col: 0 },
                        CellAddr {
                            row: b,
                            col: max_col,
                        },
                    )
                }
                _ => {
                    return Err(FormulaParseError::UnexpectedToken(
                        "mixed row/col/cell range".to_string(),
                    ))
                }
            };

            Ok(Expr::RangeRef(RangeRef { sheet, start, end }))
        } else {
            match start {
                RefStart::Cell(addr) => Ok(Expr::CellRef(CellRef { sheet, addr })),
                RefStart::Col(_) | RefStart::Row(_) => Err(FormulaParseError::UnexpectedToken(
                    "expected cell address".to_string(),
                )),
            }
        }
    }

    fn peek(&self) -> &Token {
        self.tokens.get(self.pos).unwrap_or(&Token::End)
    }

    fn peek_n(&self, n: usize) -> Token {
        self.tokens.get(self.pos + n).cloned().unwrap_or(Token::End)
    }

    fn next(&mut self) -> Token {
        let tok = self.peek().clone();
        self.pos += 1;
        tok
    }

    fn expect(&mut self, expected: Token) -> Result<(), FormulaParseError> {
        let got = self.next();
        if got == expected {
            Ok(())
        } else {
            Err(FormulaParseError::Expected {
                expected: format!("{expected:?}"),
                got: format!("{got:?}"),
            })
        }
    }
}

fn try_parse_cell_addr(id: &str) -> Result<CellAddr, AddressParseError> {
    parse_a1(id)
}

fn try_parse_col_ref(id: &str) -> Option<u32> {
    let filtered: String = id.chars().filter(|&ch| ch != '$').collect();
    if filtered.is_empty() || !filtered.chars().all(|ch| ch.is_ascii_alphabetic()) {
        return None;
    }

    let mut col: u32 = 0;
    for ch in filtered.chars() {
        let up = ch.to_ascii_uppercase();
        let digit = (up as u8).wrapping_sub(b'A').wrapping_add(1) as u32;
        col = col.checked_mul(26)?.checked_add(digit)?;
    }
    if col == 0 || col > EXCEL_MAX_COLS {
        return None;
    }
    Some(col - 1)
}

fn row_index_from_number(n: f64) -> Result<u32, FormulaParseError> {
    if !n.is_finite() || n.fract() != 0.0 {
        return Err(FormulaParseError::UnexpectedToken(format!(
            "invalid row reference: {n}"
        )));
    }
    let row_1_based = n as i64;
    if row_1_based <= 0 || row_1_based > i64::from(EXCEL_MAX_ROWS) {
        return Err(FormulaParseError::UnexpectedToken(format!(
            "row out of range: {n}"
        )));
    }
    Ok((row_1_based as u32) - 1)
}

struct Lexer<'a> {
    input: &'a str,
    pos: usize,
    prev: Option<Token>,
}

impl<'a> Lexer<'a> {
    fn new(input: &'a str) -> Self {
        // Permit formulas with or without leading '='.
        let input = input.trim_start();
        let input = input.strip_prefix('=').unwrap_or(input);
        Self {
            input,
            pos: 0,
            prev: None,
        }
    }

    fn tokenize(&mut self) -> Result<Vec<Token>, FormulaParseError> {
        let mut tokens = Vec::new();
        while let Some(ch) = self.peek_char() {
            if ch.is_whitespace() {
                self.pos += ch.len_utf8();
                continue;
            }

            let tok = match ch {
                '(' => {
                    self.pos += 1;
                    Token::LParen
                }
                ')' => {
                    self.pos += 1;
                    Token::RParen
                }
                ',' => {
                    self.pos += 1;
                    Token::Comma
                }
                ':' => {
                    self.pos += 1;
                    Token::Colon
                }
                '!' => {
                    self.pos += 1;
                    Token::Bang
                }
                '@' => {
                    self.pos += 1;
                    Token::At
                }
                '#' => self.lex_hash_or_error()?,
                '[' => self.lex_structured_ref()?,
                '+' => {
                    self.pos += 1;
                    Token::Plus
                }
                '-' => {
                    self.pos += 1;
                    Token::Minus
                }
                '*' => {
                    self.pos += 1;
                    Token::Star
                }
                '/' => {
                    self.pos += 1;
                    Token::Slash
                }
                '^' => {
                    self.pos += 1;
                    Token::Caret
                }
                '=' => {
                    self.pos += 1;
                    Token::Eq
                }
                '<' => {
                    if self.peek_str("<=") {
                        self.pos += 2;
                        Token::Le
                    } else if self.peek_str("<>") {
                        self.pos += 2;
                        Token::Ne
                    } else {
                        self.pos += 1;
                        Token::Lt
                    }
                }
                '>' => {
                    if self.peek_str(">=") {
                        self.pos += 2;
                        Token::Ge
                    } else {
                        self.pos += 1;
                        Token::Gt
                    }
                }
                '"' => self.lex_string()?,
                '\'' => self.lex_sheet_name()?,
                '.' | '0'..='9' => self.lex_number()?,
                _ if is_ident_start(ch) => self.lex_ident()?,
                _ => {
                    return Err(FormulaParseError::UnexpectedToken(format!(
                        "unexpected character '{ch}'"
                    )))
                }
            };
            tokens.push(tok);
            self.prev = tokens.last().cloned();
        }
        tokens.push(Token::End);
        Ok(tokens)
    }

    fn peek_char(&self) -> Option<char> {
        self.input[self.pos..].chars().next()
    }

    fn peek_str(&self, s: &str) -> bool {
        self.input[self.pos..].starts_with(s)
    }

    fn lex_ident(&mut self) -> Result<Token, FormulaParseError> {
        let start = self.pos;
        while let Some(ch) = self.peek_char() {
            if ch.is_ascii_alphanumeric() || ch == '_' || ch == '.' || ch == '$' {
                self.pos += ch.len_utf8();
            } else {
                break;
            }
        }
        if self.peek_char() == Some('[') {
            self.lex_structured_ref_from(start)
        } else {
            Ok(Token::Ident(self.input[start..self.pos].to_string()))
        }
    }

    fn lex_structured_ref(&mut self) -> Result<Token, FormulaParseError> {
        let start = self.pos;
        self.lex_structured_ref_from(start)
    }

    fn lex_structured_ref_from(&mut self, start: usize) -> Result<Token, FormulaParseError> {
        let mut depth: i32 = 0;
        while let Some(ch) = self.peek_char() {
            self.pos += ch.len_utf8();
            match ch {
                '[' => depth += 1,
                ']' => {
                    depth -= 1;
                    if depth == 0 {
                        return Ok(Token::StructuredRef(
                            self.input[start..self.pos].to_string(),
                        ));
                    }
                }
                _ => {}
            }
        }
        Err(FormulaParseError::UnexpectedEof)
    }

    fn lex_number(&mut self) -> Result<Token, FormulaParseError> {
        let start = self.pos;
        let mut saw_dot = false;
        while let Some(ch) = self.peek_char() {
            match ch {
                '0'..='9' => self.pos += 1,
                '.' if !saw_dot => {
                    saw_dot = true;
                    self.pos += 1;
                }
                'E' | 'e' => {
                    self.pos += 1;
                    if matches!(self.peek_char(), Some('+') | Some('-')) {
                        self.pos += 1;
                    }
                }
                _ => break,
            }
        }
        let s = &self.input[start..self.pos];
        let n: f64 = s.parse().map_err(|_| {
            FormulaParseError::UnexpectedToken(format!("invalid number literal: {s}"))
        })?;
        Ok(Token::Number(n))
    }

    fn lex_string(&mut self) -> Result<Token, FormulaParseError> {
        // Consume opening quote.
        self.pos += 1;
        let mut out = String::new();
        loop {
            match self.peek_char() {
                Some('"') => {
                    if self.peek_str("\"\"") {
                        out.push('"');
                        self.pos += 2;
                        continue;
                    }
                    self.pos += 1;
                    break;
                }
                Some(ch) => {
                    out.push(ch);
                    self.pos += ch.len_utf8();
                }
                None => return Err(FormulaParseError::UnexpectedEof),
            }
        }
        Ok(Token::String(out))
    }

    fn lex_sheet_name(&mut self) -> Result<Token, FormulaParseError> {
        // Consume opening quote.
        self.pos += 1;
        let mut out = String::new();
        loop {
            match self.peek_char() {
                Some('\'') => {
                    if self.peek_str("''") {
                        out.push('\'');
                        self.pos += 2;
                        continue;
                    }
                    self.pos += 1;
                    break;
                }
                Some(ch) => {
                    out.push(ch);
                    self.pos += ch.len_utf8();
                }
                None => return Err(FormulaParseError::UnexpectedEof),
            }
        }
        Ok(Token::SheetName(out))
    }

    fn lex_error(&mut self) -> Result<Token, FormulaParseError> {
        let start = self.pos;
        while let Some(ch) = self.peek_char() {
            if ch.is_ascii_alphanumeric() || ch == '#' || ch == '/' || ch == '!' || ch == '?' {
                self.pos += ch.len_utf8();
            } else {
                break;
            }
        }
        let s = &self.input[start..self.pos];
        let kind = match s.to_ascii_uppercase().as_str() {
            "#DIV/0!" => ErrorKind::Div0,
            "#VALUE!" => ErrorKind::Value,
            "#REF!" => ErrorKind::Ref,
            "#NAME?" => ErrorKind::Name,
            "#N/A" => ErrorKind::NA,
            "#NULL!" => ErrorKind::Null,
            "#SPILL!" => ErrorKind::Spill,
            "#CALC!" => ErrorKind::Calc,
            _ => ErrorKind::Value,
        };
        Ok(Token::Error(kind))
    }

    fn lex_hash_or_error(&mut self) -> Result<Token, FormulaParseError> {
        // Spill-range operator is postfix (`A1#`), while error literals start with `#` (`#REF!`).
        let is_postfix = self.prev.as_ref().is_some_and(|t| matches!(t, Token::Ident(_) | Token::StructuredRef(_) | Token::RParen));
        if is_postfix {
            self.pos += 1;
            return Ok(Token::Hash);
        }
        self.lex_error()
    }
}

fn is_ident_start(ch: char) -> bool {
    ch.is_ascii_alphabetic() || ch == '_' || ch == '$'
}
