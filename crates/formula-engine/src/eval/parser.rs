use crate::eval::address::{parse_a1, AddressParseError, CellAddr};
use crate::eval::ast::{
    BinaryOp, CellRef, CompareOp, Expr, ParsedExpr, RangeRef, SheetReference, UnaryOp,
};
use crate::value::ErrorKind;
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
    Plus,
    Minus,
    Star,
    Slash,
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
                Ok(Expr::ImplicitIntersection(Box::new(
                    self.parse_unary()?,
                )))
            }
            _ => self.parse_primary(),
        }
    }

    fn parse_primary(&mut self) -> Result<ParsedExpr, FormulaParseError> {
        match self.peek().clone() {
            Token::Number(n) => {
                self.next();
                Ok(Expr::Number(n))
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
                let (sref, end) = crate::structured_refs::parse_structured_ref(&text, 0).ok_or_else(|| {
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
                        _ => match try_parse_cell_addr(&id) {
                            Ok(addr) => self.parse_cell_or_range(SheetReference::Current, addr),
                            Err(_) => Ok(Expr::Error(ErrorKind::Name)),
                        },
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
        }
    }

    fn parse_function_call(&mut self) -> Result<ParsedExpr, FormulaParseError> {
        let name = match self.next() {
            Token::Ident(s) => s.to_ascii_uppercase(),
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
        Ok(Expr::FunctionCall { name, args })
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

        let addr_token = self.next();
        let addr_str = match addr_token {
            Token::Ident(s) => s,
            other => {
                return Err(FormulaParseError::Expected {
                    expected: "cell address".to_string(),
                    got: format!("{other:?}"),
                })
            }
        };
        let addr = parse_a1(&addr_str)?;
        self.parse_cell_or_range(sheet, addr)
    }

    fn parse_cell_or_range(
        &mut self,
        sheet: SheetReference<String>,
        start: CellAddr,
    ) -> Result<ParsedExpr, FormulaParseError> {
        if *self.peek() == Token::Colon {
            self.next();
            let end_token = self.next();
            let end_str = match end_token {
                Token::Ident(s) => s,
                other => {
                    return Err(FormulaParseError::Expected {
                        expected: "cell address".to_string(),
                        got: format!("{other:?}"),
                    })
                }
            };
            let end = parse_a1(&end_str)?;
            Ok(Expr::RangeRef(RangeRef {
                sheet,
                start,
                end,
            }))
        } else {
            Ok(Expr::CellRef(CellRef { sheet, addr: start }))
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

struct Lexer<'a> {
    input: &'a str,
    pos: usize,
}

impl<'a> Lexer<'a> {
    fn new(input: &'a str) -> Self {
        // Permit formulas with or without leading '='.
        let input = input.trim_start();
        let input = input.strip_prefix('=').unwrap_or(input);
        Self { input, pos: 0 }
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
                '#' => self.lex_error()?,
                '.' | '0'..='9' => self.lex_number()?,
                _ if is_ident_start(ch) => self.lex_ident()?,
                _ => {
                    return Err(FormulaParseError::UnexpectedToken(format!(
                        "unexpected character '{ch}'"
                    )))
                }
            };
            tokens.push(tok);
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
                        return Ok(Token::StructuredRef(self.input[start..self.pos].to_string()));
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
            _ => ErrorKind::Value,
        };
        Ok(Token::Error(kind))
    }
}

fn is_ident_start(ch: char) -> bool {
    ch.is_ascii_alphabetic() || ch == '_' || ch == '$'
}
