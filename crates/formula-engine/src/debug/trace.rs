use crate::error::ExcelError;
use crate::eval::{
    parse_a1, CellAddr, CompareOp, EvalContext, FormulaParseError, SheetReference, UnaryOp,
};
use crate::functions::{ArgValue as FnArgValue, FunctionContext, SheetId as FnSheetId};
use crate::value::{Array, ErrorKind, NumberLocale, Value};
use std::cmp::Ordering;

/// Half-open byte span into the original formula string.
#[derive(Debug, Clone, Copy, PartialEq, Eq, Hash)]
pub struct Span {
    pub start: usize,
    pub end: usize,
}

impl Span {
    pub fn new(start: usize, end: usize) -> Self {
        Self { start, end }
    }
}

#[derive(Debug, Clone, PartialEq, Eq, Hash)]
pub enum TraceRef {
    Cell { sheet: FnSheetId, addr: CellAddr },
    Range {
        sheet: FnSheetId,
        start: CellAddr,
        end: CellAddr,
    },
}

#[derive(Debug, Clone, PartialEq, Eq)]
pub enum TraceKind {
    Number,
    Text,
    Bool,
    Blank,
    Error,
    ArrayLiteral { rows: usize, cols: usize },
    CellRef,
    RangeRef,
    NameRef { name: String },
    Group,
    Unary { op: UnaryOp },
    Binary { op: crate::eval::BinaryOp },
    Compare { op: CompareOp },
    FunctionCall { name: String },
    ImplicitIntersection,
    SpillRange,
}

#[derive(Debug, Clone, PartialEq)]
pub struct TraceNode {
    pub kind: TraceKind,
    pub span: Span,
    pub value: Value,
    pub reference: Option<TraceRef>,
    pub children: Vec<TraceNode>,
}

#[derive(Debug, Clone, PartialEq)]
pub struct DebugEvaluation {
    pub formula: String,
    pub value: Value,
    pub trace: TraceNode,
}

/// A span-aware expression tree used exclusively for on-demand debugging.
#[derive(Debug, Clone, PartialEq)]
pub struct SpannedExpr<S> {
    pub span: Span,
    pub kind: SpannedExprKind<S>,
}

#[derive(Debug, Clone, PartialEq)]
pub enum SpannedExprKind<S> {
    Number(f64),
    Text(String),
    Bool(bool),
    Blank,
    Error(ErrorKind),
    ArrayLiteral {
        rows: Vec<Vec<SpannedExpr<S>>>,
    },
    CellRef(crate::eval::CellRef<S>),
    RangeRef(crate::eval::RangeRef<S>),
    NameRef(crate::eval::NameRef<S>),
    Group(Box<SpannedExpr<S>>),
    Unary {
        op: UnaryOp,
        expr: Box<SpannedExpr<S>>,
    },
    Binary {
        op: crate::eval::BinaryOp,
        left: Box<SpannedExpr<S>>,
        right: Box<SpannedExpr<S>>,
    },
    Compare {
        op: CompareOp,
        left: Box<SpannedExpr<S>>,
        right: Box<SpannedExpr<S>>,
    },
    FunctionCall {
        name: String,
        args: Vec<SpannedExpr<S>>,
    },
    ImplicitIntersection(Box<SpannedExpr<S>>),
    /// Dynamic array spill range operator (`#`), e.g. `A1#`.
    SpillRange(Box<SpannedExpr<S>>),
}

impl<S: Clone> SpannedExpr<S> {
    pub fn map_sheets<T: Clone, F>(&self, f: &mut F) -> SpannedExpr<T>
    where
        F: FnMut(&SheetReference<S>) -> SheetReference<T>,
    {
        let kind = match &self.kind {
            SpannedExprKind::Number(n) => SpannedExprKind::Number(*n),
            SpannedExprKind::Text(s) => SpannedExprKind::Text(s.clone()),
            SpannedExprKind::Bool(b) => SpannedExprKind::Bool(*b),
            SpannedExprKind::Blank => SpannedExprKind::Blank,
            SpannedExprKind::Error(e) => SpannedExprKind::Error(*e),
            SpannedExprKind::ArrayLiteral { rows } => SpannedExprKind::ArrayLiteral {
                rows: rows
                    .iter()
                    .map(|row| row.iter().map(|e| e.map_sheets(f)).collect())
                    .collect(),
            },
            SpannedExprKind::CellRef(r) => SpannedExprKind::CellRef(crate::eval::CellRef {
                sheet: f(&r.sheet),
                addr: r.addr,
            }),
            SpannedExprKind::RangeRef(r) => SpannedExprKind::RangeRef(crate::eval::RangeRef {
                sheet: f(&r.sheet),
                start: r.start,
                end: r.end,
            }),
            SpannedExprKind::NameRef(n) => SpannedExprKind::NameRef(crate::eval::NameRef {
                sheet: f(&n.sheet),
                name: n.name.clone(),
            }),
            SpannedExprKind::Group(expr) => SpannedExprKind::Group(Box::new(expr.map_sheets(f))),
            SpannedExprKind::Unary { op, expr } => SpannedExprKind::Unary {
                op: *op,
                expr: Box::new(expr.map_sheets(f)),
            },
            SpannedExprKind::Binary { op, left, right } => SpannedExprKind::Binary {
                op: *op,
                left: Box::new(left.map_sheets(f)),
                right: Box::new(right.map_sheets(f)),
            },
            SpannedExprKind::Compare { op, left, right } => SpannedExprKind::Compare {
                op: *op,
                left: Box::new(left.map_sheets(f)),
                right: Box::new(right.map_sheets(f)),
            },
            SpannedExprKind::FunctionCall { name, args } => SpannedExprKind::FunctionCall {
                name: name.clone(),
                args: args.iter().map(|a| a.map_sheets(f)).collect(),
            },
            SpannedExprKind::ImplicitIntersection(inner) => {
                SpannedExprKind::ImplicitIntersection(Box::new(inner.map_sheets(f)))
            }
            SpannedExprKind::SpillRange(inner) => {
                SpannedExprKind::SpillRange(Box::new(inner.map_sheets(f)))
            }
        };
        SpannedExpr {
            span: self.span,
            kind,
        }
    }
}

pub fn parse_spanned_formula(input: &str) -> Result<SpannedExpr<String>, FormulaParseError> {
    let mut lexer = Lexer::new(input);
    let tokens = lexer.tokenize()?;
    let mut p = ParserImpl::new(tokens);
    let expr = p.parse_formula()?;
    p.expect(TokenKind::End)?;
    Ok(expr)
}

pub(crate) fn evaluate_with_trace<R: crate::eval::ValueResolver>(
    resolver: &R,
    ctx: EvalContext,
    expr: &SpannedExpr<usize>,
) -> (Value, TraceNode) {
    let recalc_ctx = crate::eval::RecalcContext::new(0);
    let evaluator = TracedEvaluator {
        resolver,
        ctx,
        recalc_ctx: &recalc_ctx,
    };
    evaluator.eval_formula(expr)
}

#[derive(Debug, Clone, PartialEq)]
struct Token {
    kind: TokenKind,
    span: Span,
}

#[derive(Debug, Clone, PartialEq)]
enum TokenKind {
    Number(f64),
    String(String),
    Ident(String),
    SheetName(String),
    Error(ErrorKind),
    LBrace,
    RBrace,
    LParen,
    RParen,
    Comma,
    Semi,
    Colon,
    Bang,
    At,
    Hash,
    Plus,
    Minus,
    Star,
    Slash,
    Caret,
    Amp,
    Eq,
    Ne,
    Lt,
    Le,
    Gt,
    Ge,
    End,
}

struct Lexer<'a> {
    input: &'a str,
    pos: usize,
    prev_can_spill: bool,
}

impl<'a> Lexer<'a> {
    fn new(input: &'a str) -> Self {
        let mut pos = 0;
        while let Some(ch) = input[pos..].chars().next() {
            if ch.is_whitespace() {
                pos += ch.len_utf8();
            } else {
                break;
            }
        }
        if input[pos..].starts_with('=') {
            pos += 1;
        }
        Self {
            input,
            pos,
            prev_can_spill: false,
        }
    }

    fn tokenize(&mut self) -> Result<Vec<Token>, FormulaParseError> {
        let mut tokens = Vec::new();
        while let Some(ch) = self.peek_char() {
            if ch.is_whitespace() {
                self.pos += ch.len_utf8();
                continue;
            }

            let start = self.pos;
            let kind = match ch {
                '(' => {
                    self.pos += 1;
                    TokenKind::LParen
                }
                ')' => {
                    self.pos += 1;
                    TokenKind::RParen
                }
                '{' => {
                    self.pos += 1;
                    TokenKind::LBrace
                }
                '}' => {
                    self.pos += 1;
                    TokenKind::RBrace
                }
                ',' => {
                    self.pos += 1;
                    TokenKind::Comma
                }
                ';' => {
                    self.pos += 1;
                    TokenKind::Semi
                }
                ':' => {
                    self.pos += 1;
                    TokenKind::Colon
                }
                '!' => {
                    self.pos += 1;
                    TokenKind::Bang
                }
                '@' => {
                    self.pos += 1;
                    TokenKind::At
                }
                '+' => {
                    self.pos += 1;
                    TokenKind::Plus
                }
                '-' => {
                    self.pos += 1;
                    TokenKind::Minus
                }
                '*' => {
                    self.pos += 1;
                    TokenKind::Star
                }
                '/' => {
                    self.pos += 1;
                    TokenKind::Slash
                }
                '^' => {
                    self.pos += 1;
                    TokenKind::Caret
                }
                '&' => {
                    self.pos += 1;
                    TokenKind::Amp
                }
                '=' => {
                    self.pos += 1;
                    TokenKind::Eq
                }
                '<' => {
                    if self.peek_str("<=") {
                        self.pos += 2;
                        TokenKind::Le
                    } else if self.peek_str("<>") {
                        self.pos += 2;
                        TokenKind::Ne
                    } else {
                        self.pos += 1;
                        TokenKind::Lt
                    }
                }
                '>' => {
                    if self.peek_str(">=") {
                        self.pos += 2;
                        TokenKind::Ge
                    } else {
                        self.pos += 1;
                        TokenKind::Gt
                    }
                }
                '"' => self.lex_string()?,
                '\'' => self.lex_sheet_name()?,
                '#' => self.lex_hash_or_error()?,
                '.' | '0'..='9' => self.lex_number()?,
                _ if is_ident_start(ch) => self.lex_ident(),
                _ => {
                    return Err(FormulaParseError::UnexpectedToken(format!(
                        "unexpected character '{ch}'"
                    )))
                }
            };
            let span = Span::new(start, self.pos);
            let can_spill = matches!(&kind, TokenKind::Ident(_) | TokenKind::RParen);
            tokens.push(Token { kind, span });
            self.prev_can_spill = can_spill;
        }
        tokens.push(Token {
            kind: TokenKind::End,
            span: Span::new(self.pos, self.pos),
        });
        Ok(tokens)
    }

    fn peek_char(&self) -> Option<char> {
        self.input[self.pos..].chars().next()
    }

    fn peek_str(&self, s: &str) -> bool {
        self.input[self.pos..].starts_with(s)
    }

    fn lex_ident(&mut self) -> TokenKind {
        let start = self.pos;

        // External workbook prefixes (`[Book.xlsx]Sheet1!A1`) are treated as a single identifier
        // token, and the workbook portion inside `[...]` is more permissive than a normal Excel
        // identifier (it may contain spaces, dashes, etc). Mirror the canonical lexer by consuming
        // everything up to the closing `]` before switching back to strict identifier rules for
        // the sheet name portion.
        if self.peek_char() == Some('[') {
            self.pos += 1; // '['
            while let Some(ch) = self.peek_char() {
                self.pos += ch.len_utf8();
                if ch == ']' {
                    break;
                }
            }
        }

        while let Some(ch) = self.peek_char() {
            if ch.is_ascii_alphanumeric() || matches!(ch, '_' | '.' | '$') {
                self.pos += ch.len_utf8();
            } else {
                break;
            }
        }
        TokenKind::Ident(self.input[start..self.pos].to_string())
    }

    fn lex_number(&mut self) -> Result<TokenKind, FormulaParseError> {
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
        Ok(TokenKind::Number(n))
    }

    fn lex_string(&mut self) -> Result<TokenKind, FormulaParseError> {
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
        Ok(TokenKind::String(out))
    }

    fn lex_sheet_name(&mut self) -> Result<TokenKind, FormulaParseError> {
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
        Ok(TokenKind::SheetName(out))
    }

    fn lex_error(&mut self) -> Result<TokenKind, FormulaParseError> {
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
        Ok(TokenKind::Error(kind))
    }

    fn lex_hash_or_error(&mut self) -> Result<TokenKind, FormulaParseError> {
        // Spill-range operator is postfix (`A1#`), while error literals start with `#` (`#REF!`).
        let next = self.input[self.pos..].chars().nth(1);
        let looks_like_error = next
            .is_some_and(|c| c.is_ascii_alphanumeric() || matches!(c, '_' | '/' | '.' | '!' | '?'));
        if self.prev_can_spill && !looks_like_error {
            self.pos += 1;
            Ok(TokenKind::Hash)
        } else {
            self.lex_error()
        }
    }
}

fn is_ident_start(ch: char) -> bool {
    // Allow `[` for external workbook prefixes like `[Book.xlsx]Sheet1!A1`.
    ch.is_ascii_alphabetic() || matches!(ch, '_' | '$' | '[')
}

fn split_sheet_span_name(name: &str) -> Option<(String, String)> {
    let (start, end) = name.split_once(':')?;
    if start.is_empty() || end.is_empty() {
        return None;
    }
    Some((start.to_string(), end.to_string()))
}

struct ParserImpl {
    tokens: Vec<Token>,
    pos: usize,
}

impl ParserImpl {
    fn new(tokens: Vec<Token>) -> Self {
        Self { tokens, pos: 0 }
    }

    fn parse_formula(&mut self) -> Result<SpannedExpr<String>, FormulaParseError> {
        self.parse_compare()
    }

    fn parse_compare(&mut self) -> Result<SpannedExpr<String>, FormulaParseError> {
        let mut left = self.parse_concat()?;
        loop {
            let op = match self.peek().kind {
                TokenKind::Eq => CompareOp::Eq,
                TokenKind::Ne => CompareOp::Ne,
                TokenKind::Lt => CompareOp::Lt,
                TokenKind::Le => CompareOp::Le,
                TokenKind::Gt => CompareOp::Gt,
                TokenKind::Ge => CompareOp::Ge,
                _ => break,
            };
            self.next();
            let right = self.parse_concat()?;
            let span = Span::new(left.span.start, right.span.end);
            left = SpannedExpr {
                span,
                kind: SpannedExprKind::Compare {
                    op,
                    left: Box::new(left),
                    right: Box::new(right),
                },
            };
        }
        Ok(left)
    }

    fn parse_concat(&mut self) -> Result<SpannedExpr<String>, FormulaParseError> {
        let mut left = self.parse_add_sub()?;
        loop {
            if !matches!(self.peek().kind, TokenKind::Amp) {
                break;
            }
            self.next();
            let right = self.parse_add_sub()?;
            let span = Span::new(left.span.start, right.span.end);
            left = SpannedExpr {
                span,
                kind: SpannedExprKind::Binary {
                    op: crate::eval::BinaryOp::Concat,
                    left: Box::new(left),
                    right: Box::new(right),
                },
            };
        }
        Ok(left)
    }

    fn parse_add_sub(&mut self) -> Result<SpannedExpr<String>, FormulaParseError> {
        let mut left = self.parse_mul_div()?;
        loop {
            let op = match self.peek().kind {
                TokenKind::Plus => crate::eval::BinaryOp::Add,
                TokenKind::Minus => crate::eval::BinaryOp::Sub,
                _ => break,
            };
            self.next();
            let right = self.parse_mul_div()?;
            let span = Span::new(left.span.start, right.span.end);
            left = SpannedExpr {
                span,
                kind: SpannedExprKind::Binary {
                    op,
                    left: Box::new(left),
                    right: Box::new(right),
                },
            };
        }
        Ok(left)
    }

    fn parse_mul_div(&mut self) -> Result<SpannedExpr<String>, FormulaParseError> {
        let mut left = self.parse_unary()?;
        loop {
            let op = match self.peek().kind {
                TokenKind::Star => crate::eval::BinaryOp::Mul,
                TokenKind::Slash => crate::eval::BinaryOp::Div,
                _ => break,
            };
            self.next();
            let right = self.parse_unary()?;
            let span = Span::new(left.span.start, right.span.end);
            left = SpannedExpr {
                span,
                kind: SpannedExprKind::Binary {
                    op,
                    left: Box::new(left),
                    right: Box::new(right),
                },
            };
        }
        Ok(left)
    }

    fn parse_power(&mut self) -> Result<SpannedExpr<String>, FormulaParseError> {
        let left = self.parse_primary()?;
        if matches!(self.peek().kind, TokenKind::Caret) {
            self.next();
            // Excel exponentiation is right-associative and binds tighter than unary.
            let right = self.parse_unary()?;
            let span = Span::new(left.span.start, right.span.end);
            return Ok(SpannedExpr {
                span,
                kind: SpannedExprKind::Binary {
                    op: crate::eval::BinaryOp::Pow,
                    left: Box::new(left),
                    right: Box::new(right),
                },
            });
        }
        Ok(left)
    }

    fn parse_unary(&mut self) -> Result<SpannedExpr<String>, FormulaParseError> {
        match self.peek().kind {
            TokenKind::Plus => {
                let tok = self.next();
                let expr = self.parse_unary()?;
                Ok(SpannedExpr {
                    span: Span::new(tok.span.start, expr.span.end),
                    kind: SpannedExprKind::Unary {
                        op: UnaryOp::Plus,
                        expr: Box::new(expr),
                    },
                })
            }
            TokenKind::Minus => {
                let tok = self.next();
                let expr = self.parse_unary()?;
                Ok(SpannedExpr {
                    span: Span::new(tok.span.start, expr.span.end),
                    kind: SpannedExprKind::Unary {
                        op: UnaryOp::Minus,
                        expr: Box::new(expr),
                    },
                })
            }
            TokenKind::At => {
                let tok = self.next();
                let expr = self.parse_unary()?;
                Ok(SpannedExpr {
                    span: Span::new(tok.span.start, expr.span.end),
                    kind: SpannedExprKind::ImplicitIntersection(Box::new(expr)),
                })
            }
            _ => self.parse_power(),
        }
    }

    fn parse_primary(&mut self) -> Result<SpannedExpr<String>, FormulaParseError> {
        let tok = self.peek().clone();
        let mut expr = match &tok.kind {
            TokenKind::Number(n) => {
                self.next();
                Ok(SpannedExpr {
                    span: tok.span,
                    kind: SpannedExprKind::Number(*n),
                })
            }
            TokenKind::String(s) => {
                self.next();
                Ok(SpannedExpr {
                    span: tok.span,
                    kind: SpannedExprKind::Text(s.clone()),
                })
            }
            TokenKind::Error(e) => {
                self.next();
                Ok(SpannedExpr {
                    span: tok.span,
                    kind: SpannedExprKind::Error(*e),
                })
            }
            TokenKind::Ident(id) => {
                if matches!(self.peek_n(1).kind, TokenKind::LParen) {
                    self.parse_function_call()
                } else if matches!(self.peek_n(1).kind, TokenKind::Bang)
                    || (matches!(self.peek_n(1).kind, TokenKind::Colon)
                        && matches!(self.peek_n(3).kind, TokenKind::Bang))
                    || (id.starts_with('[')
                        && id.ends_with(']')
                        && matches!(
                            self.peek_n(1).kind,
                            TokenKind::Ident(_) | TokenKind::SheetName(_)
                        )
                        && (matches!(self.peek_n(2).kind, TokenKind::Bang)
                            || (matches!(self.peek_n(2).kind, TokenKind::Colon)
                                && matches!(self.peek_n(4).kind, TokenKind::Bang))))
                {
                    self.parse_sheet_ref()
                } else {
                    self.next();
                    match id.to_ascii_uppercase().as_str() {
                        "TRUE" => Ok(SpannedExpr {
                            span: tok.span,
                            kind: SpannedExprKind::Bool(true),
                        }),
                        "FALSE" => Ok(SpannedExpr {
                            span: tok.span,
                            kind: SpannedExprKind::Bool(false),
                        }),
                        _ => match parse_a1(id) {
                            Ok(addr) => self.parse_cell_or_range(
                                SheetReference::Current,
                                tok.span.start,
                                addr,
                                tok.span.end,
                            ),
                            Err(_) => Ok(SpannedExpr {
                                span: tok.span,
                                kind: SpannedExprKind::NameRef(crate::eval::NameRef {
                                    sheet: SheetReference::Current,
                                    name: id.clone(),
                                }),
                            }),
                        },
                    }
                }
            }
            TokenKind::SheetName(_name) => {
                if matches!(self.peek_n(1).kind, TokenKind::Bang)
                    || (matches!(self.peek_n(1).kind, TokenKind::Colon)
                        && matches!(self.peek_n(3).kind, TokenKind::Bang))
                {
                    self.parse_sheet_ref()
                } else {
                    self.next();
                    Ok(SpannedExpr {
                        span: tok.span,
                        kind: SpannedExprKind::Error(ErrorKind::Name),
                    })
                }
            }
            TokenKind::LParen => {
                let open = self.next();
                let expr = self.parse_compare()?;
                let close = self.expect(TokenKind::RParen)?;
                Ok(SpannedExpr {
                    span: Span::new(open.span.start, close.span.end),
                    kind: SpannedExprKind::Group(Box::new(expr)),
                })
            }
            TokenKind::LBrace => self.parse_array_literal(),
            other => Err(FormulaParseError::UnexpectedToken(format!("{other:?}"))),
        }?;

        while matches!(self.peek().kind, TokenKind::Hash) {
            let hash = self.next();
            expr = SpannedExpr {
                span: Span::new(expr.span.start, hash.span.end),
                kind: SpannedExprKind::SpillRange(Box::new(expr)),
            };
        }

        Ok(expr)
    }

    fn parse_array_literal(&mut self) -> Result<SpannedExpr<String>, FormulaParseError> {
        let open = self.expect(TokenKind::LBrace)?;
        let mut rows: Vec<Vec<SpannedExpr<String>>> = Vec::new();
        let mut current_row: Vec<SpannedExpr<String>> = Vec::new();
        let mut expecting_value = true;

        let blank_at = |pos: usize| SpannedExpr {
            span: Span::new(pos, pos),
            kind: SpannedExprKind::Blank,
        };

        loop {
            match &self.peek().kind {
                TokenKind::RBrace => {
                    let close = self.next();
                    if expecting_value && (!current_row.is_empty() || !rows.is_empty()) {
                        current_row.push(blank_at(close.span.start));
                    }
                    if !current_row.is_empty() || !rows.is_empty() {
                        rows.push(current_row);
                    }
                    return Ok(SpannedExpr {
                        span: Span::new(open.span.start, close.span.end),
                        kind: SpannedExprKind::ArrayLiteral { rows },
                    });
                }
                TokenKind::End => return Err(FormulaParseError::UnexpectedEof),
                TokenKind::Comma => {
                    // Blank element (e.g. `{1,,3}`).
                    let comma = self.next();
                    current_row.push(blank_at(comma.span.start));
                    expecting_value = true;
                    continue;
                }
                TokenKind::Semi => {
                    // Blank element at end of row (e.g. `{1,;2,3}`).
                    let semi = self.next();
                    current_row.push(blank_at(semi.span.start));
                    rows.push(current_row);
                    current_row = Vec::new();
                    expecting_value = true;
                    continue;
                }
                _ => {}
            }

            let el = self.parse_compare()?;
            expecting_value = false;
            current_row.push(el);

            match &self.peek().kind {
                TokenKind::Comma => {
                    self.next();
                    expecting_value = true;
                }
                TokenKind::Semi => {
                    self.next();
                    rows.push(current_row);
                    current_row = Vec::new();
                    expecting_value = true;
                }
                TokenKind::RBrace => {
                    // loop will close
                }
                TokenKind::End => return Err(FormulaParseError::UnexpectedEof),
                other => {
                    return Err(FormulaParseError::UnexpectedToken(format!(
                        "expected array separator or '}}', got {other:?}"
                    )))
                }
            }
        }
    }

    fn parse_function_call(&mut self) -> Result<SpannedExpr<String>, FormulaParseError> {
        let name_tok = self.next();
        let name = match name_tok.kind {
            TokenKind::Ident(s) => {
                let upper = s.to_ascii_uppercase();
                upper.strip_prefix("_XLFN.").unwrap_or(&upper).to_string()
            }
            other => {
                return Err(FormulaParseError::Expected {
                    expected: "identifier".to_string(),
                    got: format!("{other:?}"),
                })
            }
        };
        self.expect(TokenKind::LParen)?;
        let mut args = Vec::new();
        if !matches!(self.peek().kind, TokenKind::RParen) {
            loop {
                args.push(self.parse_compare()?);
                if matches!(self.peek().kind, TokenKind::Comma) {
                    self.next();
                    continue;
                }
                break;
            }
        }
        let close = self.expect(TokenKind::RParen)?;
        Ok(SpannedExpr {
            span: Span::new(name_tok.span.start, close.span.end),
            kind: SpannedExprKind::FunctionCall { name, args },
        })
    }

    fn parse_sheet_ref(&mut self) -> Result<SpannedExpr<String>, FormulaParseError> {
        let sheet_tok = self.next();
        let mut start_name = match sheet_tok.kind {
            TokenKind::Ident(s) | TokenKind::SheetName(s) => s,
            other => {
                return Err(FormulaParseError::Expected {
                    expected: "sheet name".to_string(),
                    got: format!("{other:?}"),
                })
            }
        };

        // External workbook references can be written with the workbook prefix unquoted and the
        // sheet name quoted separately: `[Book.xlsx]'My Sheet'!A1`.
        //
        // The canonical parser treats this as an external workbook sheet ref, but our debug lexer
        // tokenizes it as two tokens (`[Book.xlsx]` then `My Sheet`). Combine them so the rest of
        // the parser can operate on a single sheet name string.
        if start_name.starts_with('[')
            && start_name.ends_with(']')
            && matches!(self.peek().kind, TokenKind::Ident(_) | TokenKind::SheetName(_))
        {
            let sheet_name_tok = self.next();
            let sheet_name = match sheet_name_tok.kind {
                TokenKind::Ident(s) | TokenKind::SheetName(s) => s,
                other => {
                    return Err(FormulaParseError::Expected {
                        expected: "sheet name".to_string(),
                        got: format!("{other:?}"),
                    })
                }
            };
            start_name.push_str(&sheet_name);
        }

        let sheet = if matches!(self.peek().kind, TokenKind::Colon) {
            // Sheet span (3D ref) like `Sheet1:Sheet3!A1` / `'Sheet 1':'Sheet 3'!A1`.
            self.next(); // ':'
            let end_tok = self.next();
            let end_name = match end_tok.kind {
                TokenKind::Ident(s) | TokenKind::SheetName(s) => s,
                other => {
                    return Err(FormulaParseError::Expected {
                        expected: "sheet name".to_string(),
                        got: format!("{other:?}"),
                    })
                }
            };
            self.expect(TokenKind::Bang)?;
            if crate::eval::is_valid_external_sheet_key(&start_name) {
                // Excel treats `[Book]Sheet1:Sheet3!A1` as an external workbook 3D span where the
                // bracketed workbook prefix applies to both endpoints. We don't support external
                // workbook 3D spans today, but we still want debug tracing to behave like the main
                // engine for degenerate spans like `[Book]Sheet1:Sheet1!A1`.
                //
                // When the endpoint sheet names match, collapse to the single external sheet key
                // (`[Book]Sheet1`) so evaluation can consult the external provider.
                let Some((_, sheet_part)) = start_name.split_once(']') else {
                    return Ok(SpannedExpr {
                        span: Span::new(sheet_tok.span.start, end_tok.span.end),
                        kind: SpannedExprKind::Error(ErrorKind::Ref),
                    });
                };
                if sheet_part.eq_ignore_ascii_case(&end_name) {
                    SheetReference::External(start_name)
                } else {
                    // Preserve the full span in the sheet key so `resolve_sheet_id` reliably
                    // yields `#REF!`.
                    SheetReference::External(format!("{start_name}:{end_name}"))
                }
            } else {
                SheetReference::SheetRange(start_name, end_name)
            }
        } else {
            self.expect(TokenKind::Bang)?;
            match split_sheet_span_name(&start_name) {
                Some((start, end)) => {
                    if crate::eval::is_valid_external_sheet_key(&start) {
                        let Some((_, sheet_part)) = start.split_once(']') else {
                            return Ok(SpannedExpr {
                                span: Span::new(sheet_tok.span.start, sheet_tok.span.end),
                                kind: SpannedExprKind::Error(ErrorKind::Ref),
                            });
                        };
                        if sheet_part.eq_ignore_ascii_case(&end) {
                            SheetReference::External(start)
                        } else {
                            SheetReference::External(format!("{start}:{end}"))
                        }
                    } else {
                        SheetReference::SheetRange(start, end)
                    }
                }
                None => SheetReference::Sheet(start_name),
            }
        };

        if matches!(self.peek().kind, TokenKind::Ident(ref id) if id.starts_with('[')) {
            self.next();
            return Ok(SpannedExpr {
                span: Span::new(sheet_tok.span.start, sheet_tok.span.end),
                kind: SpannedExprKind::Error(ErrorKind::Ref),
            });
        }

        let addr_tok = self.next();
        let addr_str = match addr_tok.kind {
            TokenKind::Ident(s) => s,
            other => {
                return Err(FormulaParseError::Expected {
                    expected: "cell address".to_string(),
                    got: format!("{other:?}"),
                })
            }
        };
        match parse_a1(&addr_str) {
            Ok(addr) => {
                self.parse_cell_or_range(sheet, sheet_tok.span.start, addr, addr_tok.span.end)
            }
            Err(_) => Ok(SpannedExpr {
                span: Span::new(sheet_tok.span.start, addr_tok.span.end),
                kind: SpannedExprKind::NameRef(crate::eval::NameRef {
                    sheet,
                    name: addr_str,
                }),
            }),
        }
    }

    fn parse_cell_or_range(
        &mut self,
        sheet: SheetReference<String>,
        start_span: usize,
        start: CellAddr,
        end_span: usize,
    ) -> Result<SpannedExpr<String>, FormulaParseError> {
        if matches!(self.peek().kind, TokenKind::Colon) {
            self.next();
            let end_tok = self.next();
            let end_str = match end_tok.kind {
                TokenKind::Ident(s) => s,
                other => {
                    return Err(FormulaParseError::Expected {
                        expected: "cell address".to_string(),
                        got: format!("{other:?}"),
                    })
                }
            };
            let end = parse_a1(&end_str)?;
            Ok(SpannedExpr {
                span: Span::new(start_span, end_tok.span.end),
                kind: SpannedExprKind::RangeRef(crate::eval::RangeRef { sheet, start, end }),
            })
        } else {
            Ok(SpannedExpr {
                span: Span::new(start_span, end_span),
                kind: SpannedExprKind::CellRef(crate::eval::CellRef { sheet, addr: start }),
            })
        }
    }

    fn peek(&self) -> &Token {
        self.tokens
            .get(self.pos)
            .unwrap_or_else(|| self.tokens.last().unwrap())
    }

    fn peek_n(&self, n: usize) -> &Token {
        self.tokens
            .get(self.pos + n)
            .unwrap_or_else(|| self.tokens.last().unwrap())
    }

    fn next(&mut self) -> Token {
        let tok = self.peek().clone();
        self.pos += 1;
        tok
    }

    fn expect(&mut self, expected: TokenKind) -> Result<Token, FormulaParseError> {
        let got = self.next();
        if got.kind == expected {
            Ok(got)
        } else {
            Err(FormulaParseError::Expected {
                expected: format!("{expected:?}"),
                got: format!("{:?}", got.kind),
            })
        }
    }
}

#[derive(Debug, Clone)]
struct ResolvedRange {
    sheet_id: FnSheetId,
    start: CellAddr,
    end: CellAddr,
}

impl ResolvedRange {
    fn normalized(&self) -> Self {
        let (r1, r2) = if self.start.row <= self.end.row {
            (self.start.row, self.end.row)
        } else {
            (self.end.row, self.start.row)
        };
        let (c1, c2) = if self.start.col <= self.end.col {
            (self.start.col, self.end.col)
        } else {
            (self.end.col, self.start.col)
        };
        Self {
            sheet_id: self.sheet_id.clone(),
            start: CellAddr { row: r1, col: c1 },
            end: CellAddr { row: r2, col: c2 },
        }
    }

    fn is_single_cell(&self) -> bool {
        self.start == self.end
    }

    fn iter_cells(&self) -> impl Iterator<Item = CellAddr> {
        let norm = self.normalized();
        let rows = norm.start.row..=norm.end.row;
        let cols = norm.start.col..=norm.end.col;
        rows.flat_map(move |row| cols.clone().map(move |col| CellAddr { row, col }))
    }
}

#[derive(Debug, Clone)]
enum EvalValue {
    Scalar(Value),
    Reference(Vec<ResolvedRange>),
}

struct TracedEvaluator<'a, R: crate::eval::ValueResolver> {
    resolver: &'a R,
    ctx: EvalContext,
    recalc_ctx: &'a crate::eval::RecalcContext,
}

impl<'a, R: crate::eval::ValueResolver> TracedEvaluator<'a, R> {
    fn eval_formula(&self, expr: &SpannedExpr<usize>) -> (Value, TraceNode) {
        let (v, mut trace) = self.eval_value(expr);
        match v {
            EvalValue::Scalar(v) => (v, trace),
            EvalValue::Reference(ranges) => {
                let value = self.deref_reference_dynamic(ranges);
                trace.value = value.clone();
                (value, trace)
            }
        }
    }

    fn eval_scalar(&self, expr: &SpannedExpr<usize>) -> (Value, TraceNode) {
        let (v, mut trace) = self.eval_value(expr);
        match v {
            EvalValue::Scalar(v) => (v, trace),
            EvalValue::Reference(ranges) => {
                let scalar = self.deref_reference_scalar(&ranges);
                trace.value = scalar.clone();
                (scalar, trace)
            }
        }
    }

    fn deref_eval_value_dynamic(&self, value: EvalValue) -> Value {
        match value {
            EvalValue::Scalar(v) => v,
            EvalValue::Reference(ranges) => self.deref_reference_dynamic(ranges),
        }
    }

    fn deref_reference_dynamic(&self, ranges: Vec<ResolvedRange>) -> Value {
        match ranges.as_slice() {
            [] => Value::Error(ErrorKind::Ref),
            [only] => self.deref_reference_dynamic_single(only),
            _ => Value::Error(ErrorKind::Value),
        }
    }

    fn deref_reference_dynamic_single(&self, range: &ResolvedRange) -> Value {
        if range.is_single_cell() {
            return self.get_sheet_cell_value(&range.sheet_id, range.start);
        }
        let range = range.normalized();
        let rows = (range.end.row - range.start.row + 1) as usize;
        let cols = (range.end.col - range.start.col + 1) as usize;
        let mut values = Vec::with_capacity(rows.saturating_mul(cols));
        for row in range.start.row..=range.end.row {
            for col in range.start.col..=range.end.col {
                values.push(self.get_sheet_cell_value(&range.sheet_id, CellAddr { row, col }));
            }
        }
        Value::Array(Array::new(rows, cols, values))
    }

    fn get_sheet_cell_value(&self, sheet_id: &FnSheetId, addr: CellAddr) -> Value {
        match sheet_id {
            FnSheetId::Local(id) => self.resolver.get_cell_value(*id, addr),
            FnSheetId::External(key) => self
                .resolver
                .get_external_value(key, addr)
                .unwrap_or(Value::Error(ErrorKind::Ref)),
        }
    }

    fn eval_value(&self, expr: &SpannedExpr<usize>) -> (EvalValue, TraceNode) {
        match &expr.kind {
            SpannedExprKind::Number(n) => {
                let value = Value::Number(*n);
                (
                    EvalValue::Scalar(value.clone()),
                    TraceNode {
                        kind: TraceKind::Number,
                        span: expr.span,
                        value,
                        reference: None,
                        children: Vec::new(),
                    },
                )
            }
            SpannedExprKind::Text(s) => {
                let value = Value::Text(s.clone());
                (
                    EvalValue::Scalar(value.clone()),
                    TraceNode {
                        kind: TraceKind::Text,
                        span: expr.span,
                        value,
                        reference: None,
                        children: Vec::new(),
                    },
                )
            }
            SpannedExprKind::Bool(b) => {
                let value = Value::Bool(*b);
                (
                    EvalValue::Scalar(value.clone()),
                    TraceNode {
                        kind: TraceKind::Bool,
                        span: expr.span,
                        value,
                        reference: None,
                        children: Vec::new(),
                    },
                )
            }
            SpannedExprKind::Blank => (
                EvalValue::Scalar(Value::Blank),
                TraceNode {
                    kind: TraceKind::Blank,
                    span: expr.span,
                    value: Value::Blank,
                    reference: None,
                    children: Vec::new(),
                },
            ),
            SpannedExprKind::Error(e) => {
                let value = Value::Error(*e);
                (
                    EvalValue::Scalar(value.clone()),
                    TraceNode {
                        kind: TraceKind::Error,
                        span: expr.span,
                        value,
                        reference: None,
                        children: Vec::new(),
                    },
                )
            }
            SpannedExprKind::ArrayLiteral { rows } => {
                let row_count = rows.len();
                let col_count = rows.first().map(|r| r.len()).unwrap_or(0);

                if row_count == 0 || col_count == 0 || rows.iter().any(|r| r.len() != col_count) {
                    let value = Value::Error(ErrorKind::Value);
                    return (
                        EvalValue::Scalar(value.clone()),
                        TraceNode {
                            kind: TraceKind::ArrayLiteral {
                                rows: row_count,
                                cols: col_count,
                            },
                            span: expr.span,
                            value,
                            reference: None,
                            children: Vec::new(),
                        },
                    );
                }

                let mut children = Vec::with_capacity(row_count.saturating_mul(col_count));
                let mut out_values = Vec::with_capacity(row_count.saturating_mul(col_count));

                for row in rows {
                    for el in row {
                        let (ev, mut trace) = self.eval_value(el);
                        let v = match ev {
                            EvalValue::Scalar(v) => v,
                            EvalValue::Reference(ranges) => {
                                self.apply_implicit_intersection(&ranges)
                            }
                        };
                        let v = match v {
                            Value::Array(_) | Value::Spill { .. } => Value::Error(ErrorKind::Value),
                            other => other,
                        };
                        trace.value = v.clone();
                        out_values.push(v);
                        children.push(trace);
                    }
                }

                let value =
                    Value::Array(crate::value::Array::new(row_count, col_count, out_values));
                (
                    EvalValue::Scalar(value.clone()),
                    TraceNode {
                        kind: TraceKind::ArrayLiteral {
                            rows: row_count,
                            cols: col_count,
                        },
                        span: expr.span,
                        value,
                        reference: None,
                        children,
                    },
                )
            }
            SpannedExprKind::CellRef(r) => match self.resolve_sheet_ids(&r.sheet) {
                Some(sheet_ids)
                    if !sheet_ids.is_empty()
                        && sheet_ids.iter().all(|sheet_id| {
                            !matches!(sheet_id, FnSheetId::Local(id) if !self.resolver.sheet_exists(*id))
                        }) =>
                {
                    let reference = if sheet_ids.len() == 1 {
                        Some(TraceRef::Cell {
                            sheet: sheet_ids[0].clone(),
                            addr: r.addr,
                        })
                    } else {
                        None
                    };

                    (
                        EvalValue::Reference(
                            sheet_ids
                                .into_iter()
                                .map(|sheet_id| ResolvedRange {
                                    sheet_id,
                                    start: r.addr,
                                    end: r.addr,
                                })
                                .collect(),
                        ),
                        TraceNode {
                            kind: TraceKind::CellRef,
                            span: expr.span,
                            value: Value::Blank,
                            reference,
                            children: Vec::new(),
                        },
                    )
                }
                _ => {
                    let value = Value::Error(ErrorKind::Ref);
                    (
                        EvalValue::Scalar(value.clone()),
                        TraceNode {
                            kind: TraceKind::CellRef,
                            span: expr.span,
                            value,
                            reference: None,
                            children: Vec::new(),
                        },
                    )
                }
            },
            SpannedExprKind::RangeRef(r) => match self.resolve_sheet_ids(&r.sheet) {
                Some(sheet_ids)
                    if !sheet_ids.is_empty()
                        && sheet_ids.iter().all(|sheet_id| {
                            !matches!(sheet_id, FnSheetId::Local(id) if !self.resolver.sheet_exists(*id))
                        }) =>
                {
                    let reference = if sheet_ids.len() == 1 {
                        Some(TraceRef::Range {
                            sheet: sheet_ids[0].clone(),
                            start: r.start,
                            end: r.end,
                        })
                    } else {
                        None
                    };

                    (
                        EvalValue::Reference(
                            sheet_ids
                                .into_iter()
                                .map(|sheet_id| ResolvedRange {
                                    sheet_id,
                                    start: r.start,
                                    end: r.end,
                                })
                                .collect(),
                        ),
                        TraceNode {
                            kind: TraceKind::RangeRef,
                            span: expr.span,
                            value: Value::Blank,
                            reference,
                            children: Vec::new(),
                        },
                    )
                }
                _ => {
                    let value = Value::Error(ErrorKind::Ref);
                    (
                        EvalValue::Scalar(value.clone()),
                        TraceNode {
                            kind: TraceKind::RangeRef,
                            span: expr.span,
                            value,
                            reference: None,
                            children: Vec::new(),
                        },
                    )
                }
            },
            SpannedExprKind::NameRef(nref) => match self.resolve_sheet_id(&nref.sheet) {
                Some(FnSheetId::Local(sheet_id)) if self.resolver.sheet_exists(sheet_id) => {
                    let resolved = self.resolver.resolve_name(sheet_id, &nref.name);
                    match resolved {
                        Some(crate::eval::ResolvedName::Constant(v)) => (
                            EvalValue::Scalar(v.clone()),
                            TraceNode {
                                kind: TraceKind::NameRef {
                                    name: nref.name.clone(),
                                },
                                span: expr.span,
                                value: v,
                                reference: None,
                                children: Vec::new(),
                            },
                        ),
                        Some(crate::eval::ResolvedName::Expr(compiled)) => {
                            let evaluator = crate::eval::Evaluator::new(
                                self.resolver,
                                EvalContext {
                                    current_sheet: sheet_id,
                                    current_cell: self.ctx.current_cell,
                                },
                                self.recalc_ctx,
                            );
                            match FunctionContext::eval_arg(&evaluator, &compiled) {
                                FnArgValue::Scalar(v) => (
                                    EvalValue::Scalar(v.clone()),
                                    TraceNode {
                                        kind: TraceKind::NameRef {
                                            name: nref.name.clone(),
                                        },
                                        span: expr.span,
                                        value: v,
                                        reference: None,
                                        children: Vec::new(),
                                    },
                                ),
                                FnArgValue::Reference(r) => {
                                    let sheet_id = r.sheet_id.clone();
                                    let range = ResolvedRange {
                                        sheet_id: sheet_id.clone(),
                                        start: r.start,
                                        end: r.end,
                                    };
                                    let reference = if r.is_single_cell() {
                                        Some(TraceRef::Cell {
                                            sheet: sheet_id.clone(),
                                            addr: r.start,
                                        })
                                    } else {
                                        Some(TraceRef::Range {
                                            sheet: sheet_id.clone(),
                                            start: r.start,
                                            end: r.end,
                                        })
                                    };
                                    (
                                        EvalValue::Reference(vec![range]),
                                        TraceNode {
                                            kind: TraceKind::NameRef {
                                                name: nref.name.clone(),
                                            },
                                            span: expr.span,
                                            value: Value::Blank,
                                            reference,
                                            children: Vec::new(),
                                        },
                                    )
                                }
                                FnArgValue::ReferenceUnion(_) => {
                                    let value = Value::Error(ErrorKind::Value);
                                    (
                                        EvalValue::Scalar(value.clone()),
                                        TraceNode {
                                            kind: TraceKind::NameRef {
                                                name: nref.name.clone(),
                                            },
                                            span: expr.span,
                                            value,
                                            reference: None,
                                            children: Vec::new(),
                                        },
                                    )
                                }
                            }
                        }
                        None => {
                            let value = Value::Error(ErrorKind::Name);
                            (
                                EvalValue::Scalar(value.clone()),
                                TraceNode {
                                    kind: TraceKind::NameRef {
                                        name: nref.name.clone(),
                                    },
                                    span: expr.span,
                                    value,
                                    reference: None,
                                    children: Vec::new(),
                                },
                            )
                        }
                    }
                }
                _ => {
                    let value = Value::Error(ErrorKind::Ref);
                    (
                        EvalValue::Scalar(value.clone()),
                        TraceNode {
                            kind: TraceKind::NameRef {
                                name: nref.name.clone(),
                            },
                            span: expr.span,
                            value,
                            reference: None,
                            children: Vec::new(),
                        },
                    )
                }
            },
            SpannedExprKind::Group(inner) => {
                let (ev, child) = self.eval_value(inner);
                let (value, reference) = match &ev {
                    EvalValue::Scalar(v) => (v.clone(), None),
                    EvalValue::Reference(_) => (Value::Blank, child.reference.clone()),
                };
                (
                    ev,
                    TraceNode {
                        kind: TraceKind::Group,
                        span: expr.span,
                        value,
                        reference,
                        children: vec![child],
                    },
                )
            }
            SpannedExprKind::SpillRange(inner) => {
                let (ev, child) = self.eval_value(inner);
                let (out_ev, reference) = match ev {
                    EvalValue::Scalar(Value::Error(e)) => {
                        (EvalValue::Scalar(Value::Error(e)), None)
                    }
                    EvalValue::Scalar(_) => {
                        (EvalValue::Scalar(Value::Error(ErrorKind::Value)), None)
                    }
                    EvalValue::Reference(mut ranges) => {
                        // Spill-range references are only well-defined for a single-cell reference.
                        if ranges.len() != 1 {
                            (EvalValue::Scalar(Value::Error(ErrorKind::Value)), None)
                        } else {
                            let range = ranges.pop().expect("checked len() above");
                            if !range.is_single_cell() {
                                (EvalValue::Scalar(Value::Error(ErrorKind::Value)), None)
                            } else {
                                let addr = range.start;
                                match range.sheet_id {
                                    FnSheetId::Local(sheet_id) => {
                                        match self.resolver.spill_origin(sheet_id, addr) {
                                            Some(origin) => {
                                                match self.resolver.spill_range(sheet_id, origin) {
                                                    Some((start, end)) => {
                                                        let sheet = FnSheetId::Local(sheet_id);
                                                        (
                                                            EvalValue::Reference(vec![ResolvedRange {
                                                                sheet_id: sheet.clone(),
                                                                start,
                                                                end,
                                                            }]),
                                                            Some(TraceRef::Range { sheet, start, end }),
                                                        )
                                                    }
                                                    None => (
                                                        EvalValue::Scalar(Value::Error(ErrorKind::Ref)),
                                                        None,
                                                    ),
                                                }
                                            }
                                            None => (
                                                EvalValue::Scalar(Value::Error(ErrorKind::Ref)),
                                                None,
                                            ),
                                        }
                                    }
                                    FnSheetId::External(_) => (
                                        EvalValue::Scalar(Value::Error(ErrorKind::Ref)),
                                        None,
                                    ),
                                }
                            }
                        }
                    }
                };

                let value = match &out_ev {
                    EvalValue::Scalar(v) => v.clone(),
                    EvalValue::Reference(_) => Value::Blank,
                };

                (
                    out_ev,
                    TraceNode {
                        kind: TraceKind::SpillRange,
                        span: expr.span,
                        value,
                        reference,
                        children: vec![child],
                    },
                )
            }
            SpannedExprKind::Unary { op, expr: inner } => {
                let (ev, child) = self.eval_value(inner);
                let value = self.deref_eval_value_dynamic(ev);
                let locale = self.recalc_ctx.number_locale;
                let out = elementwise_unary(&value, |elem| numeric_unary(*op, elem, locale));
                (
                    EvalValue::Scalar(out.clone()),
                    TraceNode {
                        kind: TraceKind::Unary { op: *op },
                        span: expr.span,
                        value: out,
                        reference: None,
                        children: vec![child],
                    },
                )
            }
            SpannedExprKind::Binary { op, left, right } => {
                let (l_ev, ltrace) = self.eval_value(left);
                let (r_ev, rtrace) = self.eval_value(right);

                let l = self.deref_eval_value_dynamic(l_ev);
                let r = self.deref_eval_value_dynamic(r_ev);
                let locale = self.recalc_ctx.number_locale;

                let out = match op {
                    crate::eval::BinaryOp::Add
                    | crate::eval::BinaryOp::Sub
                    | crate::eval::BinaryOp::Mul
                    | crate::eval::BinaryOp::Div
                    | crate::eval::BinaryOp::Pow => {
                        elementwise_binary(&l, &r, |a, b| numeric_binary(*op, a, b, locale))
                    }
                    crate::eval::BinaryOp::Concat => elementwise_binary(&l, &r, concat_binary),
                    crate::eval::BinaryOp::Range
                    | crate::eval::BinaryOp::Intersect
                    | crate::eval::BinaryOp::Union => Value::Error(ErrorKind::Value),
                };
                (
                    EvalValue::Scalar(out.clone()),
                    TraceNode {
                        kind: TraceKind::Binary { op: *op },
                        span: expr.span,
                        value: out,
                        reference: None,
                        children: vec![ltrace, rtrace],
                    },
                )
            }
            SpannedExprKind::Compare { op, left, right } => {
                let (l_ev, ltrace) = self.eval_value(left);
                let (r_ev, rtrace) = self.eval_value(right);

                let l = self.deref_eval_value_dynamic(l_ev);
                let r = self.deref_eval_value_dynamic(r_ev);
                let out = elementwise_binary(&l, &r, |a, b| excel_compare(a, b, *op));
                (
                    EvalValue::Scalar(out.clone()),
                    TraceNode {
                        kind: TraceKind::Compare { op: *op },
                        span: expr.span,
                        value: out,
                        reference: None,
                        children: vec![ltrace, rtrace],
                    },
                )
            }
            SpannedExprKind::FunctionCall { name, args } => {
                let (out, children) = self.eval_function(name, args);
                (
                    EvalValue::Scalar(out.clone()),
                    TraceNode {
                        kind: TraceKind::FunctionCall { name: name.clone() },
                        span: expr.span,
                        value: out,
                        reference: None,
                        children,
                    },
                )
            }
            SpannedExprKind::ImplicitIntersection(inner) => {
                let (v, child) = self.eval_value(inner);
                let out = match v {
                    EvalValue::Scalar(v) => v,
                    EvalValue::Reference(ranges) => self.apply_implicit_intersection(&ranges),
                };
                (
                    EvalValue::Scalar(out.clone()),
                    TraceNode {
                        kind: TraceKind::ImplicitIntersection,
                        span: expr.span,
                        value: out,
                        reference: None,
                        children: vec![child],
                    },
                )
            }
        }
    }

    fn resolve_sheet_id(&self, sheet: &SheetReference<usize>) -> Option<FnSheetId> {
        match sheet {
            SheetReference::Current => Some(FnSheetId::Local(self.ctx.current_sheet)),
            SheetReference::Sheet(id) => Some(FnSheetId::Local(*id)),
            SheetReference::SheetRange(a, b) => {
                if a == b {
                    Some(FnSheetId::Local(*a))
                } else {
                    None
                }
            }
            SheetReference::External(key) => crate::eval::is_valid_external_sheet_key(key)
                .then(|| FnSheetId::External(key.clone())),
        }
    }

    fn resolve_sheet_ids(&self, sheet: &SheetReference<usize>) -> Option<Vec<FnSheetId>> {
        match sheet {
            SheetReference::Current => Some(vec![FnSheetId::Local(self.ctx.current_sheet)]),
            SheetReference::Sheet(id) => Some(vec![FnSheetId::Local(*id)]),
            SheetReference::SheetRange(a, b) => {
                let (start, end) = if a <= b { (*a, *b) } else { (*b, *a) };
                Some((start..=end).map(FnSheetId::Local).collect())
            }
            SheetReference::External(key) => crate::eval::is_valid_external_sheet_key(key)
                .then(|| vec![FnSheetId::External(key.clone())]),
        }
    }

    fn deref_reference_scalar(&self, ranges: &[ResolvedRange]) -> Value {
        match ranges {
            [only] if only.is_single_cell() => self.get_sheet_cell_value(&only.sheet_id, only.start),
            [_only] => Value::Error(ErrorKind::Spill),
            _ => Value::Error(ErrorKind::Value),
        }
    }

    fn apply_implicit_intersection(&self, ranges: &[ResolvedRange]) -> Value {
        match ranges {
            [] => Value::Error(ErrorKind::Value),
            [only] => self.apply_implicit_intersection_single(only),
            many => {
                // If multiple areas intersect, Excel's implicit intersection is ambiguous. We
                // approximate by succeeding only when exactly one area intersects.
                let mut hits = Vec::new();
                for r in many {
                    let v = self.apply_implicit_intersection_single(r);
                    if !matches!(v, Value::Error(ErrorKind::Value)) {
                        hits.push(v);
                    }
                }
                match hits.as_slice() {
                    [only] => only.clone(),
                    _ => Value::Error(ErrorKind::Value),
                }
            }
        }
    }

    fn apply_implicit_intersection_single(&self, range: &ResolvedRange) -> Value {
        if range.is_single_cell() {
            return self.get_sheet_cell_value(&range.sheet_id, range.start);
        }

        let range = range.normalized();
        let cur = self.ctx.current_cell;

        if range.start.col == range.end.col {
            if cur.row >= range.start.row && cur.row <= range.end.row {
                return self.get_sheet_cell_value(
                    &range.sheet_id,
                    CellAddr {
                        row: cur.row,
                        col: range.start.col,
                    },
                );
            }
            return Value::Error(ErrorKind::Value);
        }
        if range.start.row == range.end.row {
            if cur.col >= range.start.col && cur.col <= range.end.col {
                return self.get_sheet_cell_value(
                    &range.sheet_id,
                    CellAddr {
                        row: range.start.row,
                        col: cur.col,
                    },
                );
            }
            return Value::Error(ErrorKind::Value);
        }

        if cur.row >= range.start.row
            && cur.row <= range.end.row
            && cur.col >= range.start.col
            && cur.col <= range.end.col
        {
            return self.get_sheet_cell_value(&range.sheet_id, cur);
        }

        Value::Error(ErrorKind::Value)
    }

    fn eval_function(&self, name: &str, args: &[SpannedExpr<usize>]) -> (Value, Vec<TraceNode>) {
        match name {
            "IF" => self.fn_if(args),
            "IFERROR" => self.fn_iferror(args),
            "ISERROR" => self.fn_iserror(args),
            "SUM" => self.fn_sum(args),
            "VLOOKUP" => self.fn_vlookup(args),
            _ => (Value::Error(ErrorKind::Name), Vec::new()),
        }
    }

    fn fn_if(&self, args: &[SpannedExpr<usize>]) -> (Value, Vec<TraceNode>) {
        if args.is_empty() {
            return (Value::Error(ErrorKind::Value), Vec::new());
        }
        let (cond_val, cond_trace) = self.eval_scalar(&args[0]);
        if let Value::Error(e) = cond_val {
            return (Value::Error(e), vec![cond_trace]);
        }
        let cond = match cond_val.coerce_to_bool() {
            Ok(b) => b,
            Err(e) => return (Value::Error(e), vec![cond_trace]),
        };

        if cond {
            if args.len() >= 2 {
                let (v, trace) = self.eval_scalar(&args[1]);
                (v, vec![cond_trace, trace])
            } else {
                (Value::Bool(true), vec![cond_trace])
            }
        } else if args.len() >= 3 {
            let (v, trace) = self.eval_scalar(&args[2]);
            (v, vec![cond_trace, trace])
        } else {
            (Value::Bool(false), vec![cond_trace])
        }
    }

    fn fn_iferror(&self, args: &[SpannedExpr<usize>]) -> (Value, Vec<TraceNode>) {
        if args.len() < 2 {
            return (Value::Error(ErrorKind::Value), Vec::new());
        }
        let (first, first_trace) = self.eval_scalar(&args[0]);
        match first {
            Value::Error(_) => {
                let (fallback, fallback_trace) = self.eval_scalar(&args[1]);
                (fallback, vec![first_trace, fallback_trace])
            }
            other => (other, vec![first_trace]),
        }
    }

    fn fn_iserror(&self, args: &[SpannedExpr<usize>]) -> (Value, Vec<TraceNode>) {
        if args.len() != 1 {
            return (Value::Error(ErrorKind::Value), Vec::new());
        }
        let (v, trace) = self.eval_scalar(&args[0]);
        (Value::Bool(matches!(v, Value::Error(_))), vec![trace])
    }

    fn fn_sum(&self, args: &[SpannedExpr<usize>]) -> (Value, Vec<TraceNode>) {
        let mut acc = 0.0;
        let mut traces = Vec::new();

        for arg in args {
            let (ev, trace) = self.eval_value(arg);
            traces.push(trace);
            match ev {
                EvalValue::Scalar(v) => match v {
                    Value::Error(e) => return (Value::Error(e), traces),
                    Value::Number(n) => acc += n,
                    Value::Bool(b) => acc += if b { 1.0 } else { 0.0 },
                    Value::Blank => {}
                    Value::Text(s) => {
                        let n = match Value::Text(s).coerce_to_number() {
                            Ok(n) => n,
                            Err(e) => return (Value::Error(e), traces),
                        };
                        acc += n;
                    }
                    Value::Reference(_) | Value::ReferenceUnion(_) => {
                        return (Value::Error(ErrorKind::Value), traces);
                    }
                    Value::Array(arr) => {
                        for v in arr.iter() {
                            match v {
                                Value::Error(e) => return (Value::Error(*e), traces),
                                Value::Number(n) => acc += n,
                                Value::Bool(_)
                                | Value::Text(_)
                                | Value::Blank
                                | Value::Array(_)
                                | Value::Lambda(_)
                                | Value::Spill { .. }
                                | Value::Reference(_)
                                | Value::ReferenceUnion(_) => {}
                            }
                        }
                    }
                    Value::Lambda(_) => return (Value::Error(ErrorKind::Value), traces),
                    Value::Spill { .. } => return (Value::Error(ErrorKind::Value), traces),
                },
                EvalValue::Reference(ranges) => {
                    for range in ranges {
                        for addr in range.iter_cells() {
                            let v = self.get_sheet_cell_value(&range.sheet_id, addr);
                            match v {
                                Value::Error(e) => return (Value::Error(e), traces),
                                Value::Number(n) => acc += n,
                                Value::Bool(_)
                                | Value::Text(_)
                                | Value::Blank
                                | Value::Array(_)
                                | Value::Lambda(_)
                                | Value::Spill { .. }
                                | Value::Reference(_)
                                | Value::ReferenceUnion(_) => {}
                            }
                        }
                    }
                }
            }
        }

        (Value::Number(acc), traces)
    }

    fn fn_vlookup(&self, args: &[SpannedExpr<usize>]) -> (Value, Vec<TraceNode>) {
        if args.len() < 3 || args.len() > 4 {
            return (Value::Error(ErrorKind::Value), Vec::new());
        }

        let mut traces = Vec::new();

        let (lookup_value, lookup_trace) = self.eval_scalar(&args[0]);
        traces.push(lookup_trace);
        if let Value::Error(e) = lookup_value {
            return (Value::Error(e), traces);
        }

        let (table_ev, table_trace) = self.eval_value(&args[1]);
        traces.push(table_trace);
        let table_range = match table_ev {
            EvalValue::Reference(mut ranges) => match ranges.as_mut_slice() {
                [only] => only.normalized(),
                _ => return (Value::Error(ErrorKind::Value), traces),
            },
            EvalValue::Scalar(Value::Error(e)) => return (Value::Error(e), traces),
            EvalValue::Scalar(_) => return (Value::Error(ErrorKind::Value), traces),
        };

        let (col_index_val, col_trace) = self.eval_scalar(&args[2]);
        traces.push(col_trace);
        if let Value::Error(e) = col_index_val {
            return (Value::Error(e), traces);
        }
        let col_index_num = match col_index_val.coerce_to_number() {
            Ok(n) => n,
            Err(e) => return (Value::Error(e), traces),
        };
        let col_index = col_index_num as i64;
        if col_index <= 0 {
            return (Value::Error(ErrorKind::Value), traces);
        }

        // Optional `range_lookup` argument. Excel defaults to TRUE (approx match).
        let range_lookup = if args.len() == 4 {
            let (v, trace) = self.eval_scalar(&args[3]);
            traces.push(trace);
            if let Value::Error(e) = v {
                return (Value::Error(e), traces);
            }
            match v.coerce_to_bool() {
                Ok(b) => b,
                Err(e) => return (Value::Error(e), traces),
            }
        } else {
            true
        };
        let exact = !range_lookup;

        let width = (table_range.end.col - table_range.start.col + 1) as i64;
        if col_index > width {
            return (Value::Error(ErrorKind::Ref), traces);
        }
        let target_col = table_range.start.col + (col_index as u32) - 1;

        if exact {
            for row in table_range.start.row..=table_range.end.row {
                let key = CellAddr {
                    row,
                    col: table_range.start.col,
                };
                let candidate = self.get_sheet_cell_value(&table_range.sheet_id, key);
                if matches!(candidate, Value::Error(_)) {
                    continue;
                }
                let is_match = excel_order(&candidate, &lookup_value)
                    .map(|o| o == Ordering::Equal)
                    .unwrap_or(false);
                if is_match {
                    let result_addr = CellAddr {
                        row,
                        col: target_col,
                    };
                    return (
                        self.get_sheet_cell_value(&table_range.sheet_id, result_addr),
                        traces,
                    );
                }
            }
            (Value::Error(ErrorKind::NA), traces)
        } else {
            let mut best_row: Option<u32> = None;
            for row in table_range.start.row..=table_range.end.row {
                let key = CellAddr {
                    row,
                    col: table_range.start.col,
                };
                let candidate = self.get_sheet_cell_value(&table_range.sheet_id, key);
                if matches!(candidate, Value::Error(_)) {
                    continue;
                }
                let ord = match excel_order(&candidate, &lookup_value) {
                    Ok(o) => o,
                    Err(_) => continue,
                };
                if ord != Ordering::Greater {
                    best_row = Some(row);
                }
            }
            if let Some(row) = best_row {
                let result_addr = CellAddr {
                    row,
                    col: target_col,
                };
                (
                    self.get_sheet_cell_value(&table_range.sheet_id, result_addr),
                    traces,
                )
            } else {
                (Value::Error(ErrorKind::NA), traces)
            }
        }
    }
}

fn numeric_unary(op: UnaryOp, value: &Value, locale: NumberLocale) -> Value {
    match value {
        Value::Error(e) => Value::Error(*e),
        other => {
            let n = match other.coerce_to_number_with_locale(locale) {
                Ok(n) => n,
                Err(e) => return Value::Error(e),
            };
            match op {
                UnaryOp::Plus => Value::Number(n),
                UnaryOp::Minus => Value::Number(-n),
            }
        }
    }
}

fn concat_binary(left: &Value, right: &Value) -> Value {
    if let Value::Error(e) = left {
        return Value::Error(*e);
    }
    if let Value::Error(e) = right {
        return Value::Error(*e);
    }

    let ls = match left.coerce_to_string() {
        Ok(s) => s,
        Err(e) => return Value::Error(e),
    };
    let rs = match right.coerce_to_string() {
        Ok(s) => s,
        Err(e) => return Value::Error(e),
    };
    Value::Text(format!("{ls}{rs}"))
}

fn numeric_binary(op: crate::eval::BinaryOp, left: &Value, right: &Value, locale: NumberLocale) -> Value {
    if let Value::Error(e) = left {
        return Value::Error(*e);
    }
    if let Value::Error(e) = right {
        return Value::Error(*e);
    }

    let ln = match left.coerce_to_number_with_locale(locale) {
        Ok(n) => n,
        Err(e) => return Value::Error(e),
    };
    let rn = match right.coerce_to_number_with_locale(locale) {
        Ok(n) => n,
        Err(e) => return Value::Error(e),
    };

    match op {
        crate::eval::BinaryOp::Add => Value::Number(ln + rn),
        crate::eval::BinaryOp::Sub => Value::Number(ln - rn),
        crate::eval::BinaryOp::Mul => Value::Number(ln * rn),
        crate::eval::BinaryOp::Div => {
            if rn == 0.0 {
                Value::Error(ErrorKind::Div0)
            } else {
                Value::Number(ln / rn)
            }
        }
        crate::eval::BinaryOp::Pow => match crate::functions::math::power(ln, rn) {
            Ok(n) => Value::Number(n),
            Err(e) => Value::Error(match e {
                ExcelError::Div0 => ErrorKind::Div0,
                ExcelError::Value => ErrorKind::Value,
                ExcelError::Num => ErrorKind::Num,
            }),
        },
        _ => Value::Error(ErrorKind::Value),
    }
}

fn elementwise_unary(value: &Value, f: impl Fn(&Value) -> Value) -> Value {
    match value {
        Value::Array(arr) => Value::Array(Array::new(arr.rows, arr.cols, arr.iter().map(f).collect())),
        other => f(other),
    }
}

fn elementwise_binary(left: &Value, right: &Value, f: impl Fn(&Value, &Value) -> Value) -> Value {
    match (left, right) {
        (Value::Array(left_arr), Value::Array(right_arr)) => {
            let out_rows = if left_arr.rows == right_arr.rows {
                left_arr.rows
            } else if left_arr.rows == 1 {
                right_arr.rows
            } else if right_arr.rows == 1 {
                left_arr.rows
            } else {
                return Value::Error(ErrorKind::Value);
            };

            let out_cols = if left_arr.cols == right_arr.cols {
                left_arr.cols
            } else if left_arr.cols == 1 {
                right_arr.cols
            } else if right_arr.cols == 1 {
                left_arr.cols
            } else {
                return Value::Error(ErrorKind::Value);
            };

            let mut out = Vec::with_capacity(out_rows.saturating_mul(out_cols));
            for row in 0..out_rows {
                let l_row = if left_arr.rows == 1 { 0 } else { row };
                let r_row = if right_arr.rows == 1 { 0 } else { row };
                for col in 0..out_cols {
                    let l_col = if left_arr.cols == 1 { 0 } else { col };
                    let r_col = if right_arr.cols == 1 { 0 } else { col };
                    let l = left_arr.get(l_row, l_col).unwrap_or(&Value::Blank);
                    let r = right_arr.get(r_row, r_col).unwrap_or(&Value::Blank);
                    out.push(f(l, r));
                }
            }
            Value::Array(Array::new(out_rows, out_cols, out))
        }
        (Value::Array(left_arr), right_scalar) => Value::Array(Array::new(
            left_arr.rows,
            left_arr.cols,
            left_arr.values.iter().map(|a| f(a, right_scalar)).collect(),
        )),
        (left_scalar, Value::Array(right_arr)) => Value::Array(Array::new(
            right_arr.rows,
            right_arr.cols,
            right_arr.values.iter().map(|b| f(left_scalar, b)).collect(),
        )),
        (left_scalar, right_scalar) => f(left_scalar, right_scalar),
    }
}

fn excel_compare(left: &Value, right: &Value, op: CompareOp) -> Value {
    let ord = match excel_order(left, right) {
        Ok(ord) => ord,
        Err(e) => return Value::Error(e),
    };

    let result = match op {
        CompareOp::Eq => ord == Ordering::Equal,
        CompareOp::Ne => ord != Ordering::Equal,
        CompareOp::Lt => ord == Ordering::Less,
        CompareOp::Le => ord != Ordering::Greater,
        CompareOp::Gt => ord == Ordering::Greater,
        CompareOp::Ge => ord != Ordering::Less,
    };

    Value::Bool(result)
}

fn excel_order(left: &Value, right: &Value) -> Result<Ordering, ErrorKind> {
    if let Value::Error(e) = left {
        return Err(*e);
    }
    if let Value::Error(e) = right {
        return Err(*e);
    }
    if matches!(
        left,
        Value::Array(_)
            | Value::Lambda(_)
            | Value::Spill { .. }
            | Value::Reference(_)
            | Value::ReferenceUnion(_)
    ) || matches!(
        right,
        Value::Array(_)
            | Value::Lambda(_)
            | Value::Spill { .. }
            | Value::Reference(_)
            | Value::ReferenceUnion(_)
    ) {
        return Err(ErrorKind::Value);
    }

    let (l, r) = match (left, right) {
        (Value::Blank, Value::Number(_)) => (Value::Number(0.0), right.clone()),
        (Value::Number(_), Value::Blank) => (left.clone(), Value::Number(0.0)),
        (Value::Blank, Value::Bool(_)) => (Value::Bool(false), right.clone()),
        (Value::Bool(_), Value::Blank) => (left.clone(), Value::Bool(false)),
        (Value::Blank, Value::Text(_)) => (Value::Text(String::new()), right.clone()),
        (Value::Text(_), Value::Blank) => (left.clone(), Value::Text(String::new())),
        _ => (left.clone(), right.clone()),
    };

    Ok(match (&l, &r) {
        (Value::Number(a), Value::Number(b)) => a.partial_cmp(b).unwrap_or(Ordering::Equal),
        (Value::Text(a), Value::Text(b)) => {
            let au = a.to_ascii_uppercase();
            let bu = b.to_ascii_uppercase();
            au.cmp(&bu)
        }
        (Value::Bool(a), Value::Bool(b)) => a.cmp(b),
        (Value::Number(_), Value::Text(_) | Value::Bool(_)) => Ordering::Less,
        (Value::Text(_), Value::Bool(_)) => Ordering::Less,
        (Value::Text(_), Value::Number(_)) => Ordering::Greater,
        (Value::Bool(_), Value::Number(_) | Value::Text(_)) => Ordering::Greater,
        (Value::Blank, Value::Blank) => Ordering::Equal,
        (Value::Blank, _) => Ordering::Less,
        (_, Value::Blank) => Ordering::Greater,
        (Value::Error(_), _) | (_, Value::Error(_)) => Ordering::Equal,
        (Value::Array(_), _)
        | (_, Value::Array(_))
        | (Value::Lambda(_), _)
        | (_, Value::Lambda(_))
        | (Value::Spill { .. }, _)
        | (_, Value::Spill { .. })
        | (Value::Reference(_), _)
        | (_, Value::Reference(_))
        | (Value::ReferenceUnion(_), _)
        | (_, Value::ReferenceUnion(_)) => Ordering::Equal,
    })
}
