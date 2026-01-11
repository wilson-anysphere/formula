use crate::eval::{parse_a1, CellAddr, CompareOp, EvalContext, FormulaParseError, SheetReference, UnaryOp};
use crate::error::ExcelError;
use crate::value::{ErrorKind, Value};
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
    Cell { sheet: usize, addr: CellAddr },
    Range { sheet: usize, start: CellAddr, end: CellAddr },
}

#[derive(Debug, Clone, PartialEq, Eq)]
pub enum TraceKind {
    Number,
    Text,
    Bool,
    Blank,
    Error,
    CellRef,
    RangeRef,
    Group,
    Unary { op: UnaryOp },
    Binary { op: crate::eval::BinaryOp },
    Compare { op: CompareOp },
    FunctionCall { name: String },
    ImplicitIntersection,
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
    Error(ErrorKind),
    CellRef(crate::eval::CellRef<S>),
    RangeRef(crate::eval::RangeRef<S>),
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
            SpannedExprKind::Error(e) => SpannedExprKind::Error(*e),
            SpannedExprKind::CellRef(r) => SpannedExprKind::CellRef(crate::eval::CellRef {
                sheet: f(&r.sheet),
                addr: r.addr,
            }),
            SpannedExprKind::RangeRef(r) => SpannedExprKind::RangeRef(crate::eval::RangeRef {
                sheet: f(&r.sheet),
                start: r.start,
                end: r.end,
            }),
            SpannedExprKind::Group(expr) => {
                SpannedExprKind::Group(Box::new(expr.map_sheets(f)))
            }
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
    let evaluator = TracedEvaluator { resolver, ctx };
    evaluator.eval_scalar(expr)
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
    Caret,
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
        Self { input, pos }
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
                ',' => {
                    self.pos += 1;
                    TokenKind::Comma
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
                '#' => self.lex_error()?,
                '.' | '0'..='9' => self.lex_number()?,
                _ if is_ident_start(ch) => self.lex_ident(),
                _ => {
                    return Err(FormulaParseError::UnexpectedToken(format!(
                        "unexpected character '{ch}'"
                    )))
                }
            };
            let span = Span::new(start, self.pos);
            tokens.push(Token { kind, span });
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
        while let Some(ch) = self.peek_char() {
            if ch.is_ascii_alphanumeric() || ch == '_' || ch == '.' || ch == '$' {
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
}

fn is_ident_start(ch: char) -> bool {
    ch.is_ascii_alphabetic() || ch == '_' || ch == '$'
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
        let mut left = self.parse_add_sub()?;
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
            let right = self.parse_add_sub()?;
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
        match &tok.kind {
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
                } else if matches!(self.peek_n(1).kind, TokenKind::Bang) {
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
                                kind: SpannedExprKind::Error(ErrorKind::Name),
                            }),
                        },
                    }
                }
            }
            TokenKind::SheetName(_name) => {
                if matches!(self.peek_n(1).kind, TokenKind::Bang) {
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
            other => Err(FormulaParseError::UnexpectedToken(format!("{other:?}"))),
        }
    }

    fn parse_function_call(&mut self) -> Result<SpannedExpr<String>, FormulaParseError> {
        let name_tok = self.next();
        let name = match name_tok.kind {
            TokenKind::Ident(s) => s.to_ascii_uppercase(),
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
        let sheet = match sheet_tok.kind {
            TokenKind::Ident(s) | TokenKind::SheetName(s) => SheetReference::Sheet(s),
            other => {
                return Err(FormulaParseError::Expected {
                    expected: "sheet name".to_string(),
                    got: format!("{other:?}"),
                })
            }
        };
        self.expect(TokenKind::Bang)?;

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
        let addr = parse_a1(&addr_str)?;
        self.parse_cell_or_range(sheet, sheet_tok.span.start, addr, addr_tok.span.end)
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
        self.tokens.get(self.pos).unwrap_or_else(|| self.tokens.last().unwrap())
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

#[derive(Debug, Clone, Copy)]
struct ResolvedRange {
    sheet_id: usize,
    start: CellAddr,
    end: CellAddr,
}

impl ResolvedRange {
    fn normalized(self) -> Self {
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
            sheet_id: self.sheet_id,
            start: CellAddr { row: r1, col: c1 },
            end: CellAddr { row: r2, col: c2 },
        }
    }

    fn is_single_cell(self) -> bool {
        self.start == self.end
    }

    fn iter_cells(self) -> impl Iterator<Item = CellAddr> {
        let norm = self.normalized();
        let rows = norm.start.row..=norm.end.row;
        let cols = norm.start.col..=norm.end.col;
        rows.flat_map(move |row| cols.clone().map(move |col| CellAddr { row, col }))
    }
}

#[derive(Debug, Clone)]
enum EvalValue {
    Scalar(Value),
    Reference(ResolvedRange),
}

struct TracedEvaluator<'a, R: crate::eval::ValueResolver> {
    resolver: &'a R,
    ctx: EvalContext,
}

impl<'a, R: crate::eval::ValueResolver> TracedEvaluator<'a, R> {
    fn eval_scalar(&self, expr: &SpannedExpr<usize>) -> (Value, TraceNode) {
        let (v, mut trace) = self.eval_value(expr);
        match v {
            EvalValue::Scalar(v) => (v, trace),
            EvalValue::Reference(range) => {
                let scalar = self.deref_reference_scalar(range);
                trace.value = scalar.clone();
                (scalar, trace)
            }
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
            SpannedExprKind::CellRef(r) => match self.resolve_sheet_id(&r.sheet) {
                Some(sheet_id) if self.resolver.sheet_exists(sheet_id) => {
                    let range = ResolvedRange {
                        sheet_id,
                        start: r.addr,
                        end: r.addr,
                    };
                    (
                        EvalValue::Reference(range),
                        TraceNode {
                            kind: TraceKind::CellRef,
                            span: expr.span,
                            value: Value::Blank,
                            reference: Some(TraceRef::Cell {
                                sheet: sheet_id,
                                addr: r.addr,
                            }),
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
            SpannedExprKind::RangeRef(r) => match self.resolve_sheet_id(&r.sheet) {
                Some(sheet_id) if self.resolver.sheet_exists(sheet_id) => {
                    let range = ResolvedRange {
                        sheet_id,
                        start: r.start,
                        end: r.end,
                    };
                    (
                        EvalValue::Reference(range),
                        TraceNode {
                            kind: TraceKind::RangeRef,
                            span: expr.span,
                            value: Value::Blank,
                            reference: Some(TraceRef::Range {
                                sheet: sheet_id,
                                start: r.start,
                                end: r.end,
                            }),
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
            SpannedExprKind::Unary { op, expr: inner } => {
                let (v, child) = self.eval_scalar(inner);
                let out = match v {
                    Value::Error(e) => Value::Error(e),
                    other => {
                        match coerce_to_number(&other) {
                            Ok(n) => match op {
                                UnaryOp::Plus => Value::Number(n),
                                UnaryOp::Minus => Value::Number(-n),
                            },
                            Err(e) => Value::Error(e),
                        }
                    }
                };
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
                let (l, ltrace) = self.eval_scalar(left);
                if let Value::Error(e) = l {
                    return (
                        EvalValue::Scalar(Value::Error(e)),
                        TraceNode {
                            kind: TraceKind::Binary { op: *op },
                            span: expr.span,
                            value: Value::Error(e),
                            reference: None,
                            children: vec![ltrace],
                        },
                    );
                }
                let (r, rtrace) = self.eval_scalar(right);
                if let Value::Error(e) = r {
                    return (
                        EvalValue::Scalar(Value::Error(e)),
                        TraceNode {
                            kind: TraceKind::Binary { op: *op },
                            span: expr.span,
                            value: Value::Error(e),
                            reference: None,
                            children: vec![ltrace, rtrace],
                        },
                    );
                }
                let ln = match coerce_to_number(&l) {
                    Ok(n) => n,
                    Err(e) => {
                        let out = Value::Error(e);
                        return (
                            EvalValue::Scalar(out.clone()),
                            TraceNode {
                                kind: TraceKind::Binary { op: *op },
                                span: expr.span,
                                value: out,
                                reference: None,
                                children: vec![ltrace, rtrace],
                            },
                        );
                    }
                };
                let rn = match coerce_to_number(&r) {
                    Ok(n) => n,
                    Err(e) => {
                        let out = Value::Error(e);
                        return (
                            EvalValue::Scalar(out.clone()),
                            TraceNode {
                                kind: TraceKind::Binary { op: *op },
                                span: expr.span,
                                value: out,
                                reference: None,
                                children: vec![ltrace, rtrace],
                            },
                        );
                    }
                };
                let out = match op {
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
                let (l, ltrace) = self.eval_scalar(left);
                if let Value::Error(e) = l {
                    return (
                        EvalValue::Scalar(Value::Error(e)),
                        TraceNode {
                            kind: TraceKind::Compare { op: *op },
                            span: expr.span,
                            value: Value::Error(e),
                            reference: None,
                            children: vec![ltrace],
                        },
                    );
                }
                let (r, rtrace) = self.eval_scalar(right);
                if let Value::Error(e) = r {
                    return (
                        EvalValue::Scalar(Value::Error(e)),
                        TraceNode {
                            kind: TraceKind::Compare { op: *op },
                            span: expr.span,
                            value: Value::Error(e),
                            reference: None,
                            children: vec![ltrace, rtrace],
                        },
                    );
                }
                let out = excel_compare(&l, &r, *op);
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
                    EvalValue::Reference(range) => self.apply_implicit_intersection(range),
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

    fn resolve_sheet_id(&self, sheet: &SheetReference<usize>) -> Option<usize> {
        match sheet {
            SheetReference::Current => Some(self.ctx.current_sheet),
            SheetReference::Sheet(id) => Some(*id),
            SheetReference::External(_) => None,
        }
    }

    fn deref_reference_scalar(&self, range: ResolvedRange) -> Value {
        if range.is_single_cell() {
            self.resolver.get_cell_value(range.sheet_id, range.start)
        } else {
            Value::Error(ErrorKind::Spill)
        }
    }

    fn apply_implicit_intersection(&self, range: ResolvedRange) -> Value {
        if range.is_single_cell() {
            return self.resolver.get_cell_value(range.sheet_id, range.start);
        }

        let range = range.normalized();
        let cur = self.ctx.current_cell;

        if range.start.col == range.end.col {
            if cur.row >= range.start.row && cur.row <= range.end.row {
                return self
                    .resolver
                    .get_cell_value(range.sheet_id, CellAddr { row: cur.row, col: range.start.col });
            }
            return Value::Error(ErrorKind::Value);
        }
        if range.start.row == range.end.row {
            if cur.col >= range.start.col && cur.col <= range.end.col {
                return self.resolver.get_cell_value(
                    range.sheet_id,
                    CellAddr { row: range.start.row, col: cur.col },
                );
            }
            return Value::Error(ErrorKind::Value);
        }

        if cur.row >= range.start.row
            && cur.row <= range.end.row
            && cur.col >= range.start.col
            && cur.col <= range.end.col
        {
            return self.resolver.get_cell_value(range.sheet_id, cur);
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
        let cond = match coerce_to_bool(&cond_val) {
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
                        if let Some(n) = parse_number_from_text(&s) {
                            acc += n;
                        }
                    }
                    Value::Array(arr) => {
                        for v in arr.iter() {
                            match v {
                                Value::Error(e) => return (Value::Error(*e), traces),
                                Value::Number(n) => acc += n,
                                Value::Bool(_) | Value::Text(_) | Value::Blank | Value::Array(_) | Value::Spill { .. } => {}
                            }
                        }
                    }
                    Value::Spill { .. } => return (Value::Error(ErrorKind::Value), traces),
                },
                EvalValue::Reference(range) => {
                    for addr in range.iter_cells() {
                        let v = self.resolver.get_cell_value(range.sheet_id, addr);
                        match v {
                            Value::Error(e) => return (Value::Error(e), traces),
                            Value::Number(n) => acc += n,
                            Value::Bool(_) | Value::Text(_) | Value::Blank | Value::Array(_) | Value::Spill { .. } => {}
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
            EvalValue::Reference(range) => range.normalized(),
            EvalValue::Scalar(Value::Error(e)) => return (Value::Error(e), traces),
            EvalValue::Scalar(_) => return (Value::Error(ErrorKind::Value), traces),
        };

        let (col_index_val, col_trace) = self.eval_scalar(&args[2]);
        traces.push(col_trace);
        if let Value::Error(e) = col_index_val {
            return (Value::Error(e), traces);
        }
        let col_index_num = match coerce_to_number(&col_index_val) {
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
            match coerce_to_bool(&v) {
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
                let candidate = self.resolver.get_cell_value(table_range.sheet_id, key);
                if matches!(candidate, Value::Error(_)) {
                    continue;
                }
                let is_match = excel_order(&candidate, &lookup_value)
                    .map(|o| o == Ordering::Equal)
                    .unwrap_or(false);
                if is_match {
                    let result_addr = CellAddr { row, col: target_col };
                    return (
                        self.resolver.get_cell_value(table_range.sheet_id, result_addr),
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
                let candidate = self.resolver.get_cell_value(table_range.sheet_id, key);
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
                let result_addr = CellAddr { row, col: target_col };
                (
                    self.resolver.get_cell_value(table_range.sheet_id, result_addr),
                    traces,
                )
            } else {
                (Value::Error(ErrorKind::NA), traces)
            }
        }
    }
}

fn parse_number_from_text(s: &str) -> Option<f64> {
    let trimmed = s.trim();
    if trimmed.is_empty() {
        return None;
    }
    trimmed.parse::<f64>().ok()
}

fn coerce_to_number(v: &Value) -> Result<f64, ErrorKind> {
    match v {
        Value::Number(n) => Ok(*n),
        Value::Bool(b) => Ok(if *b { 1.0 } else { 0.0 }),
        Value::Blank => Ok(0.0),
        Value::Text(s) => parse_number_from_text(s).ok_or(ErrorKind::Value),
        Value::Error(e) => Err(*e),
        Value::Array(_) | Value::Spill { .. } => Err(ErrorKind::Value),
    }
}

fn coerce_to_bool(v: &Value) -> Result<bool, ErrorKind> {
    match v {
        Value::Bool(b) => Ok(*b),
        Value::Number(n) => Ok(*n != 0.0),
        Value::Blank => Ok(false),
        Value::Text(s) => {
            let t = s.trim();
            if t.eq_ignore_ascii_case("TRUE") {
                return Ok(true);
            }
            if t.eq_ignore_ascii_case("FALSE") {
                return Ok(false);
            }
            if let Some(n) = parse_number_from_text(t) {
                return Ok(n != 0.0);
            }
            Err(ErrorKind::Value)
        }
        Value::Error(e) => Err(*e),
        Value::Array(_) | Value::Spill { .. } => Err(ErrorKind::Value),
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
    if matches!(left, Value::Array(_) | Value::Spill { .. })
        || matches!(right, Value::Array(_) | Value::Spill { .. })
    {
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
        | (Value::Spill { .. }, _)
        | (_, Value::Spill { .. }) => Ordering::Equal,
    })
}
