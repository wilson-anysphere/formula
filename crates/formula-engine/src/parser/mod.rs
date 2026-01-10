//! Formula lexer and parser.

use crate::{
    ArrayLiteral, Ast, BinaryExpr, BinaryOp, CellRef, Coord, Expr, FunctionCall, FunctionName,
    LocaleConfig, ParseError, ParseOptions, PostfixExpr, PostfixOp, Span, StructuredRef, UnaryExpr,
    UnaryOp,
};

#[derive(Debug, Clone, PartialEq, Eq)]
pub enum TokenKind {
    Number(String),
    String(String),
    Boolean(bool),
    Error(String),
    Cell(CellToken),
    Ident(String),
    QuotedIdent(String),
    Whitespace(String),
    Intersect(String),
    LParen,
    RParen,
    LBrace,
    RBrace,
    LBracket,
    RBracket,
    Bang,
    Colon,
    ArgSep,
    Union,
    ArrayRowSep,
    ArrayColSep,
    Plus,
    Minus,
    Star,
    Slash,
    Caret,
    Amp,
    Percent,
    Eq,
    Ne,
    Lt,
    Gt,
    Le,
    Ge,
    At,
    Eof,
}

#[derive(Debug, Clone, PartialEq, Eq)]
pub struct Token {
    pub kind: TokenKind,
    pub span: Span,
}

#[derive(Debug, Clone, PartialEq, Eq)]
pub struct CellToken {
    pub col: u32,
    pub row: u32,
    pub col_abs: bool,
    pub row_abs: bool,
}

pub fn parse_formula(formula: &str, opts: ParseOptions) -> Result<Ast, ParseError> {
    let (has_equals, expr_src) = if let Some(rest) = formula.strip_prefix('=') {
        (true, rest)
    } else {
        (false, formula)
    };

    let tokens = lex(expr_src, &opts.locale)?;
    let mut parser = Parser::new(expr_src, tokens, opts.locale.clone());
    let expr = parser.parse_expression(0)?;
    parser.expect(TokenKind::Eof)?;

    let mut ast = Ast::new(has_equals, expr);
    if let Some(origin) = opts.normalize_relative_to {
        ast = ast.normalize_relative(origin);
    }
    Ok(ast)
}

#[derive(Debug, Clone, PartialEq, Eq)]
pub struct FunctionContext {
    pub name: String,
    /// 0-indexed argument index.
    pub arg_index: usize,
}

#[derive(Debug, Clone, PartialEq, Eq, Default)]
pub struct ParseContext {
    pub function: Option<FunctionContext>,
}

#[derive(Debug, Clone, PartialEq, Eq)]
pub struct PartialParse {
    pub ast: Ast,
    pub error: Option<ParseError>,
    pub context: ParseContext,
}

/// Best-effort parsing used for editor/autocomplete scenarios.
///
/// Unlike [`parse_formula`], this API never returns an error. Instead, it returns:
/// - `error`: the first parse error encountered (if any)
/// - `context`: a coarse context (e.g. current function call + arg index)
/// - `ast`: a partial AST with missing nodes filled as [`Expr::Missing`]
pub fn parse_formula_partial(formula: &str, opts: ParseOptions) -> PartialParse {
    let (has_equals, expr_src) = if let Some(rest) = formula.strip_prefix('=') {
        (true, rest)
    } else {
        (false, formula)
    };

    let tokens = match lex(expr_src, &opts.locale) {
        Ok(t) => t,
        Err(e) => {
            return PartialParse {
                ast: Ast::new(has_equals, Expr::Missing),
                error: Some(e),
                context: ParseContext::default(),
            };
        }
    };

    let mut parser = Parser::new(expr_src, tokens, opts.locale.clone());
    let expr = parser.parse_expression_best_effort(0);

    let mut ast = Ast::new(has_equals, expr);
    if let Some(origin) = opts.normalize_relative_to {
        ast = ast.normalize_relative(origin);
    }

    let context = parser.context();
    let error = parser.first_error;

    PartialParse {
        ast,
        error,
        context,
    }
}

pub fn lex(formula: &str, locale: &LocaleConfig) -> Result<Vec<Token>, ParseError> {
    Lexer::new(formula, locale.clone()).lex()
}

#[derive(Debug, Clone)]
enum ParenContext {
    FunctionCall,
    Group,
}

struct Lexer<'a> {
    src: &'a str,
    chars: std::str::Chars<'a>,
    idx: usize,
    locale: LocaleConfig,
    tokens: Vec<Token>,
    paren_stack: Vec<ParenContext>,
    brace_depth: usize,
    prev_sig: Option<TokenKind>,
}

impl<'a> Lexer<'a> {
    fn new(src: &'a str, locale: LocaleConfig) -> Self {
        Self {
            src,
            chars: src.chars(),
            idx: 0,
            locale,
            tokens: Vec::new(),
            paren_stack: Vec::new(),
            brace_depth: 0,
            prev_sig: None,
        }
    }

    fn lex(mut self) -> Result<Vec<Token>, ParseError> {
        while let Some(ch) = self.peek_char() {
            let start = self.idx;
            match ch {
                ' ' | '\t' | '\r' | '\n' => {
                    let raw = self.take_while(|c| matches!(c, ' ' | '\t' | '\r' | '\n'));
                    self.push(TokenKind::Whitespace(raw), start, self.idx);
                }
                '"' => {
                    self.bump();
                    let mut value = String::new();
                    loop {
                        match self.peek_char() {
                            Some('"') => {
                                self.bump();
                                if self.peek_char() == Some('"') {
                                    self.bump();
                                    value.push('"');
                                    continue;
                                }
                                break;
                            }
                            Some(c) => {
                                self.bump();
                                value.push(c);
                            }
                            None => {
                                return Err(ParseError::new(
                                    "Unterminated string literal",
                                    Span::new(start, self.idx),
                                ));
                            }
                        }
                    }
                    self.push(TokenKind::String(value), start, self.idx);
                }
                '\'' => {
                    // Quoted identifier, typically for sheet names.
                    self.bump();
                    let mut value = String::new();
                    loop {
                        match self.peek_char() {
                            Some('\'') => {
                                self.bump();
                                if self.peek_char() == Some('\'') {
                                    self.bump();
                                    value.push('\'');
                                    continue;
                                }
                                break;
                            }
                            Some(c) => {
                                self.bump();
                                value.push(c);
                            }
                            None => {
                                return Err(ParseError::new(
                                    "Unterminated quoted identifier",
                                    Span::new(start, self.idx),
                                ));
                            }
                        }
                    }
                    self.push(TokenKind::QuotedIdent(value), start, self.idx);
                }
                '#' => {
                    self.bump();
                    let rest = self.take_while(|c| !is_delim(c));
                    let mut raw = String::from("#");
                    raw.push_str(&rest);
                    self.push(TokenKind::Error(raw), start, self.idx);
                }
                '(' => {
                    self.bump();
                    let is_func = matches!(self.prev_sig, Some(TokenKind::Ident(_)));
                    self.paren_stack.push(if is_func {
                        ParenContext::FunctionCall
                    } else {
                        ParenContext::Group
                    });
                    self.push(TokenKind::LParen, start, self.idx);
                }
                ')' => {
                    self.bump();
                    self.paren_stack.pop();
                    self.push(TokenKind::RParen, start, self.idx);
                }
                '{' => {
                    self.bump();
                    self.brace_depth += 1;
                    self.push(TokenKind::LBrace, start, self.idx);
                }
                '}' => {
                    self.bump();
                    self.brace_depth = self.brace_depth.saturating_sub(1);
                    self.push(TokenKind::RBrace, start, self.idx);
                }
                '[' => {
                    self.bump();
                    self.push(TokenKind::LBracket, start, self.idx);
                }
                ']' => {
                    self.bump();
                    self.push(TokenKind::RBracket, start, self.idx);
                }
                '!' => {
                    self.bump();
                    self.push(TokenKind::Bang, start, self.idx);
                }
                ':' => {
                    self.bump();
                    self.push(TokenKind::Colon, start, self.idx);
                }
                c if c == self.locale.arg_separator => {
                    self.bump();
                    if self.brace_depth > 0 {
                        // In array literals, commas/semicolons map to array separators.
                        if c == self.locale.array_row_separator {
                            self.push(TokenKind::ArrayRowSep, start, self.idx);
                        } else if c == self.locale.array_col_separator {
                            self.push(TokenKind::ArrayColSep, start, self.idx);
                        } else {
                            self.push(TokenKind::ArrayColSep, start, self.idx);
                        }
                    } else if matches!(self.paren_stack.last(), Some(ParenContext::FunctionCall)) {
                        self.push(TokenKind::ArgSep, start, self.idx);
                    } else {
                        self.push(TokenKind::Union, start, self.idx);
                    }
                }
                c if self.brace_depth > 0
                    && (c == self.locale.array_row_separator
                        || c == self.locale.array_col_separator) =>
                {
                    self.bump();
                    if c == self.locale.array_row_separator {
                        self.push(TokenKind::ArrayRowSep, start, self.idx);
                    } else {
                        self.push(TokenKind::ArrayColSep, start, self.idx);
                    }
                }
                '+' => {
                    self.bump();
                    self.push(TokenKind::Plus, start, self.idx);
                }
                '-' => {
                    self.bump();
                    self.push(TokenKind::Minus, start, self.idx);
                }
                '*' => {
                    self.bump();
                    self.push(TokenKind::Star, start, self.idx);
                }
                '/' => {
                    self.bump();
                    self.push(TokenKind::Slash, start, self.idx);
                }
                '^' => {
                    self.bump();
                    self.push(TokenKind::Caret, start, self.idx);
                }
                '&' => {
                    self.bump();
                    self.push(TokenKind::Amp, start, self.idx);
                }
                '%' => {
                    self.bump();
                    self.push(TokenKind::Percent, start, self.idx);
                }
                '@' => {
                    self.bump();
                    self.push(TokenKind::At, start, self.idx);
                }
                '=' => {
                    self.bump();
                    self.push(TokenKind::Eq, start, self.idx);
                }
                '<' => {
                    self.bump();
                    match self.peek_char() {
                        Some('=') => {
                            self.bump();
                            self.push(TokenKind::Le, start, self.idx);
                        }
                        Some('>') => {
                            self.bump();
                            self.push(TokenKind::Ne, start, self.idx);
                        }
                        _ => self.push(TokenKind::Lt, start, self.idx),
                    }
                }
                '>' => {
                    self.bump();
                    if self.peek_char() == Some('=') {
                        self.bump();
                        self.push(TokenKind::Ge, start, self.idx);
                    } else {
                        self.push(TokenKind::Gt, start, self.idx);
                    }
                }
                c if is_digit(c)
                    || (c == self.locale.decimal_separator && self.peek_next_is_digit()) =>
                {
                    let raw = self.lex_number();
                    self.push(TokenKind::Number(raw), start, self.idx);
                }
                '$' | '_' | '\\' | 'A'..='Z' | 'a'..='z' => {
                    if let Some(cell) = self.try_lex_cell_ref() {
                        self.push(TokenKind::Cell(cell), start, self.idx);
                    } else {
                        let ident = self.lex_ident();
                        let upper = ident.to_ascii_uppercase();
                        if upper == "TRUE" {
                            self.push(TokenKind::Boolean(true), start, self.idx);
                        } else if upper == "FALSE" {
                            self.push(TokenKind::Boolean(false), start, self.idx);
                        } else {
                            self.push(TokenKind::Ident(ident), start, self.idx);
                        }
                    }
                }
                _ => {
                    return Err(ParseError::new(
                        format!("Unexpected character `{ch}`"),
                        Span::new(start, self.idx + ch.len_utf8()),
                    ));
                }
            }
        }

        self.push(TokenKind::Eof, self.idx, self.idx);
        self.post_process_intersections();
        Ok(self.tokens)
    }

    fn post_process_intersections(&mut self) {
        let mut i = 0;
        while i < self.tokens.len() {
            if let TokenKind::Whitespace(raw) = &self.tokens[i].kind {
                let prev = prev_significant(&self.tokens, i);
                let next = next_significant(&self.tokens, i);
                if let (Some(p), Some(n)) = (prev, next) {
                    if is_intersect_operand(&self.tokens[p].kind)
                        && is_intersect_operand(&self.tokens[n].kind)
                        && raw.chars().any(|c| c == ' ' || c == '\t')
                    {
                        self.tokens[i].kind = TokenKind::Intersect(raw.clone());
                    }
                }
            }
            i += 1;
        }
    }

    fn push(&mut self, kind: TokenKind, start: usize, end: usize) {
        let sig = !matches!(kind, TokenKind::Whitespace(_));
        if sig {
            self.prev_sig = Some(kind.clone());
        }
        self.tokens.push(Token {
            kind,
            span: Span::new(start, end),
        });
    }

    fn bump(&mut self) -> Option<char> {
        let ch = self.chars.next()?;
        self.idx += ch.len_utf8();
        Some(ch)
    }

    fn peek_char(&self) -> Option<char> {
        self.src[self.idx..].chars().next()
    }

    fn peek_next_is_digit(&self) -> bool {
        let mut iter = self.src[self.idx..].chars();
        iter.next();
        matches!(iter.next(), Some(c) if is_digit(c))
    }

    fn take_while<F>(&mut self, mut pred: F) -> String
    where
        F: FnMut(char) -> bool,
    {
        let mut out = String::new();
        while let Some(ch) = self.peek_char() {
            if !pred(ch) {
                break;
            }
            self.bump();
            out.push(ch);
        }
        out
    }

    fn lex_number(&mut self) -> String {
        let mut out = String::new();
        // integer / leading decimal
        while let Some(ch) = self.peek_char() {
            if is_digit(ch) {
                self.bump();
                out.push(ch);
            } else {
                break;
            }
        }
        if self.peek_char() == Some(self.locale.decimal_separator) {
            self.bump();
            out.push(self.locale.decimal_separator);
            while let Some(ch) = self.peek_char() {
                if is_digit(ch) {
                    self.bump();
                    out.push(ch);
                } else {
                    break;
                }
            }
        }
        if matches!(self.peek_char(), Some('E' | 'e')) {
            let save_idx = self.idx;
            let save_chars = self.chars.clone();
            self.bump();
            let mut exp = String::from("E");
            if matches!(self.peek_char(), Some('+' | '-')) {
                let sign = self.bump().unwrap();
                exp.push(sign);
            }
            let mut digits = String::new();
            while let Some(ch) = self.peek_char() {
                if is_digit(ch) {
                    self.bump();
                    digits.push(ch);
                } else {
                    break;
                }
            }
            if digits.is_empty() {
                // roll back: the 'E' was part of an identifier maybe.
                self.idx = save_idx;
                self.chars = save_chars;
            } else {
                exp.push_str(&digits);
                out.push_str(&exp);
            }
        }
        out
    }

    fn lex_ident(&mut self) -> String {
        self.take_while(|c| matches!(c, '_' | '\\' | '.' | 'A'..='Z' | 'a'..='z' | '0'..='9'))
    }

    fn try_lex_cell_ref(&mut self) -> Option<CellToken> {
        let save_idx = self.idx;
        let save_chars = self.chars.clone();

        let mut col_abs = false;
        if self.peek_char() == Some('$') {
            col_abs = true;
            self.bump();
        }
        let mut col_letters = String::new();
        while let Some(ch) = self.peek_char() {
            if matches!(ch, 'A'..='Z' | 'a'..='z') {
                self.bump();
                col_letters.push(ch);
            } else {
                break;
            }
        }
        if col_letters.is_empty() {
            self.idx = save_idx;
            self.chars = save_chars;
            return None;
        }
        let mut row_abs = false;
        if self.peek_char() == Some('$') {
            row_abs = true;
            self.bump();
        }
        let mut row_digits = String::new();
        while let Some(ch) = self.peek_char() {
            if is_digit(ch) {
                self.bump();
                row_digits.push(ch);
            } else {
                break;
            }
        }
        if row_digits.is_empty() {
            self.idx = save_idx;
            self.chars = save_chars;
            return None;
        }

        let Some(col) = col_from_a1(&col_letters) else {
            self.idx = save_idx;
            self.chars = save_chars;
            return None;
        };
        let Some(row) = row_digits.parse::<u32>().ok() else {
            self.idx = save_idx;
            self.chars = save_chars;
            return None;
        };
        if row == 0 {
            self.idx = save_idx;
            self.chars = save_chars;
            return None;
        }
        Some(CellToken {
            col,
            row: row - 1,
            col_abs,
            row_abs,
        })
    }
}

fn is_digit(c: char) -> bool {
    matches!(c, '0'..='9')
}

fn is_delim(c: char) -> bool {
    matches!(
        c,
        ' ' | '\t'
            | '\r'
            | '\n'
            | '('
            | ')'
            | '{'
            | '}'
            | '['
            | ']'
            | '!'
            | ':'
            | '+'
            | '-'
            | '*'
            | '/'
            | '^'
            | '&'
            | '%'
            | '='
            | '<'
            | '>'
            | ','
            | ';'
    )
}

fn prev_significant(tokens: &[Token], idx: usize) -> Option<usize> {
    let mut j = idx;
    while j > 0 {
        j -= 1;
        if !matches!(tokens[j].kind, TokenKind::Whitespace(_)) {
            return Some(j);
        }
    }
    None
}

fn next_significant(tokens: &[Token], idx: usize) -> Option<usize> {
    let mut j = idx + 1;
    while j < tokens.len() {
        if !matches!(tokens[j].kind, TokenKind::Whitespace(_)) {
            return Some(j);
        }
        j += 1;
    }
    None
}

fn is_intersect_operand(kind: &TokenKind) -> bool {
    matches!(
        kind,
        TokenKind::Cell(_)
            | TokenKind::Ident(_)
            | TokenKind::QuotedIdent(_)
            | TokenKind::RParen
            | TokenKind::RBracket
    )
}

fn col_from_a1(letters: &str) -> Option<u32> {
    let mut col: u32 = 0;
    for (i, ch) in letters.chars().enumerate() {
        let v = (ch.to_ascii_uppercase() as u8).wrapping_sub(b'A') as u32;
        if v >= 26 {
            return None;
        }
        col = col * 26 + v + 1;
        if i >= 3 {
            return None;
        }
    }
    Some(col - 1)
}

struct Parser<'a> {
    src: &'a str,
    tokens: Vec<Token>,
    pos: usize,
    locale: LocaleConfig,
    func_stack: Vec<(String, usize)>,
    first_error: Option<ParseError>,
}

impl<'a> Parser<'a> {
    fn new(src: &'a str, tokens: Vec<Token>, locale: LocaleConfig) -> Self {
        Self {
            src,
            tokens,
            pos: 0,
            locale,
            func_stack: Vec::new(),
            first_error: None,
        }
    }

    fn parse_expression(&mut self, min_bp: u8) -> Result<Expr, ParseError> {
        self.skip_trivia();
        let mut lhs = self.parse_prefix()?;

        loop {
            self.skip_trivia();
            // Postfix percent.
            let percent_bp = 60;
            if matches!(self.peek_kind(), TokenKind::Percent) && percent_bp >= min_bp {
                self.next();
                lhs = Expr::Postfix(PostfixExpr {
                    op: PostfixOp::Percent,
                    expr: Box::new(lhs),
                });
                continue;
            }

            let op = match self.peek_kind() {
                TokenKind::Colon => Some(BinaryOp::Range),
                TokenKind::Intersect(_) => Some(BinaryOp::Intersect),
                TokenKind::Union => Some(BinaryOp::Union),
                TokenKind::Caret => Some(BinaryOp::Pow),
                TokenKind::Star => Some(BinaryOp::Mul),
                TokenKind::Slash => Some(BinaryOp::Div),
                TokenKind::Plus => Some(BinaryOp::Add),
                TokenKind::Minus => Some(BinaryOp::Sub),
                TokenKind::Amp => Some(BinaryOp::Concat),
                TokenKind::Eq => Some(BinaryOp::Eq),
                TokenKind::Ne => Some(BinaryOp::Ne),
                TokenKind::Lt => Some(BinaryOp::Lt),
                TokenKind::Gt => Some(BinaryOp::Gt),
                TokenKind::Le => Some(BinaryOp::Le),
                TokenKind::Ge => Some(BinaryOp::Ge),
                _ => None,
            };

            let Some(op) = op else { break };
            let (l_bp, r_bp) = infix_binding_power(op);
            if l_bp < min_bp {
                break;
            }
            self.next(); // consume operator
            let rhs = self.parse_expression(r_bp)?;
            lhs = Expr::Binary(BinaryExpr {
                op,
                left: Box::new(lhs),
                right: Box::new(rhs),
            });
        }

        Ok(lhs)
    }

    fn context(&self) -> ParseContext {
        let function = self
            .func_stack
            .last()
            .map(|(name, arg_index)| FunctionContext {
                name: name.clone(),
                arg_index: *arg_index,
            });
        ParseContext { function }
    }

    fn record_error(&mut self, err: ParseError) {
        if self.first_error.is_none() {
            self.first_error = Some(err);
        }
    }

    fn parse_expression_best_effort(&mut self, min_bp: u8) -> Expr {
        self.skip_trivia();
        let mut lhs = self.parse_prefix_best_effort();

        loop {
            self.skip_trivia();
            let percent_bp = 60;
            if matches!(self.peek_kind(), TokenKind::Percent) && percent_bp >= min_bp {
                self.next();
                lhs = Expr::Postfix(PostfixExpr {
                    op: PostfixOp::Percent,
                    expr: Box::new(lhs),
                });
                continue;
            }

            let op = match self.peek_kind() {
                TokenKind::Colon => Some(BinaryOp::Range),
                TokenKind::Intersect(_) => Some(BinaryOp::Intersect),
                TokenKind::Union => Some(BinaryOp::Union),
                TokenKind::Caret => Some(BinaryOp::Pow),
                TokenKind::Star => Some(BinaryOp::Mul),
                TokenKind::Slash => Some(BinaryOp::Div),
                TokenKind::Plus => Some(BinaryOp::Add),
                TokenKind::Minus => Some(BinaryOp::Sub),
                TokenKind::Amp => Some(BinaryOp::Concat),
                TokenKind::Eq => Some(BinaryOp::Eq),
                TokenKind::Ne => Some(BinaryOp::Ne),
                TokenKind::Lt => Some(BinaryOp::Lt),
                TokenKind::Gt => Some(BinaryOp::Gt),
                TokenKind::Le => Some(BinaryOp::Le),
                TokenKind::Ge => Some(BinaryOp::Ge),
                _ => None,
            };

            let Some(op) = op else { break };
            let (l_bp, r_bp) = infix_binding_power(op);
            if l_bp < min_bp {
                break;
            }
            self.next(); // consume operator
            let rhs = self.parse_expression_best_effort(r_bp);
            lhs = Expr::Binary(BinaryExpr {
                op,
                left: Box::new(lhs),
                right: Box::new(rhs),
            });
        }

        lhs
    }

    fn parse_prefix_best_effort(&mut self) -> Expr {
        self.skip_trivia();
        match self.peek_kind() {
            TokenKind::Plus => {
                self.next();
                let expr = self.parse_expression_best_effort(70);
                Expr::Unary(UnaryExpr {
                    op: UnaryOp::Plus,
                    expr: Box::new(expr),
                })
            }
            TokenKind::Minus => {
                self.next();
                let expr = self.parse_expression_best_effort(70);
                Expr::Unary(UnaryExpr {
                    op: UnaryOp::Minus,
                    expr: Box::new(expr),
                })
            }
            TokenKind::At => {
                self.next();
                let expr = self.parse_expression_best_effort(70);
                Expr::Unary(UnaryExpr {
                    op: UnaryOp::ImplicitIntersection,
                    expr: Box::new(expr),
                })
            }
            _ => self.parse_primary_best_effort(),
        }
    }

    fn parse_primary_best_effort(&mut self) -> Expr {
        self.skip_trivia();
        match self.peek_kind() {
            TokenKind::Number(raw) => {
                let raw = raw.clone();
                self.next();
                Expr::Number(raw)
            }
            TokenKind::String(value) => {
                let value = value.clone();
                self.next();
                Expr::String(value)
            }
            TokenKind::Boolean(v) => {
                let v = *v;
                self.next();
                Expr::Boolean(v)
            }
            TokenKind::Error(e) => {
                let e = e.clone();
                self.next();
                Expr::Error(e)
            }
            TokenKind::LParen => {
                self.next();
                let expr = self.parse_expression_best_effort(0);
                if let Err(e) = self.expect(TokenKind::RParen) {
                    self.record_error(e);
                }
                expr
            }
            TokenKind::LBrace => self.parse_array_literal_best_effort(),
            TokenKind::LBracket => match self.parse_bracket_start() {
                Ok(expr) => expr,
                Err(e) => {
                    self.record_error(e);
                    // Consume the '[' to avoid infinite loops.
                    if matches!(self.peek_kind(), TokenKind::LBracket) {
                        self.next();
                    }
                    Expr::Missing
                }
            },
            TokenKind::Cell(_) | TokenKind::Ident(_) | TokenKind::QuotedIdent(_) => {
                self.parse_reference_or_name_or_func_best_effort()
            }
            TokenKind::ArgSep | TokenKind::RParen | TokenKind::Eof => Expr::Missing,
            _ => {
                self.record_error(ParseError::new("Unexpected token", self.current_span()));
                // Consume one token and continue.
                if !matches!(self.peek_kind(), TokenKind::Eof) {
                    self.next();
                }
                Expr::Missing
            }
        }
    }

    fn parse_reference_or_name_or_func_best_effort(&mut self) -> Expr {
        // This is similar to `parse_reference_or_name_or_func`, but uses best-effort
        // function call parsing so editor states like `=SUM(A1,` still yield a useful AST.

        // Optional sheet prefix: Sheet1!A1 / 'My Sheet'!A1
        let save_pos = self.pos;
        let sheet_prefix = match self.peek_kind() {
            TokenKind::Ident(_) | TokenKind::QuotedIdent(_) => {
                let name = match self.take_name_token() {
                    Ok(s) => s,
                    Err(e) => {
                        self.record_error(e);
                        return Expr::Missing;
                    }
                };
                self.skip_trivia();
                if matches!(self.peek_kind(), TokenKind::Bang) {
                    self.next();
                    Some(name)
                } else {
                    self.pos = save_pos;
                    None
                }
            }
            _ => None,
        };

        if let Some(sheet) = sheet_prefix {
            return match self.parse_ref_after_prefix(None, Some(sheet)) {
                Ok(e) => e,
                Err(err) => {
                    self.record_error(err);
                    Expr::Missing
                }
            };
        }

        match self.peek_kind() {
            TokenKind::Ident(name) => {
                let name = name.clone();
                self.next();
                self.skip_trivia();
                if matches!(self.peek_kind(), TokenKind::LParen) {
                    self.parse_function_call_best_effort(name)
                } else if matches!(self.peek_kind(), TokenKind::LBracket) {
                    match self.parse_structured_ref(Some(name)) {
                        Ok(expr) => expr,
                        Err(err) => {
                            self.record_error(err);
                            Expr::Missing
                        }
                    }
                } else {
                    Expr::Name(name)
                }
            }
            TokenKind::Cell(cell) => {
                let cell = cell.clone();
                self.next();
                Expr::CellRef(CellRef {
                    workbook: None,
                    sheet: None,
                    col: Coord::A1 {
                        index: cell.col,
                        abs: cell.col_abs,
                    },
                    row: Coord::A1 {
                        index: cell.row,
                        abs: cell.row_abs,
                    },
                })
            }
            TokenKind::QuotedIdent(name) => {
                let name = name.clone();
                self.next();
                Expr::Name(name)
            }
            _ => {
                self.record_error(ParseError::new(
                    "Expected reference or name",
                    self.current_span(),
                ));
                Expr::Missing
            }
        }
    }

    fn parse_function_call_best_effort(&mut self, name: String) -> Expr {
        if let Err(e) = self.expect(TokenKind::LParen) {
            self.record_error(e);
            return Expr::Missing;
        }

        self.func_stack.push((name.clone(), 0));
        let mut args = Vec::new();

        loop {
            self.skip_trivia();
            match self.peek_kind() {
                TokenKind::RParen => {
                    self.next();
                    self.func_stack.pop();
                    break;
                }
                TokenKind::Eof => {
                    self.record_error(ParseError::new(
                        "Unterminated function call",
                        self.current_span(),
                    ));
                    // Don't pop the stack: context matters for autocomplete.
                    break;
                }
                _ => {}
            }

            // Parse an argument (or record it as missing).
            if matches!(self.peek_kind(), TokenKind::ArgSep) {
                args.push(Expr::Missing);
            } else {
                let arg = self.parse_expression_best_effort(0);
                args.push(arg);
            }

            self.skip_trivia();
            match self.peek_kind() {
                TokenKind::ArgSep => {
                    self.next();
                    if let Some((_n, idx)) = self.func_stack.last_mut() {
                        *idx += 1;
                    }
                    continue;
                }
                TokenKind::RParen => {
                    self.next();
                    self.func_stack.pop();
                    break;
                }
                TokenKind::Eof => {
                    self.record_error(ParseError::new(
                        "Unterminated function call",
                        self.current_span(),
                    ));
                    break;
                }
                _ => {
                    self.record_error(ParseError::new(
                        "Expected argument separator or `)`",
                        self.current_span(),
                    ));
                    // Attempt to resync by consuming one token.
                    if !matches!(self.peek_kind(), TokenKind::Eof) {
                        self.next();
                    }
                }
            }
        }

        Expr::FunctionCall(FunctionCall {
            name: FunctionName::new(name),
            args,
        })
    }

    fn parse_array_literal_best_effort(&mut self) -> Expr {
        if let Err(e) = self.expect(TokenKind::LBrace) {
            self.record_error(e);
            return Expr::Missing;
        }
        let mut rows: Vec<Vec<Expr>> = Vec::new();
        let mut current_row: Vec<Expr> = Vec::new();
        loop {
            self.skip_trivia();
            match self.peek_kind() {
                TokenKind::RBrace => {
                    self.next();
                    if !current_row.is_empty() || !rows.is_empty() {
                        rows.push(current_row);
                    }
                    break;
                }
                TokenKind::Eof => {
                    self.record_error(ParseError::new(
                        "Unterminated array literal",
                        self.current_span(),
                    ));
                    if !current_row.is_empty() || !rows.is_empty() {
                        rows.push(current_row);
                    }
                    break;
                }
                _ => {}
            }

            let el = self.parse_expression_best_effort(0);
            current_row.push(el);
            self.skip_trivia();
            match self.peek_kind() {
                TokenKind::ArrayColSep => {
                    self.next();
                }
                TokenKind::ArrayRowSep => {
                    self.next();
                    rows.push(current_row);
                    current_row = Vec::new();
                }
                TokenKind::RBrace => {}
                TokenKind::Eof => {}
                _ => {
                    self.record_error(ParseError::new(
                        "Expected array separator or `}`",
                        self.current_span(),
                    ));
                    // Try to continue by consuming one token.
                    if !matches!(self.peek_kind(), TokenKind::Eof) {
                        self.next();
                    }
                }
            }
        }
        Expr::Array(ArrayLiteral { rows })
    }

    fn parse_prefix(&mut self) -> Result<Expr, ParseError> {
        self.skip_trivia();
        match self.peek_kind() {
            TokenKind::Plus => {
                self.next();
                let expr = self.parse_expression(70)?;
                Ok(Expr::Unary(UnaryExpr {
                    op: UnaryOp::Plus,
                    expr: Box::new(expr),
                }))
            }
            TokenKind::Minus => {
                self.next();
                let expr = self.parse_expression(70)?;
                Ok(Expr::Unary(UnaryExpr {
                    op: UnaryOp::Minus,
                    expr: Box::new(expr),
                }))
            }
            TokenKind::At => {
                self.next();
                let expr = self.parse_expression(70)?;
                Ok(Expr::Unary(UnaryExpr {
                    op: UnaryOp::ImplicitIntersection,
                    expr: Box::new(expr),
                }))
            }
            _ => self.parse_primary(),
        }
    }

    fn parse_primary(&mut self) -> Result<Expr, ParseError> {
        self.skip_trivia();
        match self.peek_kind() {
            TokenKind::Number(raw) => {
                let raw = raw.clone();
                self.next();
                Ok(Expr::Number(raw))
            }
            TokenKind::String(value) => {
                let value = value.clone();
                self.next();
                Ok(Expr::String(value))
            }
            TokenKind::Boolean(v) => {
                let v = *v;
                self.next();
                Ok(Expr::Boolean(v))
            }
            TokenKind::Error(e) => {
                let e = e.clone();
                self.next();
                Ok(Expr::Error(e))
            }
            TokenKind::LParen => {
                self.next();
                let expr = self.parse_expression(0)?;
                self.expect(TokenKind::RParen)?;
                Ok(expr)
            }
            TokenKind::LBrace => self.parse_array_literal(),
            TokenKind::LBracket => self.parse_bracket_start(),
            TokenKind::Cell(_) | TokenKind::Ident(_) | TokenKind::QuotedIdent(_) => {
                self.parse_reference_or_name_or_func()
            }
            TokenKind::ArgSep => {
                // Missing argument, caller decides how to handle.
                Ok(Expr::Missing)
            }
            TokenKind::RParen | TokenKind::Eof => Ok(Expr::Missing),
            _ => {
                let span = self.current_span();
                Err(ParseError::new("Unexpected token", span))
            }
        }
    }

    fn parse_reference_or_name_or_func(&mut self) -> Result<Expr, ParseError> {
        // Handle optional external workbook prefix and/or sheet prefix.
        // We do this by peeking patterns: [Book]Sheet!..., Sheet!... etc.
        let save_pos = self.pos;

        let (workbook, sheet) = match self.peek_kind() {
            TokenKind::LBracket => unreachable!("handled elsewhere"),
            TokenKind::QuotedIdent(_) | TokenKind::Ident(_) => {
                // Could be sheet prefix (if followed by Bang), or function/name.
                let name = self.take_name_token()?;
                self.skip_trivia();
                if matches!(self.peek_kind(), TokenKind::Bang) {
                    self.next();
                    (None, Some(name))
                } else {
                    self.pos = save_pos;
                    (None, None)
                }
            }
            _ => (None, None),
        };

        // If we consumed a sheet prefix, parse the remainder as a reference/name.
        if sheet.is_some() {
            return self.parse_ref_after_prefix(workbook, sheet);
        }

        // No sheet prefix. Check function call.
        match self.peek_kind() {
            TokenKind::Ident(name) => {
                let name = name.clone();
                self.next();
                self.skip_trivia();
                if matches!(self.peek_kind(), TokenKind::LParen) {
                    self.parse_function_call(name)
                } else if matches!(self.peek_kind(), TokenKind::LBracket) {
                    self.parse_structured_ref(Some(name))
                } else {
                    Ok(Expr::Name(name))
                }
            }
            TokenKind::Cell(cell) => {
                let cell = cell.clone();
                self.next();
                Ok(Expr::CellRef(CellRef {
                    workbook: None,
                    sheet: None,
                    col: Coord::A1 {
                        index: cell.col,
                        abs: cell.col_abs,
                    },
                    row: Coord::A1 {
                        index: cell.row,
                        abs: cell.row_abs,
                    },
                }))
            }
            TokenKind::QuotedIdent(name) => {
                let name = name.clone();
                self.next();
                Ok(Expr::Name(name))
            }
            _ => Err(ParseError::new(
                "Expected reference or name",
                self.current_span(),
            )),
        }
    }

    fn parse_ref_after_prefix(
        &mut self,
        workbook: Option<String>,
        sheet: Option<String>,
    ) -> Result<Expr, ParseError> {
        self.skip_trivia();
        match self.peek_kind() {
            TokenKind::Cell(cell) => {
                let cell = cell.clone();
                self.next();
                Ok(Expr::CellRef(CellRef {
                    workbook,
                    sheet,
                    col: Coord::A1 {
                        index: cell.col,
                        abs: cell.col_abs,
                    },
                    row: Coord::A1 {
                        index: cell.row,
                        abs: cell.row_abs,
                    },
                }))
            }
            TokenKind::Ident(name) => {
                let name = name.clone();
                self.next();
                self.skip_trivia();
                if matches!(self.peek_kind(), TokenKind::LBracket) {
                    Ok(Expr::StructuredRef(StructuredRef {
                        table: Some(name),
                        spec: self.parse_bracket_spec()?,
                    }))
                } else {
                    Ok(Expr::Name(format!(
                        "{}!{}",
                        sheet.unwrap_or_default(),
                        name
                    )))
                }
            }
            TokenKind::LBracket => {
                let sr = self.parse_structured_ref(None)?;
                if let Expr::StructuredRef(mut s) = sr {
                    // Attach sheet/workbook prefix by serializing into a name-like string.
                    // Excel structured references can be prefixed with sheet/workbook; we model this
                    // as a name for now.
                    let mut rendered = String::new();
                    if let Some(book) = workbook {
                        rendered.push('[');
                        rendered.push_str(&book);
                        rendered.push(']');
                    }
                    if let Some(sh) = sheet {
                        rendered.push_str(&sh);
                        rendered.push('!');
                    }
                    // structured ref itself
                    if let Some(table) = s.table.take() {
                        rendered.push_str(&table);
                    }
                    rendered.push('[');
                    rendered.push_str(&s.spec);
                    rendered.push(']');
                    Ok(Expr::Name(rendered))
                } else {
                    Ok(sr)
                }
            }
            _ => Err(ParseError::new(
                "Expected reference after sheet prefix",
                self.current_span(),
            )),
        }
    }

    fn parse_function_call(&mut self, name: String) -> Result<Expr, ParseError> {
        self.expect(TokenKind::LParen)?;
        self.func_stack.push((name.clone(), 0));
        let mut args = Vec::new();
        self.skip_trivia();
        if matches!(self.peek_kind(), TokenKind::RParen) {
            self.next();
        } else {
            loop {
                self.skip_trivia();
                if matches!(self.peek_kind(), TokenKind::ArgSep) {
                    // Missing argument.
                    args.push(Expr::Missing);
                } else {
                    let arg = self.parse_expression(0)?;
                    args.push(arg);
                }
                self.skip_trivia();
                match self.peek_kind() {
                    TokenKind::ArgSep => {
                        self.next();
                        if let Some((_n, idx)) = self.func_stack.last_mut() {
                            *idx += 1;
                        }
                        continue;
                    }
                    TokenKind::RParen => {
                        self.next();
                        break;
                    }
                    _ => {
                        return Err(ParseError::new(
                            "Expected argument separator or `)`",
                            self.current_span(),
                        ));
                    }
                }
            }
        }
        self.func_stack.pop();
        Ok(Expr::FunctionCall(FunctionCall {
            name: FunctionName::new(name),
            args,
        }))
    }

    fn parse_array_literal(&mut self) -> Result<Expr, ParseError> {
        self.expect(TokenKind::LBrace)?;
        let mut rows: Vec<Vec<Expr>> = Vec::new();
        let mut current_row: Vec<Expr> = Vec::new();
        loop {
            self.skip_trivia();
            if matches!(self.peek_kind(), TokenKind::RBrace) {
                self.next();
                if !current_row.is_empty() || !rows.is_empty() {
                    rows.push(current_row);
                }
                break;
            }

            let el = self.parse_expression(0)?;
            current_row.push(el);
            self.skip_trivia();
            match self.peek_kind() {
                TokenKind::ArrayColSep => {
                    self.next();
                }
                TokenKind::ArrayRowSep => {
                    self.next();
                    rows.push(current_row);
                    current_row = Vec::new();
                }
                TokenKind::RBrace => {
                    // loop will close
                }
                _ => {
                    return Err(ParseError::new(
                        "Expected array separator or `}`",
                        self.current_span(),
                    ));
                }
            }
        }
        Ok(Expr::Array(ArrayLiteral { rows }))
    }

    fn parse_bracket_start(&mut self) -> Result<Expr, ParseError> {
        // Could be an external workbook prefix ([Book]Sheet!A1) or a structured ref like [@Col].
        // Look ahead for pattern: [ ... ] <sheet> !
        let save = self.pos;
        self.expect(TokenKind::LBracket)?;
        let book_start = self.pos;
        while !matches!(self.peek_kind(), TokenKind::RBracket | TokenKind::Eof) {
            self.next();
        }
        self.expect(TokenKind::RBracket)?;
        let book_span = Span::new(
            self.tokens[book_start].span.start,
            self.tokens[self.pos - 1].span.end,
        );
        let workbook = self.src[book_span.start..book_span.end]
            .trim_matches(&['[', ']'][..])
            .to_string();
        self.skip_trivia();
        let sheet = match self.peek_kind() {
            TokenKind::Ident(_) | TokenKind::QuotedIdent(_) => self.take_name_token()?,
            _ => {
                // Not an external ref; rewind and parse as structured.
                self.pos = save;
                return self.parse_structured_ref(None);
            }
        };
        self.skip_trivia();
        if !matches!(self.peek_kind(), TokenKind::Bang) {
            self.pos = save;
            return self.parse_structured_ref(None);
        }
        self.next(); // bang
        self.parse_ref_after_prefix(Some(workbook), Some(sheet))
    }

    fn parse_structured_ref(&mut self, table: Option<String>) -> Result<Expr, ParseError> {
        let spec = if matches!(self.peek_kind(), TokenKind::LBracket) {
            self.next();
            self.parse_bracket_spec()?
        } else {
            self.expect(TokenKind::LBracket)?;
            self.parse_bracket_spec()?
        };
        Ok(Expr::StructuredRef(StructuredRef { table, spec }))
    }

    fn parse_bracket_spec(&mut self) -> Result<String, ParseError> {
        let mut spec = String::new();
        loop {
            match self.peek_kind() {
                TokenKind::RBracket => {
                    self.next();
                    break;
                }
                TokenKind::At => {
                    self.next();
                    spec.push('@');
                }
                TokenKind::Ident(s) => {
                    let s = s.clone();
                    self.next();
                    spec.push_str(&s);
                }
                TokenKind::QuotedIdent(s) => {
                    let s = s.clone();
                    self.next();
                    spec.push('\'');
                    spec.push_str(&s);
                    spec.push('\'');
                }
                TokenKind::Whitespace(w) => {
                    let w = w.clone();
                    self.next();
                    spec.push_str(&w);
                }
                TokenKind::Colon => {
                    self.next();
                    spec.push(':');
                }
                TokenKind::Union => {
                    self.next();
                    spec.push(self.locale.arg_separator);
                }
                TokenKind::ArgSep => {
                    self.next();
                    spec.push(self.locale.arg_separator);
                }
                TokenKind::Number(n) => {
                    let n = n.clone();
                    self.next();
                    spec.push_str(&n);
                }
                TokenKind::Error(e) => {
                    let e = e.clone();
                    self.next();
                    spec.push_str(&e);
                }
                TokenKind::Eof => {
                    return Err(ParseError::new(
                        "Unterminated structured reference",
                        self.current_span(),
                    ));
                }
                _ => {
                    // Fallback: include raw source span of the token.
                    let span = self.current_span();
                    let raw = self.src[span.start..span.end].to_string();
                    self.next();
                    spec.push_str(&raw);
                }
            }
        }
        Ok(spec)
    }

    fn take_name_token(&mut self) -> Result<String, ParseError> {
        self.skip_trivia();
        match self.peek_kind() {
            TokenKind::Ident(s) => {
                let s = s.clone();
                self.next();
                Ok(s)
            }
            TokenKind::QuotedIdent(s) => {
                let s = s.clone();
                self.next();
                Ok(s)
            }
            _ => Err(ParseError::new("Expected name", self.current_span())),
        }
    }

    fn expect(&mut self, kind: TokenKind) -> Result<(), ParseError> {
        self.skip_trivia();
        if std::mem::discriminant(self.peek_kind()) == std::mem::discriminant(&kind) {
            self.next();
            Ok(())
        } else {
            Err(ParseError::new(
                format!("Expected {:?}", kind),
                self.current_span(),
            ))
        }
    }

    fn skip_trivia(&mut self) {
        while matches!(self.peek_kind(), TokenKind::Whitespace(_)) {
            self.pos += 1;
        }
    }

    fn peek_kind(&self) -> &TokenKind {
        &self.tokens[self.pos].kind
    }

    fn next(&mut self) -> &Token {
        let tok = &self.tokens[self.pos];
        self.pos += 1;
        tok
    }

    fn current_span(&self) -> Span {
        self.tokens
            .get(self.pos)
            .map(|t| t.span)
            .unwrap_or_else(|| Span::new(self.src.len(), self.src.len()))
    }
}

fn infix_binding_power(op: BinaryOp) -> (u8, u8) {
    match op {
        BinaryOp::Range => (82, 83),
        BinaryOp::Intersect => (81, 82),
        BinaryOp::Union => (80, 81),
        BinaryOp::Pow => (50, 50), // right associative
        BinaryOp::Mul | BinaryOp::Div => (40, 41),
        BinaryOp::Add | BinaryOp::Sub => (30, 31),
        BinaryOp::Concat => (20, 21),
        BinaryOp::Eq | BinaryOp::Ne | BinaryOp::Lt | BinaryOp::Gt | BinaryOp::Le | BinaryOp::Ge => {
            (10, 11)
        }
    }
}
