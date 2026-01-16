//! Formula lexer and parser.

use crate::{
    ArrayLiteral, Ast, BinaryExpr, BinaryOp, CallExpr, CellRef, ColRef, Coord, Expr,
    FieldAccessExpr, FunctionCall, FunctionName, LocaleConfig, NameRef, ParseError, ParseOptions,
    PostfixExpr, PostfixOp, ReferenceStyle, RowRef, SheetRef, Span, StructuredRef, UnaryExpr,
    UnaryOp,
};
use formula_model::{column_label_to_index_lenient, sheet_name_eq_case_insensitive};
use std::borrow::Cow;

/// Excel formula limits enforced by this parser.
///
/// These are primarily intended to:
/// - match Excel compatibility constraints
/// - prevent pathological formulas from consuming excessive CPU/memory or overflowing the Rust
///   stack during parsing/evaluation
///
/// Reference: `instructions/core-engine.md`.
const EXCEL_MAX_FORMULA_CHARS: usize = 8_192;
const EXCEL_MAX_TOKENIZED_BYTES: usize = 16_384;
const EXCEL_MAX_NESTED_CALLS: usize = 64;

#[derive(Debug, Clone, PartialEq, Eq)]
pub enum TokenKind {
    Number(String),
    String(String),
    Boolean(bool),
    Error(String),
    Cell(CellToken),
    R1C1Cell(R1C1CellToken),
    R1C1Row(R1C1RowToken),
    R1C1Col(R1C1ColToken),
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
    Dot,
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
    Hash,
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

#[derive(Debug, Clone, Copy, PartialEq, Eq)]
pub struct CellToken {
    pub col: u32,
    pub row: u32,
    pub col_abs: bool,
    pub row_abs: bool,
}

#[derive(Debug, Clone, Copy, PartialEq, Eq)]
pub struct R1C1CellToken {
    pub row: Coord,
    pub col: Coord,
}

#[derive(Debug, Clone, Copy, PartialEq, Eq)]
pub struct R1C1RowToken {
    pub row: Coord,
}

#[derive(Debug, Clone, Copy, PartialEq, Eq)]
pub struct R1C1ColToken {
    pub col: Coord,
}

pub fn parse_formula(formula: &str, opts: ParseOptions) -> Result<Ast, ParseError> {
    // Excel's formula display limit is 8,192 characters. We count Unicode scalar values (`char`)
    // rather than bytes to behave reasonably for non-ASCII formulas.
    let char_len = formula.chars().count();
    if char_len > EXCEL_MAX_FORMULA_CHARS {
        return Err(ParseError::new(
            format!(
                "Formula exceeds Excel's {EXCEL_MAX_FORMULA_CHARS}-character limit (got {char_len})"
            ),
            Span::new(0, formula.len()),
        ));
    }

    let (has_equals, expr_src, span_offset) = if let Some(rest) = formula.strip_prefix('=') {
        (true, rest, 1)
    } else {
        (false, formula, 0)
    };

    let tokens = lex(expr_src, &opts).map_err(|e| e.add_offset(span_offset))?;
    let mut parser = Parser::new(expr_src, tokens);
    let expr = parser
        .parse_expression(0)
        .map_err(|e| e.add_offset(span_offset))?;
    parser
        .expect(TokenKind::Eof)
        .map_err(|e| e.add_offset(span_offset))?;

    // Excel also enforces a 16,384-byte limit on the internal tokenized form of a formula.
    //
    // We do not implement Excel's BIFF ptg serializer here, but we *approximate* the tokenized
    // size using a deterministic, conservative per-AST-node byte estimate derived from common ptg
    // sizes (e.g. numbers are 3 or 9 bytes, cell refs ~5 bytes, operators ~1 byte).
    //
    // This provides a practical guard against formulas that are short in text form but expand to
    // a very large internal representation (e.g. thousands of numeric literals).
    let estimated_bytes = estimate_tokenized_bytes(&expr);
    if estimated_bytes > EXCEL_MAX_TOKENIZED_BYTES {
        return Err(
            ParseError::new(
                format!(
                    "Formula exceeds Excel's {EXCEL_MAX_TOKENIZED_BYTES}-byte tokenized limit (estimated {estimated_bytes} bytes)"
                ),
                Span::new(0, expr_src.len()),
            )
            .add_offset(span_offset),
        );
    }

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
    let (has_equals, expr_src, span_offset) = if let Some(rest) = formula.strip_prefix('=') {
        (true, rest, 1)
    } else {
        (false, formula, 0)
    };

    let char_len = formula.chars().count();
    if char_len > EXCEL_MAX_FORMULA_CHARS {
        return PartialParse {
            ast: Ast::new(has_equals, Expr::Missing),
            error: Some(ParseError::new(
                format!(
                    "Formula exceeds Excel's {EXCEL_MAX_FORMULA_CHARS}-character limit (got {char_len})"
                ),
                Span::new(0, formula.len()),
            )),
            context: ParseContext::default(),
        };
    }

    let tokens = match lex(expr_src, &opts) {
        Ok(t) => t,
        Err(e) => {
            return PartialParse {
                ast: Ast::new(has_equals, Expr::Missing),
                error: Some(e.add_offset(span_offset)),
                context: ParseContext::default(),
            };
        }
    };

    let mut parser = Parser::new(expr_src, tokens);
    let expr = parser.parse_expression_best_effort(0);

    let mut ast = Ast::new(has_equals, expr);
    if let Some(origin) = opts.normalize_relative_to {
        ast = ast.normalize_relative(origin);
    }

    let context = parser.take_context();
    let error = parser.first_error.map(|e| e.add_offset(span_offset));

    PartialParse {
        ast,
        error,
        context,
    }
}

pub fn lex(formula: &str, opts: &ParseOptions) -> Result<Vec<Token>, ParseError> {
    Lexer::new(formula, opts.locale, opts.reference_style).lex()
}

#[derive(Debug, Clone, PartialEq, Eq)]
pub struct PartialLex {
    pub tokens: Vec<Token>,
    pub error: Option<ParseError>,
}

/// Best-effort lexing used for editor/syntax-highlighting scenarios.
///
/// Unlike [`lex`], this API never returns an error. Instead, it returns:
/// - `tokens`: as many tokens as possible (always ending with [`TokenKind::Eof`])
/// - `error`: the first lex error encountered (if any)
pub fn lex_partial(formula: &str, opts: &ParseOptions) -> PartialLex {
    Lexer::new(formula, opts.locale, opts.reference_style).lex_partial()
}

#[derive(Debug, Clone, Copy, PartialEq, Eq)]
enum ParenContext {
    /// Parentheses opened as part of a function call, along with the brace depth at the `(`.
    ///
    /// This is used to disambiguate locale separators that overlap between function argument
    /// separators and array literal separators. For example, in `={SUM(1,2),3}` the comma inside
    /// `SUM(1,2)` should be lexed as a function argument separator, while the comma after the
    /// closing `)` should be lexed as an array column separator.
    FunctionCall {
        brace_depth: usize,
    },
    Group,
}

#[derive(Debug, Clone, Copy, PartialEq, Eq)]
enum LexMode {
    Strict,
    BestEffort,
}

#[derive(Debug, Clone, Copy, PartialEq, Eq)]
enum PrevSig {
    Number,
    String,
    Boolean,
    Error,
    Cell,
    R1C1Cell,
    R1C1Row,
    R1C1Col,
    Ident,
    QuotedIdent,
    RParen,
    RBrace,
    RBracket,
    Hash,
    Percent,
    Other,
}

impl PrevSig {
    fn from_kind(kind: &TokenKind) -> Self {
        match kind {
            TokenKind::Number(_) => Self::Number,
            TokenKind::String(_) => Self::String,
            TokenKind::Boolean(_) => Self::Boolean,
            TokenKind::Error(_) => Self::Error,
            TokenKind::Cell(_) => Self::Cell,
            TokenKind::R1C1Cell(_) => Self::R1C1Cell,
            TokenKind::R1C1Row(_) => Self::R1C1Row,
            TokenKind::R1C1Col(_) => Self::R1C1Col,
            TokenKind::Ident(_) => Self::Ident,
            TokenKind::QuotedIdent(_) => Self::QuotedIdent,
            TokenKind::RParen => Self::RParen,
            TokenKind::RBrace => Self::RBrace,
            TokenKind::RBracket => Self::RBracket,
            TokenKind::Hash => Self::Hash,
            TokenKind::Percent => Self::Percent,
            _ => Self::Other,
        }
    }
}

struct Lexer<'a> {
    src: &'a str,
    chars: std::str::Chars<'a>,
    idx: usize,
    locale: LocaleConfig,
    reference_style: ReferenceStyle,
    tokens: Vec<Token>,
    paren_stack: Vec<ParenContext>,
    brace_depth: usize,
    bracket_depth: usize,
    prev_sig: Option<PrevSig>,
}

fn find_workbook_prefix_end_if_valid(src: &str, start: usize) -> Option<usize> {
    formula_model::external_refs::find_external_workbook_prefix_end_if_followed_by_sheet_or_name_token(
        src, start,
    )
}

impl<'a> Lexer<'a> {
    fn new(src: &'a str, locale: LocaleConfig, reference_style: ReferenceStyle) -> Self {
        Self {
            src,
            chars: src.chars(),
            idx: 0,
            locale,
            reference_style,
            tokens: Vec::new(),
            paren_stack: Vec::new(),
            brace_depth: 0,
            bracket_depth: 0,
            prev_sig: None,
        }
    }

    fn lex(self) -> Result<Vec<Token>, ParseError> {
        let (tokens, _) = self.lex_with_mode(LexMode::Strict)?;
        Ok(tokens)
    }

    fn lex_partial(self) -> PartialLex {
        let (tokens, error) = self
            .lex_with_mode(LexMode::BestEffort)
            .expect("best-effort lexer should not return an error");
        PartialLex { tokens, error }
    }

    fn lex_with_mode(
        mut self,
        mode: LexMode,
    ) -> Result<(Vec<Token>, Option<ParseError>), ParseError> {
        let mut first_error: Option<ParseError> = None;

        let mut handle_error = |err: ParseError, stop_scanning: bool| -> Result<bool, ParseError> {
            match mode {
                LexMode::Strict => Err(err),
                LexMode::BestEffort => {
                    if first_error.is_none() {
                        first_error = Some(err);
                    }
                    Ok(stop_scanning)
                }
            }
        };

        while let Some(ch) = self.peek_char() {
            let start = self.idx;
            if self.bracket_depth > 0 && !matches!(ch, '[' | ']' | '"') {
                // Inside workbook/structured reference brackets, treat everything as raw text so
                // locale separators (e.g. `,` in `Table1[[#Headers],[Col]]`) don't get lexed as
                // unions/arg separators and non-locale delimiters don't fail lexing.
                let raw = self.take_while(|c| !matches!(c, '[' | ']'));
                self.push(TokenKind::Ident(raw), start, self.idx);
                continue;
            }
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
                                handle_error(
                                    ParseError::new(
                                    "Unterminated string literal",
                                    Span::new(start, self.idx),
                                    ),
                                    false,
                                )?;
                                break;
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
                                handle_error(
                                    ParseError::new(
                                    "Unterminated quoted identifier",
                                    Span::new(start, self.idx),
                                    ),
                                    false,
                                )?;
                                break;
                            }
                        }
                    }
                    self.push(TokenKind::QuotedIdent(value), start, self.idx);
                }
                '#' => {
                    // Excel's spill-range reference operator (`#`) is postfix (e.g. `A1#`),
                    // but error literals also start with `#` (e.g. `#REF!`).
                    //
                    // Treat `#` as a postfix operator only when it is *immediately* after an
                    // expression-like token (no intervening whitespace).
                    let is_immediate = self.tokens.last().is_some_and(|t| {
                        t.span.end == start && !matches!(t.kind, TokenKind::Whitespace(_))
                    });
                    let is_postfix_spill = is_immediate
                        && matches!(
                            self.prev_sig,
                            Some(
                                PrevSig::Cell
                                    | PrevSig::Ident
                                    | PrevSig::QuotedIdent
                                    | PrevSig::RParen
                                    | PrevSig::RBracket
                            )
                        );

                    if is_postfix_spill {
                        self.bump();
                        self.push(TokenKind::Hash, start, self.idx);
                        continue;
                    }

                    if let Some(len) = match_error_literal(&self.src[start..]) {
                        let end = start + len;
                        while self.idx < end {
                            self.bump();
                        }
                        let raw = self.src[start..end].to_string();
                        self.push(TokenKind::Error(raw), start, self.idx);
                    } else if self
                        .src
                        .get(self.idx + 1..)
                        .and_then(|s| s.chars().next())
                        .is_some_and(is_error_body_char)
                    {
                        self.bump(); // '#'
                        let mut rest = String::from("#");
                        self.take_while_into(is_error_body_char, &mut rest);
                        if matches!(self.peek_char(), Some('!' | '?')) {
                            if let Some(ch) = self.bump() {
                                rest.push(ch);
                            } else {
                                debug_assert!(false, "peek_char ensured char exists");
                            }
                        }
                        self.push(TokenKind::Error(rest), start, self.idx);
                    } else {
                        // Standalone `#` is the spill-range reference postfix operator (e.g. `A1#`).
                        self.bump();
                        self.push(TokenKind::Hash, start, self.idx);
                    }
                }
                '(' => {
                    self.bump();
                    let is_func = matches!(
                        self.prev_sig,
                        Some(
                            PrevSig::Number
                                | PrevSig::String
                                | PrevSig::Boolean
                                | PrevSig::Error
                                | PrevSig::Cell
                                | PrevSig::R1C1Cell
                                | PrevSig::R1C1Row
                                | PrevSig::R1C1Col
                                | PrevSig::Ident
                                | PrevSig::QuotedIdent
                                | PrevSig::RParen
                                | PrevSig::RBrace
                                | PrevSig::RBracket
                                | PrevSig::Hash
                                | PrevSig::Percent
                        )
                    );
                    self.paren_stack.push(if is_func {
                        ParenContext::FunctionCall {
                            brace_depth: self.brace_depth,
                        }
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
                    if self.bracket_depth == 0 {
                        // Workbook prefixes are *not* nesting, even if the workbook name contains
                        // `[` characters (e.g. `=[A1[Name.xlsx]Sheet1!A1`). Prefer a non-nesting
                        // scan when the bracketed segment is followed by a sheet name and `!`.
                        if let Some(end) = find_workbook_prefix_end_if_valid(self.src, start) {
                            self.bump();
                            self.push(TokenKind::LBracket, start, self.idx);

                            let inner_start = self.idx;
                            let inner_end = end.saturating_sub(1);
                            if inner_end > inner_start {
                                let raw = self.src[inner_start..inner_end].to_string();
                                self.rollback_to(inner_end);
                                self.push(TokenKind::Ident(raw), inner_start, inner_end);
                            }

                            let close_start = self.idx;
                            self.bump();
                            self.push(TokenKind::RBracket, close_start, self.idx);
                            continue;
                        }
                    }

                    self.bump();
                    self.bracket_depth += 1;
                    self.push(TokenKind::LBracket, start, self.idx);
                }
                ']' => {
                    // Excel escapes `]` inside structured references as `]]`. At the outermost
                    // bracket depth, treat a double `]]` as a literal `]` rather than the end of
                    // the bracketed segment.
                    if self.bracket_depth == 1 && self.src[self.idx..].starts_with("]]") {
                        self.bump();
                        self.push(TokenKind::RBracket, start, self.idx);
                        let start2 = self.idx;
                        self.bump();
                        self.push(TokenKind::RBracket, start2, self.idx);
                        continue;
                    }
                    self.bump();
                    self.bracket_depth = self.bracket_depth.saturating_sub(1);
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
                    let is_func_arg_sep = matches!(
                        self.paren_stack.last(),
                        Some(ParenContext::FunctionCall { brace_depth }) if *brace_depth == self.brace_depth
                    );
                    if self.brace_depth > 0 && !is_func_arg_sep {
                        // In array literals, commas/semicolons map to array separators.
                        if c == self.locale.array_row_separator {
                            self.push(TokenKind::ArrayRowSep, start, self.idx);
                        } else if c == self.locale.array_col_separator {
                            self.push(TokenKind::ArrayColSep, start, self.idx);
                        } else {
                            self.push(TokenKind::ArrayColSep, start, self.idx);
                        }
                    } else if is_func_arg_sep {
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
                    || ((c == self.locale.decimal_separator || c == '.')
                        && self.peek_next_is_digit()) =>
                {
                    let raw = self.lex_number();
                    self.push(TokenKind::Number(raw), start, self.idx);
                }
                '.' => {
                    self.bump();
                    self.push(TokenKind::Dot, start, self.idx);
                }
                c if is_ident_start_char(c) => {
                    if self.reference_style == ReferenceStyle::R1C1 {
                        if let Some(cell) = self.try_lex_r1c1_cell_ref() {
                            self.push(TokenKind::R1C1Cell(cell), start, self.idx);
                            continue;
                        }
                        if let Some(row) = self.try_lex_r1c1_row_ref() {
                            self.push(TokenKind::R1C1Row(row), start, self.idx);
                            continue;
                        }
                        if let Some(col) = self.try_lex_r1c1_col_ref() {
                            self.push(TokenKind::R1C1Col(col), start, self.idx);
                            continue;
                        }
                    }

                    if self.may_be_a1_cell_ref() {
                        if let Some(cell) = self.try_lex_cell_ref() {
                            self.push(TokenKind::Cell(cell), start, self.idx);
                            continue;
                        }
                    }

                    let ident = self.lex_ident();
                    let is_true = ident.eq_ignore_ascii_case("TRUE");
                    let is_false = ident.eq_ignore_ascii_case("FALSE");
                    if is_true || is_false {
                        // Excel supports `TRUE` / `FALSE` as both boolean literals *and* zero-arg
                        // functions (`TRUE()` / `FALSE()`). Lex `TRUE`/`FALSE` as booleans only
                        // when they are standalone literals; if the next non-whitespace
                        // character is `(`, treat them as identifiers so the parser produces a
                        // `FunctionCall`.
                        let next_non_ws = self.src[self.idx..]
                            .chars()
                            .find(|c| !matches!(c, ' ' | '\t' | '\r' | '\n'));
                        if next_non_ws == Some('(') {
                            self.push(TokenKind::Ident(ident), start, self.idx);
                        } else if is_true {
                            self.push(TokenKind::Boolean(true), start, self.idx);
                        } else {
                            self.push(TokenKind::Boolean(false), start, self.idx);
                        }
                    } else {
                        self.push(TokenKind::Ident(ident), start, self.idx);
                    }
                }
                _ => {
                    if handle_error(
                        ParseError::new(
                        format!("Unexpected character `{ch}`"),
                        Span::new(start, self.idx + ch.len_utf8()),
                        ),
                        true,
                    )? {
                        break;
                    }
                }
            }
        }

        self.push(TokenKind::Eof, self.idx, self.idx);
        self.post_process_intersections();
        Ok((self.tokens, first_error))
    }

    fn post_process_intersections(&mut self) {
        let mut i = 0;
        while i < self.tokens.len() {
            let should_intersect = if let TokenKind::Whitespace(raw) = &self.tokens[i].kind {
                let prev = prev_significant(&self.tokens, i);
                let next = next_significant(&self.tokens, i);
                if let (Some(p), Some(n)) = (prev, next) {
                    is_intersect_operand(&self.tokens[p].kind)
                        && is_intersect_operand(&self.tokens[n].kind)
                        && raw.chars().any(|c| c == ' ' || c == '\t')
                } else {
                    false
                }
            } else {
                false
            };

            if should_intersect {
                let raw = match &mut self.tokens[i].kind {
                    TokenKind::Whitespace(raw) => std::mem::take(raw),
                    _ => unreachable!("should_intersect requires Whitespace token"),
                };
                self.tokens[i].kind = TokenKind::Intersect(raw);
            }
            i += 1;
        }
    }

    fn push(&mut self, kind: TokenKind, start: usize, end: usize) {
        let sig = !matches!(kind, TokenKind::Whitespace(_));
        if sig {
            self.prev_sig = Some(PrevSig::from_kind(&kind));
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

    fn rollback_to(&mut self, idx: usize) {
        self.idx = idx;
        self.chars = self.src[idx..].chars();
    }

    fn peek_char(&self) -> Option<char> {
        self.chars.clone().next()
    }

    fn peek_next_is_digit(&self) -> bool {
        let mut iter = self.chars.clone();
        iter.next();
        matches!(iter.next(), Some(c) if is_digit(c))
    }

    fn take_while<F>(&mut self, pred: F) -> String
    where
        F: FnMut(char) -> bool,
    {
        let mut out = String::new();
        self.take_while_into(pred, &mut out);
        out
    }

    fn take_while_into<F>(&mut self, mut pred: F, out: &mut String)
    where
        F: FnMut(char) -> bool,
    {
        while let Some(ch) = self.peek_char() {
            if !pred(ch) {
                break;
            }
            self.bump();
            out.push(ch);
        }
    }

    /// Determine which decimal separator (if any) should be used when lexing the current number.
    ///
    /// Rules:
    /// - Prefer the locale decimal separator when it appears anywhere in the mantissa.
    /// - Otherwise, accept canonical `.` decimals in any locale.
    /// - In locales where `.` is also the thousands separator (e.g. `de-DE`, `es-ES`), treat `.`
    ///   as a thousands separator (not a decimal point) when the mantissa matches a typical
    ///   thousands-grouping pattern like `1.234.567`.
    fn number_decimal_separator(&self) -> Option<char> {
        let start = self.idx;
        let mut end = start;

        for (rel, ch) in self.src[start..].char_indices() {
            if matches!(ch, 'E' | 'e') {
                break;
            }
            // Some locales (notably fr-FR) commonly use NBSP (U+00A0) for thousands grouping, but
            // narrow NBSP (U+202F) also appears in spreadsheets. When configured for either,
            // accept both while scanning the mantissa so we can still detect the decimal
            // separator later in the literal.
            let is_thousands_sep =
                LocaleConfig::matches_thousands_separator(self.locale.thousands_separator, ch);
            if is_digit(ch) || ch == self.locale.decimal_separator || ch == '.' || is_thousands_sep
            {
                end = start + rel + ch.len_utf8();
                continue;
            }
            break;
        }

        if end <= start {
            return None;
        }

        let mantissa = &self.src[start..end];

        if mantissa.contains(self.locale.decimal_separator) {
            return Some(self.locale.decimal_separator);
        }

        if self.locale.decimal_separator != '.' && mantissa.contains('.') {
            // Disambiguate locales where the thousands separator collides with the canonical
            // decimal separator.
            if self.locale.thousands_separator == Some('.')
                && looks_like_thousands_grouping(mantissa, '.')
            {
                return None;
            }
            return Some('.');
        }

        None
    }

    fn lex_number(&mut self) -> String {
        let decimal_sep = self.number_decimal_separator();
        let group_sep = match (decimal_sep, self.locale.thousands_separator) {
            (Some(dec), Some(group)) if dec == group => None,
            _ => self.locale.thousands_separator,
        };

        let mut out = String::new();
        // integer / leading decimal
        while let Some(ch) = self.peek_char() {
            if is_digit(ch) {
                self.bump();
                out.push(ch);
                continue;
            }

            // Locale-specific grouping separators inside the integer portion of the literal.
            //
            // Note: Some locales (notably fr-FR) commonly use NBSP (U+00A0) as the grouping
            // separator, but some spreadsheets may contain the narrow no-break space (U+202F)
            // instead. When configured for either, accept both.
            let is_thousands_sep = LocaleConfig::matches_thousands_separator(group_sep, ch);
            if is_thousands_sep && !out.is_empty() && self.peek_next_is_digit() {
                self.bump();
                continue;
            }

            break;
        }
        if decimal_sep.is_some_and(|dec| self.peek_char() == Some(dec)) {
            self.bump();
            out.push(decimal_sep.expect("is_some_and ensured decimal_sep is Some"));
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
            let save_out_len = out.len();
            self.bump();
            out.push('E');
            if matches!(self.peek_char(), Some('+' | '-')) {
                let sign = self.bump().unwrap();
                out.push(sign);
            }
            let digits_start_len = out.len();
            while let Some(ch) = self.peek_char() {
                if is_digit(ch) {
                    self.bump();
                    out.push(ch);
                } else {
                    break;
                }
            }
            if out.len() == digits_start_len {
                // roll back: the 'E' was part of an identifier maybe.
                self.rollback_to(save_idx);
                out.truncate(save_out_len);
            }
        }
        out
    }

    fn lex_ident(&mut self) -> String {
        self.take_while(is_ident_cont_char)
    }

    fn may_be_a1_cell_ref(&self) -> bool {
        let bytes = self.src.as_bytes();
        let mut i = self.idx;
        if i >= bytes.len() {
            return false;
        }

        if bytes[i] == b'$' {
            i += 1;
        }

        let mut letters = 0usize;
        while i < bytes.len() && bytes[i].is_ascii_alphabetic() {
            letters += 1;
            if letters > 3 {
                return false;
            }
            i += 1;
        }
        if letters == 0 {
            return false;
        }

        if i < bytes.len() && bytes[i] == b'$' {
            i += 1;
        }

        i < bytes.len() && bytes[i].is_ascii_digit()
    }

    fn try_lex_cell_ref(&mut self) -> Option<CellToken> {
        let save_idx = self.idx;

        let mut col_abs = false;
        if self.peek_char() == Some('$') {
            col_abs = true;
            self.bump();
        }
        let col_start = self.idx;
        while let Some(ch) = self.peek_char() {
            if matches!(ch, 'A'..='Z' | 'a'..='z') {
                self.bump();
            } else {
                break;
            }
        }
        let col_end = self.idx;
        if col_start == col_end {
            self.rollback_to(save_idx);
            return None;
        }
        let mut row_abs = false;
        if self.peek_char() == Some('$') {
            row_abs = true;
            self.bump();
        }
        let row_start = self.idx;
        while let Some(ch) = self.peek_char() {
            if is_digit(ch) {
                self.bump();
            } else {
                break;
            }
        }
        let row_end = self.idx;
        if row_start == row_end {
            self.rollback_to(save_idx);
            return None;
        }

        // Avoid mis-lexing identifiers that start with an A1 reference prefix (e.g. `A1FOO`).
        //
        // Excel allows defined names like `A1FOO` because they do not *fully* match the A1 cell
        // reference grammar. If we accept the `A1` prefix as a cell token, the remaining `FOO`
        // becomes an adjacent identifier token which is invalid formula syntax and results in
        // confusing parse errors.
        // If the next character continues an identifier (e.g. `A1FOO`), treat this as a name
        // rather than a cell reference to avoid confusing parse errors.
        //
        // Special case: allow `.` so we can parse field access expressions like `A1.Price`.
        if matches!(self.peek_char(), Some(c) if (is_ident_cont_char(c) && c != '.') || c == '(') {
            self.rollback_to(save_idx);
            return None;
        }

        let col_letters = &self.src[col_start..col_end];
        let Ok(col) = column_label_to_index_lenient(col_letters) else {
            self.rollback_to(save_idx);
            return None;
        };
        let row_digits = &self.src[row_start..row_end];
        let Some(row) = row_digits.parse::<u32>().ok() else {
            self.rollback_to(save_idx);
            return None;
        };
        if row == 0 {
            self.rollback_to(save_idx);
            return None;
        }
        Some(CellToken {
            col,
            row: row - 1,
            col_abs,
            row_abs,
        })
    }

    fn try_lex_r1c1_cell_ref(&mut self) -> Option<R1C1CellToken> {
        let ch = self.peek_char()?;
        if !matches!(ch, 'R' | 'r') {
            return None;
        }
        let save_idx = self.idx;
        self.bump(); // R

        let row = match self.peek_char() {
            Some('[') => {
                let offset = self
                    .lex_r1c1_offset_in_brackets()
                    .or_else(|| {
                        self.rollback_to(save_idx);
                        None
                    })?;
                Coord::Offset(offset)
            }
            Some(c) if is_digit(c) => {
                let raw = self.take_while(is_digit);
                let row_1: u32 = match raw.parse().ok() {
                    Some(v) => v,
                    None => {
                        self.rollback_to(save_idx);
                        return None;
                    }
                };
                if row_1 == 0 {
                    self.rollback_to(save_idx);
                    return None;
                }
                Coord::A1 {
                    index: row_1 - 1,
                    abs: true,
                }
            }
            _ => Coord::Offset(0),
        };

        let Some(ch) = self.peek_char() else {
            self.rollback_to(save_idx);
            return None;
        };
        if !matches!(ch, 'C' | 'c') {
            self.rollback_to(save_idx);
            return None;
        }
        self.bump(); // C

        let col = match self.peek_char() {
            Some('[') => {
                let offset = self
                    .lex_r1c1_offset_in_brackets()
                    .or_else(|| {
                        self.rollback_to(save_idx);
                        None
                    })?;
                Coord::Offset(offset)
            }
            Some(c) if is_digit(c) => {
                let raw = self.take_while(is_digit);
                let col_1: u32 = match raw.parse().ok() {
                    Some(v) => v,
                    None => {
                        self.rollback_to(save_idx);
                        return None;
                    }
                };
                if col_1 == 0 {
                    self.rollback_to(save_idx);
                    return None;
                }
                Coord::A1 {
                    index: col_1 - 1,
                    abs: true,
                }
            }
            _ => Coord::Offset(0),
        };

        // Avoid mis-lexing identifiers that *start* with an R1C1 reference prefix.
        //
        // In R1C1 mode, valid references like `RC` and `R1C1` can appear as prefixes of valid
        // identifiers (e.g. `RCAR`, `R1C1FOO`). Excel allows such names because they do not *fully*
        // match the R1C1 cell-reference grammar.
        //
        // If we accept the prefix as a cell token we would end up with adjacency like
        // `RC` + `AR`, which is not valid formula syntax and causes confusing parse errors.
        // Instead, reject the cell token when the next character would continue an identifier or
        // start a function call, so the full string is lexed as an identifier.
        if matches!(self.peek_char(), Some(c) if (is_ident_cont_char(c) && c != '.') || c == '(') {
            self.rollback_to(save_idx);
            return None;
        }

        Some(R1C1CellToken { row, col })
    }

    fn try_lex_r1c1_row_ref(&mut self) -> Option<R1C1RowToken> {
        let ch = self.peek_char()?;
        if !matches!(ch, 'R' | 'r') {
            return None;
        }
        let save_idx = self.idx;
        self.bump(); // R

        let row = match self.peek_char() {
            Some('[') => {
                let offset = self
                    .lex_r1c1_offset_in_brackets()
                    .or_else(|| {
                        self.rollback_to(save_idx);
                        None
                    })?;
                Coord::Offset(offset)
            }
            Some(c) if is_digit(c) => {
                let raw = self.take_while(is_digit);
                let row_1: u32 = raw.parse().ok()?;
                if row_1 == 0 {
                    self.rollback_to(save_idx);
                    return None;
                }
                Coord::A1 {
                    index: row_1 - 1,
                    abs: true,
                }
            }
            _ => Coord::Offset(0),
        };

        if matches!(self.peek_char(), Some(c) if (is_ident_cont_char(c) && c != '.') || c == '(') {
            self.rollback_to(save_idx);
            return None;
        }

        Some(R1C1RowToken { row })
    }

    fn try_lex_r1c1_col_ref(&mut self) -> Option<R1C1ColToken> {
        let ch = self.peek_char()?;
        if !matches!(ch, 'C' | 'c') {
            return None;
        }
        let save_idx = self.idx;
        self.bump(); // C

        let col = match self.peek_char() {
            Some('[') => {
                let offset = self
                    .lex_r1c1_offset_in_brackets()
                    .or_else(|| {
                        self.rollback_to(save_idx);
                        None
                    })?;
                Coord::Offset(offset)
            }
            Some(c) if is_digit(c) => {
                let raw = self.take_while(is_digit);
                let col_1: u32 = raw.parse().ok()?;
                if col_1 == 0 {
                    self.rollback_to(save_idx);
                    return None;
                }
                Coord::A1 {
                    index: col_1 - 1,
                    abs: true,
                }
            }
            _ => Coord::Offset(0),
        };

        if matches!(self.peek_char(), Some(c) if (is_ident_cont_char(c) && c != '.') || c == '(') {
            self.rollback_to(save_idx);
            return None;
        }

        Some(R1C1ColToken { col })
    }

    fn lex_r1c1_offset_in_brackets(&mut self) -> Option<i32> {
        debug_assert_eq!(self.peek_char(), Some('['));
        self.bump(); // '['
        let sign = match self.peek_char() {
            Some('+') => {
                self.bump();
                1i64
            }
            Some('-') => {
                self.bump();
                -1i64
            }
            _ => 1i64,
        };
        let digits = self.take_while(is_digit);
        if digits.is_empty() {
            return None;
        }
        if self.peek_char() != Some(']') {
            return None;
        }
        self.bump(); // ']'

        // Offsets are stored as i32s in the AST/IR. Parse via i64 so we can accept the full i32
        // range, including `-2147483648` (which requires parsing a magnitude of `2147483648`).
        let mag: i64 = digits.parse().ok()?;
        let value = sign.checked_mul(mag)?;
        if value < i64::from(i32::MIN) || value > i64::from(i32::MAX) {
            return None;
        }
        Some(value as i32)
    }
}

fn is_digit(c: char) -> bool {
    matches!(c, '0'..='9')
}

fn is_ident_start_char(c: char) -> bool {
    matches!(c, '$' | '_' | '\\' | 'A'..='Z' | 'a'..='z') || (!c.is_ascii() && c.is_alphabetic())
}

fn is_ident_cont_char(c: char) -> bool {
    matches!(
        c,
        '$' | '_' | '\\' | '.' | 'A'..='Z' | 'a'..='z' | '0'..='9'
    ) || (!c.is_ascii() && c.is_alphanumeric())
}

fn looks_like_thousands_grouping(raw: &str, sep: char) -> bool {
    let mut parts = raw.split(sep);
    let Some(first) = parts.next() else {
        return false;
    };
    if first.is_empty() || first.len() > 3 || !first.chars().all(|c| c.is_ascii_digit()) {
        return false;
    }

    let mut saw_sep = false;
    for part in parts {
        saw_sep = true;
        if part.len() != 3 || !part.chars().all(|c| c.is_ascii_digit()) {
            return false;
        }
    }

    saw_sep
}

const ERROR_LITERALS: &[&str] = &[
    "#NULL!",
    "#DIV/0!",
    "#VALUE!",
    "#REF!",
    "#NAME?",
    "#NUM!",
    "#N/A",
    "#N/A!",
    "#GETTING_DATA",
    "#SPILL!",
    "#CALC!",
    "#FIELD!",
    "#CONNECT!",
    "#BLOCKED!",
    "#UNKNOWN!",
];

fn match_error_literal(input: &str) -> Option<usize> {
    let mut best: Option<usize> = None;
    for &lit in ERROR_LITERALS {
        if input
            .get(..lit.len())
            .is_some_and(|prefix| prefix.eq_ignore_ascii_case(lit))
        {
            best = Some(best.map_or(lit.len(), |cur| cur.max(lit.len())));
        }
    }
    best
}

fn is_error_body_char(c: char) -> bool {
    // Error literals start with `#` and are followed by a locale-dependent name that can include
    // non-ASCII letters (e.g. `#ÜBERLAUF!`) and, in some locales, inverted punctuation (e.g.
    // `#¡VALOR!`, `#¿NOMBRE?`).
    //
    // We treat the error "body" as a superset of identifier-continue characters plus a small set
    // of ASCII punctuation used by canonical error names.
    matches!(c, '_' | '/' | '.' | '¡' | '¿') || unicode_ident::is_xid_continue(c)
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
            | TokenKind::R1C1Cell(_)
            | TokenKind::R1C1Row(_)
            | TokenKind::R1C1Col(_)
            | TokenKind::Ident(_)
            | TokenKind::QuotedIdent(_)
            | TokenKind::RParen
            | TokenKind::RBracket
            | TokenKind::Hash
    )
}

struct Parser<'a> {
    src: &'a str,
    tokens: Vec<Token>,
    pos: usize,
    func_stack: Vec<(String, usize)>,
    call_depth: usize,
    group_depth: usize,
    unary_depth: usize,
    pow_depth: usize,
    array_depth: usize,
    first_error: Option<ParseError>,
}

impl<'a> Parser<'a> {
    fn new(src: &'a str, tokens: Vec<Token>) -> Self {
        Self {
            src,
            tokens,
            pos: 0,
            func_stack: Vec::new(),
            call_depth: 0,
            group_depth: 0,
            unary_depth: 0,
            pow_depth: 0,
            array_depth: 0,
            first_error: None,
        }
    }

    fn parse_expression(&mut self, min_bp: u8) -> Result<Expr, ParseError> {
        self.skip_trivia();
        let mut lhs = self.parse_prefix()?;

        loop {
            self.skip_trivia();
            // Postfix call expressions: `expr(arg1, arg2, ...)` (e.g. `LAMBDA(x,x+1)(5)`).
            let call_bp = 90;
            if matches!(self.peek_kind(), TokenKind::LParen) && call_bp >= min_bp {
                lhs = self.parse_call(lhs)?;
                continue;
            }

            // Postfix field access: `expr.Field` / `expr.["Field Name"]`.
            let field_bp = 90;
            if matches!(self.peek_kind(), TokenKind::Dot) && field_bp >= min_bp {
                lhs = self.parse_field_access(lhs)?;
                continue;
            }

            // Postfix operators (`%` and spill-range `#`).
            let postfix_bp = 60;
            if matches!(self.peek_kind(), TokenKind::Percent) && postfix_bp >= min_bp {
                self.next();
                lhs = Expr::Postfix(PostfixExpr {
                    op: PostfixOp::Percent,
                    expr: Box::new(lhs),
                });
                continue;
            }
            if matches!(self.peek_kind(), TokenKind::Hash) && postfix_bp >= min_bp {
                self.next();
                lhs = Expr::Postfix(PostfixExpr {
                    op: PostfixOp::SpillRange,
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
            if op == BinaryOp::Pow && self.pow_depth >= EXCEL_MAX_NESTED_CALLS {
                return Err(ParseError::new(
                    format!(
                        "Expression nesting exceeds Excel's {EXCEL_MAX_NESTED_CALLS}-level limit"
                    ),
                    self.current_span(),
                ));
            }
            self.next(); // consume operator
            let rhs = if op == BinaryOp::Pow {
                self.pow_depth += 1;
                let result = self.parse_expression(r_bp);
                self.pow_depth = self.pow_depth.saturating_sub(1);
                result?
            } else {
                self.parse_expression(r_bp)?
            };
            let (left, right) = if op == BinaryOp::Range {
                coerce_range_operands(lhs, rhs)
            } else {
                (lhs, rhs)
            };
            lhs = Expr::Binary(BinaryExpr {
                op,
                left: Box::new(left),
                right: Box::new(right),
            });
        }

        Ok(lhs)
    }

    fn take_context(&mut self) -> ParseContext {
        let function = self
            .func_stack
            .pop()
            .map(|(name, arg_index)| FunctionContext { name, arg_index });
        ParseContext { function }
    }

    fn record_error(&mut self, err: ParseError) {
        if self.first_error.is_none() {
            self.first_error = Some(err);
        }
    }

    /// Consume tokens until the matching closing `)` for the current parenthesized group/call.
    ///
    /// Assumes the opening `(` has already been consumed.
    ///
    /// Returns `true` if a matching `)` was found and consumed, or `false` if EOF was reached
    /// first.
    fn consume_until_matching_rparen(&mut self) -> bool {
        let mut depth: usize = 1;
        while depth > 0 {
            match self.peek_kind() {
                TokenKind::LParen => {
                    depth += 1;
                    self.next();
                }
                TokenKind::RParen => {
                    depth = depth.saturating_sub(1);
                    self.next();
                }
                TokenKind::Eof => return false,
                _ => {
                    self.next();
                }
            }
        }
        true
    }

    fn parse_expression_best_effort(&mut self, min_bp: u8) -> Expr {
        self.skip_trivia();
        let mut lhs = self.parse_prefix_best_effort();

        loop {
            self.skip_trivia();
            // Postfix call expressions: `expr(arg1, arg2, ...)` (e.g. `LAMBDA(x,x+1)(5)`).
            let call_bp = 90;
            if matches!(self.peek_kind(), TokenKind::LParen) && call_bp >= min_bp {
                lhs = self.parse_call_best_effort(lhs);
                continue;
            }

            // Postfix field access: `expr.Field` / `expr.["Field Name"]`.
            let field_bp = 90;
            if matches!(self.peek_kind(), TokenKind::Dot) && field_bp >= min_bp {
                lhs = self.parse_field_access_best_effort(lhs);
                continue;
            }

            // Postfix operators (`%` and spill-range `#`).
            let postfix_bp = 60;
            if matches!(self.peek_kind(), TokenKind::Percent) && postfix_bp >= min_bp {
                self.next();
                lhs = Expr::Postfix(PostfixExpr {
                    op: PostfixOp::Percent,
                    expr: Box::new(lhs),
                });
                continue;
            }
            if matches!(self.peek_kind(), TokenKind::Hash) && postfix_bp >= min_bp {
                self.next();
                lhs = Expr::Postfix(PostfixExpr {
                    op: PostfixOp::SpillRange,
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
            if op == BinaryOp::Pow && self.pow_depth >= EXCEL_MAX_NESTED_CALLS {
                self.record_error(ParseError::new(
                    format!(
                        "Expression nesting exceeds Excel's {EXCEL_MAX_NESTED_CALLS}-level limit"
                    ),
                    self.current_span(),
                ));
                break;
            }
            self.next(); // consume operator
            let rhs = if op == BinaryOp::Pow {
                self.pow_depth += 1;
                let rhs = self.parse_expression_best_effort(r_bp);
                self.pow_depth = self.pow_depth.saturating_sub(1);
                rhs
            } else {
                self.parse_expression_best_effort(r_bp)
            };
            let (left, right) = if op == BinaryOp::Range {
                coerce_range_operands(lhs, rhs)
            } else {
                (lhs, rhs)
            };
            lhs = Expr::Binary(BinaryExpr {
                op,
                left: Box::new(left),
                right: Box::new(right),
            });
        }

        lhs
    }

    fn parse_prefix_best_effort(&mut self) -> Expr {
        self.skip_trivia();
        match self.peek_kind() {
            TokenKind::Plus => {
                if self.unary_depth >= EXCEL_MAX_NESTED_CALLS {
                    self.record_error(ParseError::new(
                        format!(
                            "Expression nesting exceeds Excel's {EXCEL_MAX_NESTED_CALLS}-level limit"
                        ),
                        self.current_span(),
                    ));
                    self.next();
                    return Expr::Missing;
                }
                self.next();
                self.unary_depth += 1;
                let expr = self.parse_expression_best_effort(50);
                self.unary_depth = self.unary_depth.saturating_sub(1);
                Expr::Unary(UnaryExpr {
                    op: UnaryOp::Plus,
                    expr: Box::new(expr),
                })
            }
            TokenKind::Minus => {
                if self.unary_depth >= EXCEL_MAX_NESTED_CALLS {
                    self.record_error(ParseError::new(
                        format!(
                            "Expression nesting exceeds Excel's {EXCEL_MAX_NESTED_CALLS}-level limit"
                        ),
                        self.current_span(),
                    ));
                    self.next();
                    return Expr::Missing;
                }
                self.next();
                self.unary_depth += 1;
                let expr = self.parse_expression_best_effort(50);
                self.unary_depth = self.unary_depth.saturating_sub(1);
                Expr::Unary(UnaryExpr {
                    op: UnaryOp::Minus,
                    expr: Box::new(expr),
                })
            }
            TokenKind::At => {
                if self.unary_depth >= EXCEL_MAX_NESTED_CALLS {
                    self.record_error(ParseError::new(
                        format!(
                            "Expression nesting exceeds Excel's {EXCEL_MAX_NESTED_CALLS}-level limit"
                        ),
                        self.current_span(),
                    ));
                    self.next();
                    return Expr::Missing;
                }
                self.next();
                self.unary_depth += 1;
                let expr = self.parse_expression_best_effort(50);
                self.unary_depth = self.unary_depth.saturating_sub(1);
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
            TokenKind::Number(_) => Expr::Number(self.take_number_token_unchecked()),
            TokenKind::String(_) => Expr::String(self.take_string_token_unchecked()),
            TokenKind::Boolean(v) => {
                let v = *v;
                self.next();
                Expr::Boolean(v)
            }
            TokenKind::Error(_) => Expr::Error(self.take_error_token_unchecked()),
            TokenKind::LParen => {
                if self.group_depth >= EXCEL_MAX_NESTED_CALLS {
                    self.record_error(ParseError::new(
                        format!(
                            "Expression nesting exceeds Excel's {EXCEL_MAX_NESTED_CALLS}-level limit"
                        ),
                        self.current_span(),
                    ));
                    // Consume the '(' to avoid infinite loops.
                    self.next();
                    return Expr::Missing;
                }
                self.next();
                self.group_depth += 1;
                let expr = self.parse_expression_best_effort(0);
                if let Err(e) = self.expect(TokenKind::RParen) {
                    self.record_error(e);
                }
                self.group_depth = self.group_depth.saturating_sub(1);
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
            TokenKind::Cell(_)
            | TokenKind::R1C1Cell(_)
            | TokenKind::R1C1Row(_)
            | TokenKind::R1C1Col(_)
            | TokenKind::Ident(_)
            | TokenKind::QuotedIdent(_) => self.parse_reference_or_name_or_func_best_effort(),
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

        // Optional sheet prefix:
        // - Sheet1!A1 / 'My Sheet'!A1
        // - Sheet1:Sheet3!A1 / 'Sheet 1':'Sheet 3'!A1
        let save_pos = self.pos;
        let sheet_prefix = match self.peek_kind() {
            TokenKind::Ident(_) | TokenKind::QuotedIdent(_) => {
                if !self.looks_like_sheet_prefix(save_pos, true) {
                    None
                } else {
                    let complete_prefix = self.looks_like_sheet_prefix(save_pos, false);

                    let start_raw = if complete_prefix {
                        self.take_name_token_unchecked()
                    } else {
                        match self.take_name_token() {
                            Ok(s) => s,
                            Err(e) => {
                                self.record_error(e);
                                return Expr::Missing;
                            }
                        }
                    };
                self.skip_trivia();
                if matches!(self.peek_kind(), TokenKind::Colon) {
                    // Sheet span.
                    self.next();
                    self.skip_trivia();
                    let end_raw = if complete_prefix {
                        self.take_name_token_unchecked()
                    } else {
                        match self.take_name_token() {
                            Ok(s) => s,
                            Err(e) => {
                                self.record_error(e);
                                self.pos = save_pos;
                                return Expr::Missing;
                            }
                        }
                    };
                    self.skip_trivia();
                    if matches!(self.peek_kind(), TokenKind::Bang) {
                        self.next();
                        let (workbook, start) = split_external_sheet_name_parts(&start_raw);
                        let (_wb2, end) = split_external_sheet_name_parts(&end_raw);
                        let sheet_ref = if sheet_name_eq_case_insensitive(start.as_ref(), end.as_ref())
                        {
                            SheetRef::Sheet(start.into_owned())
                        } else {
                            SheetRef::SheetRange {
                                start: start.into_owned(),
                                end: end.into_owned(),
                            }
                        };
                        Some((workbook.map(Cow::into_owned), sheet_ref))
                    } else {
                        self.pos = save_pos;
                        None
                    }
                } else if matches!(self.peek_kind(), TokenKind::Bang) {
                    self.next();
                    let (workbook, sheet_ref) = sheet_ref_from_raw_prefix(&start_raw);
                    Some((workbook, sheet_ref))
                } else {
                    self.pos = save_pos;
                    None
                }
                }
            }
            _ => None,
        };

        if let Some((workbook, sheet)) = sheet_prefix {
            return match self.parse_ref_after_prefix(workbook, Some(sheet)) {
                Ok(e) => e,
                Err(err) => {
                    self.record_error(err);
                    Expr::Missing
                }
            };
        }

        match self.peek_kind() {
            TokenKind::Ident(_) => {
                let name = self.take_ident_token_unchecked();
                self.skip_trivia();
                if matches!(self.peek_kind(), TokenKind::LParen) {
                    self.parse_function_call_best_effort(name)
                } else if matches!(self.peek_kind(), TokenKind::LBracket) {
                    match self.parse_structured_ref(None, None, Some(name)) {
                        Ok(expr) => expr,
                        Err(err) => {
                            self.record_error(err);
                            Expr::Missing
                        }
                    }
                } else {
                    Expr::NameRef(NameRef {
                        workbook: None,
                        sheet: None,
                        name,
                    })
                }
            }
            TokenKind::Cell(cell) => {
                let cell = *cell;
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
            TokenKind::R1C1Cell(cell) => {
                let cell = *cell;
                self.next();
                Expr::CellRef(CellRef {
                    workbook: None,
                    sheet: None,
                    col: cell.col,
                    row: cell.row,
                })
            }
            TokenKind::R1C1Row(row) => {
                let row = *row;
                self.next();
                Expr::RowRef(RowRef {
                    workbook: None,
                    sheet: None,
                    row: row.row,
                })
            }
            TokenKind::R1C1Col(col) => {
                let col = *col;
                self.next();
                Expr::ColRef(ColRef {
                    workbook: None,
                    sheet: None,
                    col: col.col,
                })
            }
            TokenKind::QuotedIdent(_name) => {
                let raw = self.take_name_token_unchecked();
                let (workbook, name) = split_external_sheet_name(&raw);
                Expr::NameRef(NameRef {
                    workbook,
                    sheet: None,
                    name,
                })
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
        if self.call_depth >= EXCEL_MAX_NESTED_CALLS {
            self.record_error(ParseError::new(
                format!("Function nesting exceeds Excel's {EXCEL_MAX_NESTED_CALLS}-level limit"),
                self.current_span(),
            ));
            // Consume the `(` (if present) to make progress and avoid deep recursion.
            if matches!(self.peek_kind(), TokenKind::LParen) {
                self.next();
            }
            return Expr::Missing;
        }

        if let Err(e) = self.expect(TokenKind::LParen) {
            self.record_error(e);
            return Expr::Missing;
        }

        self.call_depth += 1;
        self.func_stack.push((name, 0));
        let mut args = Vec::new();
        let mut should_pop_stack = false;

        loop {
            self.skip_trivia();
            match self.peek_kind() {
                TokenKind::RParen => {
                    self.next();
                    should_pop_stack = true;
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

            if args.len() == crate::EXCEL_MAX_ARGS {
                self.record_error(ParseError::new(
                    format!("Too many arguments (max {})", crate::EXCEL_MAX_ARGS),
                    self.current_span(),
                ));
                let closed = self.consume_until_matching_rparen();
                should_pop_stack = closed;
                break;
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
                    should_pop_stack = true;
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

        let name = if should_pop_stack {
            let (name, _) = self
                .func_stack
                .pop()
                .expect("parse_function_call_best_effort should balance func_stack");
            name
        } else {
            self.func_stack
                .last()
                .map(|(name, _idx)| name.clone())
                .expect("parse_function_call_best_effort should push func_stack entry")
        };
        let out = Expr::FunctionCall(FunctionCall {
            name: FunctionName::new(name),
            args,
        });
        self.call_depth = self.call_depth.saturating_sub(1);
        out
    }

    fn parse_call_best_effort(&mut self, callee: Expr) -> Expr {
        if self.call_depth >= EXCEL_MAX_NESTED_CALLS {
            self.record_error(ParseError::new(
                format!("Function nesting exceeds Excel's {EXCEL_MAX_NESTED_CALLS}-level limit"),
                self.current_span(),
            ));
            // Consume the `(` (if present) to make progress and avoid deep recursion.
            if matches!(self.peek_kind(), TokenKind::LParen) {
                self.next();
            }
            return Expr::Missing;
        }

        if let Err(e) = self.expect(TokenKind::LParen) {
            self.record_error(e);
            return Expr::Missing;
        }

        self.call_depth += 1;
        let mut args = Vec::new();

        loop {
            self.skip_trivia();
            match self.peek_kind() {
                TokenKind::RParen => {
                    self.next();
                    break;
                }
                TokenKind::Eof => {
                    self.record_error(ParseError::new("Unterminated call", self.current_span()));
                    break;
                }
                _ => {}
            }

            if args.len() == crate::EXCEL_MAX_ARGS {
                self.record_error(ParseError::new(
                    format!("Too many arguments (max {})", crate::EXCEL_MAX_ARGS),
                    self.current_span(),
                ));
                self.consume_until_matching_rparen();
                break;
            }

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
                    continue;
                }
                TokenKind::RParen => {
                    self.next();
                    break;
                }
                TokenKind::Eof => {
                    self.record_error(ParseError::new("Unterminated call", self.current_span()));
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

        let out = Expr::Call(CallExpr {
            callee: Box::new(callee),
            args,
        });
        self.call_depth = self.call_depth.saturating_sub(1);
        out
    }

    fn parse_array_literal_best_effort(&mut self) -> Expr {
        if self.array_depth >= EXCEL_MAX_NESTED_CALLS {
            self.record_error(ParseError::new(
                format!("Expression nesting exceeds Excel's {EXCEL_MAX_NESTED_CALLS}-level limit"),
                self.current_span(),
            ));
            // Consume the `{` (if present) to avoid infinite loops.
            if matches!(self.peek_kind(), TokenKind::LBrace) {
                self.next();
            }
            return Expr::Missing;
        }

        if let Err(e) = self.expect(TokenKind::LBrace) {
            self.record_error(e);
            return Expr::Missing;
        }
        self.array_depth += 1;
        let mut rows: Vec<Vec<Expr>> = Vec::new();
        let mut current_row: Vec<Expr> = Vec::new();
        let mut expecting_value = true;
        loop {
            self.skip_trivia();
            match self.peek_kind() {
                TokenKind::RBrace => {
                    self.next();
                    if expecting_value && (!current_row.is_empty() || !rows.is_empty()) {
                        current_row.push(Expr::Missing);
                    }
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
                    if expecting_value && (!current_row.is_empty() || !rows.is_empty()) {
                        current_row.push(Expr::Missing);
                    }
                    if !current_row.is_empty() || !rows.is_empty() {
                        rows.push(current_row);
                    }
                    break;
                }
                TokenKind::ArrayColSep => {
                    current_row.push(Expr::Missing);
                    self.next();
                    expecting_value = true;
                    continue;
                }
                TokenKind::ArrayRowSep => {
                    current_row.push(Expr::Missing);
                    self.next();
                    rows.push(current_row);
                    current_row = Vec::new();
                    expecting_value = true;
                    continue;
                }
                _ => {}
            }

            let el = self.parse_expression_best_effort(0);
            current_row.push(el);
            expecting_value = false;
            self.skip_trivia();
            match self.peek_kind() {
                TokenKind::ArrayColSep => {
                    self.next();
                    expecting_value = true;
                }
                TokenKind::ArrayRowSep => {
                    self.next();
                    rows.push(current_row);
                    current_row = Vec::new();
                    expecting_value = true;
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
        self.array_depth = self.array_depth.saturating_sub(1);
        Expr::Array(ArrayLiteral { rows })
    }

    fn parse_prefix(&mut self) -> Result<Expr, ParseError> {
        self.skip_trivia();
        match self.peek_kind() {
            TokenKind::Plus => {
                if self.unary_depth >= EXCEL_MAX_NESTED_CALLS {
                    return Err(ParseError::new(
                        format!(
                            "Expression nesting exceeds Excel's {EXCEL_MAX_NESTED_CALLS}-level limit"
                        ),
                        self.current_span(),
                    ));
                }
                self.next();
                self.unary_depth += 1;
                let result = self.parse_expression(50);
                self.unary_depth = self.unary_depth.saturating_sub(1);
                let expr = result?;
                Ok(Expr::Unary(UnaryExpr {
                    op: UnaryOp::Plus,
                    expr: Box::new(expr),
                }))
            }
            TokenKind::Minus => {
                if self.unary_depth >= EXCEL_MAX_NESTED_CALLS {
                    return Err(ParseError::new(
                        format!(
                            "Expression nesting exceeds Excel's {EXCEL_MAX_NESTED_CALLS}-level limit"
                        ),
                        self.current_span(),
                    ));
                }
                self.next();
                self.unary_depth += 1;
                let result = self.parse_expression(50);
                self.unary_depth = self.unary_depth.saturating_sub(1);
                let expr = result?;
                Ok(Expr::Unary(UnaryExpr {
                    op: UnaryOp::Minus,
                    expr: Box::new(expr),
                }))
            }
            TokenKind::At => {
                if self.unary_depth >= EXCEL_MAX_NESTED_CALLS {
                    return Err(ParseError::new(
                        format!(
                            "Expression nesting exceeds Excel's {EXCEL_MAX_NESTED_CALLS}-level limit"
                        ),
                        self.current_span(),
                    ));
                }
                self.next();
                self.unary_depth += 1;
                let result = self.parse_expression(50);
                self.unary_depth = self.unary_depth.saturating_sub(1);
                let expr = result?;
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
            TokenKind::Number(_) => Ok(Expr::Number(self.take_number_token_unchecked())),
            TokenKind::String(_) => Ok(Expr::String(self.take_string_token_unchecked())),
            TokenKind::Boolean(v) => {
                let v = *v;
                self.next();
                Ok(Expr::Boolean(v))
            }
            TokenKind::Error(_) => Ok(Expr::Error(self.take_error_token_unchecked())),
            TokenKind::LParen => {
                if self.group_depth >= EXCEL_MAX_NESTED_CALLS {
                    return Err(ParseError::new(
                        format!(
                            "Expression nesting exceeds Excel's {EXCEL_MAX_NESTED_CALLS}-level limit"
                        ),
                        self.current_span(),
                    ));
                }
                self.next();
                self.group_depth += 1;
                let result = (|| {
                    let expr = self.parse_expression(0)?;
                    self.expect(TokenKind::RParen)?;
                    Ok(expr)
                })();
                self.group_depth = self.group_depth.saturating_sub(1);
                result
            }
            TokenKind::LBrace => self.parse_array_literal(),
            TokenKind::LBracket => self.parse_bracket_start(),
            TokenKind::Cell(_)
            | TokenKind::R1C1Cell(_)
            | TokenKind::R1C1Row(_)
            | TokenKind::R1C1Col(_)
            | TokenKind::Ident(_)
            | TokenKind::QuotedIdent(_) => self.parse_reference_or_name_or_func(),
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
                // Could be sheet prefix (if followed by `!`), or a function/name.
                //
                // Important: many formulas start with an identifier that is *not* a sheet name
                // (`SUM`, `LET`, a defined name, etc.). Avoid cloning the identifier string unless
                // we can prove this is a sheet prefix (we previously cloned and then rewound).
                if !self.looks_like_sheet_prefix(save_pos, false) {
                    (None, None)
                } else {
                    let start_raw = self.take_name_token_unchecked();
                        self.skip_trivia();

                        // Sheet span (3D ref) like `Sheet1:Sheet3!A1`.
                        if matches!(self.peek_kind(), TokenKind::Colon) {
                            self.next();
                            self.skip_trivia();
                            if !matches!(
                                self.peek_kind(),
                                TokenKind::Ident(_) | TokenKind::QuotedIdent(_)
                            ) {
                                self.pos = save_pos;
                                (None, None)
                            } else {
                                let end_raw = self.take_name_token_unchecked();
                                self.skip_trivia();
                                if !matches!(self.peek_kind(), TokenKind::Bang) {
                                    self.pos = save_pos;
                                    (None, None)
                                } else {
                                    self.next();
                                    let (workbook, start) =
                                        split_external_sheet_name_parts(&start_raw);
                                    let (_wb2, end) = split_external_sheet_name_parts(&end_raw);
                                    let sheet_ref = if sheet_name_eq_case_insensitive(
                                        start.as_ref(),
                                        end.as_ref(),
                                    ) {
                                        SheetRef::Sheet(start.into_owned())
                                    } else {
                                        SheetRef::SheetRange {
                                            start: start.into_owned(),
                                            end: end.into_owned(),
                                        }
                                    };
                                    (workbook.map(Cow::into_owned), Some(sheet_ref))
                                }
                            }
                        } else if matches!(self.peek_kind(), TokenKind::Bang) {
                            self.next();
                            let (workbook, sheet_ref) = sheet_ref_from_raw_prefix(&start_raw);
                            (workbook, Some(sheet_ref))
                        } else {
                            self.pos = save_pos;
                            (None, None)
                        }
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
            TokenKind::Ident(_) => {
                let name = self.take_ident_token_unchecked();
                self.skip_trivia();
                if matches!(self.peek_kind(), TokenKind::LParen) {
                    self.parse_function_call(name)
                } else if matches!(self.peek_kind(), TokenKind::LBracket) {
                    self.parse_structured_ref(None, None, Some(name))
                } else {
                    Ok(Expr::NameRef(NameRef {
                        workbook: None,
                        sheet: None,
                        name,
                    }))
                }
            }
            TokenKind::Cell(cell) => {
                let cell = *cell;
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
            TokenKind::R1C1Cell(cell) => {
                let cell = *cell;
                self.next();
                Ok(Expr::CellRef(CellRef {
                    workbook: None,
                    sheet: None,
                    col: cell.col,
                    row: cell.row,
                }))
            }
            TokenKind::R1C1Row(row) => {
                let row = *row;
                self.next();
                Ok(Expr::RowRef(RowRef {
                    workbook: None,
                    sheet: None,
                    row: row.row,
                }))
            }
            TokenKind::R1C1Col(col) => {
                let col = *col;
                self.next();
                Ok(Expr::ColRef(ColRef {
                    workbook: None,
                    sheet: None,
                    col: col.col,
                }))
            }
            TokenKind::QuotedIdent(_name) => {
                let raw = self.take_name_token_unchecked();
                let (workbook, name) = split_external_sheet_name(&raw);
                Ok(Expr::NameRef(NameRef {
                    workbook,
                    sheet: None,
                    name,
                }))
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
        sheet: Option<SheetRef>,
    ) -> Result<Expr, ParseError> {
        self.skip_trivia();
        match self.peek_kind() {
            TokenKind::Cell(cell) => {
                let cell = *cell;
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
            TokenKind::R1C1Cell(cell) => {
                let cell = *cell;
                self.next();
                Ok(Expr::CellRef(CellRef {
                    workbook,
                    sheet,
                    col: cell.col,
                    row: cell.row,
                }))
            }
            TokenKind::R1C1Row(row) => {
                let row = *row;
                self.next();
                Ok(Expr::RowRef(RowRef {
                    workbook,
                    sheet,
                    row: row.row,
                }))
            }
            TokenKind::R1C1Col(col) => {
                let col = *col;
                self.next();
                Ok(Expr::ColRef(ColRef {
                    workbook,
                    sheet,
                    col: col.col,
                }))
            }
            TokenKind::Number(_) => {
                let span = self.current_span();
                let raw = self.take_number_token_unchecked();
                let Some(row) = parse_row_number_literal(&raw) else {
                    return Err(ParseError::new("Invalid row reference", span));
                };
                Ok(Expr::RowRef(RowRef {
                    workbook,
                    sheet,
                    row: Coord::A1 {
                        index: row,
                        abs: false,
                    },
                }))
            }
            TokenKind::Ident(_) => {
                let name = self.take_ident_token_unchecked();
                self.skip_trivia();
                if matches!(self.peek_kind(), TokenKind::LBracket) {
                    self.parse_structured_ref(workbook, sheet, Some(name))
                } else {
                    Ok(Expr::NameRef(NameRef {
                        workbook,
                        sheet,
                        name,
                    }))
                }
            }
            TokenKind::LBracket => self.parse_structured_ref(workbook, sheet, None),
            _ => Err(ParseError::new(
                "Expected reference after sheet prefix",
                self.current_span(),
            )),
        }
    }

    fn parse_function_call(&mut self, name: String) -> Result<Expr, ParseError> {
        if self.call_depth >= EXCEL_MAX_NESTED_CALLS {
            return Err(ParseError::new(
                format!("Function nesting exceeds Excel's {EXCEL_MAX_NESTED_CALLS}-level limit"),
                self.current_span(),
            ));
        }

        self.expect(TokenKind::LParen)?;
        self.call_depth += 1;
        self.func_stack.push((name, 0));
        let result: Result<Vec<Expr>, ParseError> = (|| {
            let mut args = Vec::new();
            self.skip_trivia();
            if matches!(self.peek_kind(), TokenKind::RParen) {
                self.next();
            } else {
                loop {
                    self.skip_trivia();
                    if args.len() == crate::EXCEL_MAX_ARGS {
                        return Err(ParseError::new(
                            format!("Too many arguments (max {})", crate::EXCEL_MAX_ARGS),
                            self.current_span(),
                        ));
                    }
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
            Ok(args)
        })();

        let (name, _) = self
            .func_stack
            .pop()
            .expect("parse_function_call should balance func_stack");
        self.call_depth = self.call_depth.saturating_sub(1);
        result.map(|args| {
            Expr::FunctionCall(FunctionCall {
                name: FunctionName::new(name),
                args,
            })
        })
    }

    fn parse_call(&mut self, callee: Expr) -> Result<Expr, ParseError> {
        if self.call_depth >= EXCEL_MAX_NESTED_CALLS {
            return Err(ParseError::new(
                format!("Function nesting exceeds Excel's {EXCEL_MAX_NESTED_CALLS}-level limit"),
                self.current_span(),
            ));
        }

        self.expect(TokenKind::LParen)?;
        self.call_depth += 1;
        let result = (|| {
            let mut args = Vec::new();
            self.skip_trivia();
            if matches!(self.peek_kind(), TokenKind::RParen) {
                self.next();
            } else {
                loop {
                    self.skip_trivia();
                    if args.len() == crate::EXCEL_MAX_ARGS {
                        return Err(ParseError::new(
                            format!("Too many arguments (max {})", crate::EXCEL_MAX_ARGS),
                            self.current_span(),
                        ));
                    }
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
            Ok(Expr::Call(CallExpr {
                callee: Box::new(callee),
                args,
            }))
        })();

        self.call_depth = self.call_depth.saturating_sub(1);
        result
    }

    fn parse_field_access(&mut self, base: Expr) -> Result<Expr, ParseError> {
        let dot_span = self.current_span();
        self.expect(TokenKind::Dot)?;
        self.skip_trivia();

        match self.peek_kind() {
            TokenKind::Ident(_) => {
                let name = self.take_ident_token_unchecked();

                if name.is_empty() {
                    return Err(ParseError::new("Expected field name", dot_span));
                }

                if !name.contains('.') {
                    return Ok(Expr::FieldAccess(FieldAccessExpr {
                        base: Box::new(base),
                        field: name,
                    }));
                }

                let mut expr = base;
                for part in name.split('.') {
                    if part.is_empty() {
                        return Err(ParseError::new("Expected field name", dot_span));
                    }
                    expr = Expr::FieldAccess(FieldAccessExpr {
                        base: Box::new(expr),
                        field: part.to_string(),
                    });
                }
                Ok(expr)
            }
            TokenKind::LBracket => {
                self.expect(TokenKind::LBracket)?;
                let field = match &mut self.tokens[self.pos].kind {
                    TokenKind::Ident(s) => {
                        let raw = std::mem::take(s);
                        self.pos += 1;
                        parse_field_selector_from_brackets(&raw, dot_span)?
                    }
                    TokenKind::String(s) => {
                        let value = std::mem::take(s);
                        self.pos += 1;
                        value
                    }
                    TokenKind::RBracket => String::new(),
                    _ => {
                        return Err(ParseError::new("Expected field selector", self.current_span()));
                    }
                };
                self.expect(TokenKind::RBracket)?;
                Ok(Expr::FieldAccess(FieldAccessExpr {
                    base: Box::new(base),
                    field,
                }))
            }
            _ => Err(ParseError::new("Expected field selector", dot_span)),
        }
    }

    fn parse_field_access_best_effort(&mut self, base: Expr) -> Expr {
        let dot_span = self.current_span();
        if let Err(e) = self.expect(TokenKind::Dot) {
            self.record_error(e);
            return base;
        }

        self.skip_trivia();
        match self.peek_kind() {
            TokenKind::Ident(_) => {
                let name = self.take_ident_token_unchecked();

                if name.is_empty() {
                    self.record_error(ParseError::new("Expected field name", dot_span));
                    return base;
                }

                if !name.contains('.') {
                    return Expr::FieldAccess(FieldAccessExpr {
                        base: Box::new(base),
                        field: name,
                    });
                }

                let mut expr = base;
                for part in name.split('.') {
                    if part.is_empty() {
                        self.record_error(ParseError::new("Expected field name", dot_span));
                        break;
                    }
                    expr = Expr::FieldAccess(FieldAccessExpr {
                        base: Box::new(expr),
                        field: part.to_string(),
                    });
                }
                expr
            }
            TokenKind::LBracket => {
                self.next(); // '['

                let field = match &mut self.tokens[self.pos].kind {
                    TokenKind::Ident(s) => {
                        let raw = std::mem::take(s);
                        self.pos += 1;
                        match parse_field_selector_from_brackets(&raw, dot_span) {
                            Ok(f) => f,
                            Err(e) => {
                                self.record_error(e);
                                raw.trim().to_string()
                            }
                        }
                    }
                    TokenKind::String(s) => {
                        let value = std::mem::take(s);
                        self.pos += 1;
                        value
                    }
                    _ => String::new(),
                };

                self.skip_trivia();
                if matches!(self.peek_kind(), TokenKind::RBracket) {
                    self.next();
                } else {
                    self.record_error(ParseError::new(
                        "Unterminated field selector",
                        self.current_span(),
                    ));
                    // Attempt to resync by consuming until ']' or EOF.
                    while !matches!(self.peek_kind(), TokenKind::RBracket | TokenKind::Eof) {
                        self.next();
                    }
                    if matches!(self.peek_kind(), TokenKind::RBracket) {
                        self.next();
                    }
                }

                Expr::FieldAccess(FieldAccessExpr {
                    base: Box::new(base),
                    field,
                })
            }
            _ => {
                self.record_error(ParseError::new("Expected field selector", dot_span));
                Expr::FieldAccess(FieldAccessExpr {
                    base: Box::new(base),
                    field: String::new(),
                })
            }
        }
    }

    fn parse_array_literal(&mut self) -> Result<Expr, ParseError> {
        if self.array_depth >= EXCEL_MAX_NESTED_CALLS {
            return Err(ParseError::new(
                format!("Expression nesting exceeds Excel's {EXCEL_MAX_NESTED_CALLS}-level limit"),
                self.current_span(),
            ));
        }

        self.expect(TokenKind::LBrace)?;
        self.array_depth += 1;
        let result = (|| {
            let mut rows: Vec<Vec<Expr>> = Vec::new();
            let mut current_row: Vec<Expr> = Vec::new();
            let mut expecting_value = true;
            loop {
                self.skip_trivia();
                match self.peek_kind() {
                    TokenKind::RBrace => {
                        self.next();
                        if expecting_value && (!current_row.is_empty() || !rows.is_empty()) {
                            current_row.push(Expr::Missing);
                        }
                        if !current_row.is_empty() || !rows.is_empty() {
                            rows.push(current_row);
                        }
                        break;
                    }
                    TokenKind::ArrayColSep => {
                        // Blank element, e.g. `{1,,3}`.
                        current_row.push(Expr::Missing);
                        self.next();
                        expecting_value = true;
                        continue;
                    }
                    TokenKind::ArrayRowSep => {
                        // Blank element at the end of a row, e.g. `{1,;2,3}`.
                        current_row.push(Expr::Missing);
                        self.next();
                        rows.push(current_row);
                        current_row = Vec::new();
                        expecting_value = true;
                        continue;
                    }
                    _ => {}
                }

                let el = self.parse_expression(0)?;
                current_row.push(el);
                expecting_value = false;
                self.skip_trivia();
                match self.peek_kind() {
                    TokenKind::ArrayColSep => {
                        self.next();
                        expecting_value = true;
                    }
                    TokenKind::ArrayRowSep => {
                        self.next();
                        rows.push(current_row);
                        current_row = Vec::new();
                        expecting_value = true;
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
        })();
        self.array_depth = self.array_depth.saturating_sub(1);
        result
    }

    fn parse_bracket_start(&mut self) -> Result<Expr, ParseError> {
        // Could be an external workbook prefix ([Book]Sheet!A1) or a structured ref like [@Col].
        // Look ahead for pattern: [ ... ] <sheet> !
        let save = self.pos;
        let open_span = self.current_span();
        self.expect(TokenKind::LBracket)?;
        let after_open = self.pos;

        // Workbook ids may include `]` (e.g. `C:\[foo]\Book.xlsx`), so the first `]` token is not
        // necessarily the workbook delimiter. Instead, treat any `]` as a *candidate* delimiter
        // and pick the one that yields a valid `[workbook]sheet!` prefix.
        #[derive(Debug, Clone, Copy, PartialEq, Eq)]
        enum ExternalPrefixKind {
            SheetBang,
            SheetSpanBang,
            WorkbookNameOrStructured,
        }

        let mut chosen: Option<(usize, ExternalPrefixKind)> = None;
        {
            let tokens = &self.tokens;
            let kind_at = |idx: usize| tokens.get(idx).map(|t| &t.kind);
            let skip_ws = |mut idx: usize| {
                while matches!(kind_at(idx), Some(TokenKind::Whitespace(_))) {
                    idx += 1;
                }
                idx
            };

            for close_idx in after_open..tokens.len() {
                match kind_at(close_idx) {
                    Some(TokenKind::RBracket) => {}
                    Some(TokenKind::Eof) | None => break,
                    _ => continue,
                }

                let close_span = tokens[close_idx].span;
                let workbook_start = open_span.end;
                let workbook_end = close_span.start;
                if workbook_end <= workbook_start {
                    continue;
                }

                let start_idx = skip_ws(close_idx + 1);
                // Only treat the token after `]` as a candidate sheet/name token if it looks like
                // something the lexer would have produced in normal mode (i.e. not raw bracket
                // content like `:` that gets emitted as `Ident` while inside `[...]`).
                let valid_name_token = match kind_at(start_idx) {
                    Some(TokenKind::QuotedIdent(_)) => true,
                    Some(TokenKind::Ident(s)) => s
                        .chars()
                        .next()
                        .is_some_and(|c| is_ident_start_char(c)),
                    _ => false,
                };
                if !valid_name_token {
                    continue;
                }
                let idx = skip_ws(start_idx + 1);

                match kind_at(idx) {
                    Some(TokenKind::Bang) => {
                        chosen = Some((close_idx, ExternalPrefixKind::SheetBang));
                        break;
                    }
                    Some(TokenKind::Colon) => {
                        let end_idx = skip_ws(idx + 1);
                        if !matches!(
                            kind_at(end_idx),
                            Some(TokenKind::Ident(_) | TokenKind::QuotedIdent(_))
                        ) {
                            continue;
                        }
                        let after_end = skip_ws(end_idx + 1);
                        if !matches!(kind_at(after_end), Some(TokenKind::Bang)) {
                            continue;
                        }
                        chosen = Some((close_idx, ExternalPrefixKind::SheetSpanBang));
                        break;
                    }
                    _ => {
                        // Reject candidates that are still inside a larger bracketed segment.
                        if matches!(kind_at(idx), Some(TokenKind::RBracket)) {
                            continue;
                        }
                        chosen = Some((close_idx, ExternalPrefixKind::WorkbookNameOrStructured));
                        break;
                    }
                }
            }
        }

        if let Some((close_idx, kind)) = chosen {
            let close_span = self.tokens[close_idx].span;
            let workbook_start = open_span.end;
            let workbook_end = close_span.start;
            let workbook = || self.src[workbook_start..workbook_end].to_string();

            self.pos = close_idx + 1;
            self.skip_trivia();
            let first = self.take_name_token_unchecked();
            self.skip_trivia();

            match kind {
                ExternalPrefixKind::SheetSpanBang => {
                    self.next(); // colon
                    self.skip_trivia();
                    let end = self.take_name_token_unchecked();
                    self.skip_trivia();
                    self.next(); // bang

                    let sheet_ref = if sheet_name_eq_case_insensitive(&first, &end) {
                        SheetRef::Sheet(first)
                    } else {
                        SheetRef::SheetRange { start: first, end }
                    };

                    return self.parse_ref_after_prefix(Some(workbook()), Some(sheet_ref));
                }
                ExternalPrefixKind::SheetBang => {
                    self.next(); // bang

                    let sheet_ref = match split_sheet_span_slices(&first) {
                        None => SheetRef::Sheet(first),
                        Some((start, _end)) => {
                            // `split_sheet_span_slices` returns slices into `first` of the form
                            // `{start}:{end}` (colon is ASCII and excluded from the slices). Reuse
                            // `first` for the start segment to avoid an extra allocation.
                            let start_len = start.len();
                            let mut start_owned = first;
                            let split_idx = start_len.saturating_add(1);
                            let end_owned = start_owned.split_off(split_idx);
                            start_owned.truncate(start_len);

                            if sheet_name_eq_case_insensitive(&start_owned, &end_owned) {
                                SheetRef::Sheet(start_owned)
                            } else {
                                SheetRef::SheetRange {
                                    start: start_owned,
                                    end: end_owned,
                                }
                            }
                        }
                    };

                    return self.parse_ref_after_prefix(Some(workbook()), Some(sheet_ref));
                }
                ExternalPrefixKind::WorkbookNameOrStructured => {
                    // If the token after the candidate name is another `]`, we're still inside a larger
                    // bracketed segment (meaning this `]` was not the workbook delimiter).
                    if matches!(self.peek_kind(), TokenKind::RBracket) {
                        // Fall through to structured-ref parsing below.
                    } else {
                        let workbook = workbook();
                        if matches!(self.peek_kind(), TokenKind::LBracket) {
                            return self.parse_structured_ref(Some(workbook), None, Some(first));
                        }
                        return Ok(Expr::NameRef(NameRef {
                            workbook: Some(workbook),
                            sheet: None,
                            name: first,
                        }));
                    }
                }
            }
        }

        // Not an external ref; rewind and parse as structured.
        self.pos = save;
        self.parse_structured_ref(None, None, None)
    }

    fn parse_structured_ref(
        &mut self,
        workbook: Option<String>,
        sheet: Option<SheetRef>,
        table: Option<String>,
    ) -> Result<Expr, ParseError> {
        self.skip_trivia();
        let open_span = self.current_span();
        self.expect(TokenKind::LBracket)?;

        let spec_start = open_span.end;

        // Structured references can contain escaped `]` as `]]`, but `]]` is also used to close
        // nested bracket groups. Use the structured-ref parser's disambiguation logic to find the
        // correct end position when possible.
        if let Some((_, end_pos)) = crate::structured_refs::parse_structured_ref(self.src, open_span.start)
        {
            // Advance to the closing `]` token for the chosen end position.
            while self.current_span().end < end_pos {
                if matches!(self.peek_kind(), TokenKind::Eof) {
                    return Err(ParseError::new(
                        "Unterminated structured reference",
                        self.current_span(),
                    ));
                }
                self.next();
            }
            let close_span = self.current_span();
            self.expect(TokenKind::RBracket)?;

            let spec = self.src[spec_start..close_span.start].to_string();
            return Ok(Expr::StructuredRef(StructuredRef {
                workbook,
                sheet,
                table,
                spec,
            }));
        }

        let mut depth: i32 = 1;
        let mut spec_end: Option<usize> = None;

        while depth > 0 {
            match self.peek_kind() {
                TokenKind::LBracket => {
                    depth += 1;
                    self.next();
                }
                TokenKind::RBracket => {
                    // Excel escapes ']' inside structured references as ']]'. When parsing the
                    // *outermost* bracket, treat a double ']]' as a literal ']' rather than the
                    // end of the structured ref.
                    if depth == 1
                        && matches!(
                            self.tokens.get(self.pos + 1).map(|t| &t.kind),
                            Some(TokenKind::RBracket)
                        )
                    {
                        self.next();
                        self.next();
                        continue;
                    }

                    let close_span = self.current_span();
                    self.next();
                    depth -= 1;
                    if depth == 0 {
                        spec_end = Some(close_span.start);
                    }
                }
                TokenKind::Eof => {
                    return Err(ParseError::new(
                        "Unterminated structured reference",
                        self.current_span(),
                    ));
                }
                _ => {
                    self.next();
                }
            }
        }

        let spec_end = spec_end.expect("loop should set spec_end when depth reaches zero");
        let spec = self.src[spec_start..spec_end].to_string();

        Ok(Expr::StructuredRef(StructuredRef {
            workbook,
            sheet,
            table,
            spec,
        }))
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

    fn take_ident_token_unchecked(&mut self) -> String {
        match &mut self.tokens[self.pos].kind {
            TokenKind::Ident(s) => {
                let out = std::mem::take(s);
                self.pos += 1;
                out
            }
            _ => unreachable!("caller should guard with TokenKind::Ident"),
        }
    }

    fn take_name_token_unchecked(&mut self) -> String {
        match &mut self.tokens[self.pos].kind {
            TokenKind::Ident(s) | TokenKind::QuotedIdent(s) => {
                let out = std::mem::take(s);
                self.pos += 1;
                out
            }
            _ => unreachable!("caller should guard with TokenKind::Ident | TokenKind::QuotedIdent"),
        }
    }

    fn take_number_token_unchecked(&mut self) -> String {
        match &mut self.tokens[self.pos].kind {
            TokenKind::Number(s) => {
                let out = std::mem::take(s);
                self.pos += 1;
                out
            }
            _ => unreachable!("caller should guard with TokenKind::Number"),
        }
    }

    fn take_string_token_unchecked(&mut self) -> String {
        match &mut self.tokens[self.pos].kind {
            TokenKind::String(s) => {
                let out = std::mem::take(s);
                self.pos += 1;
                out
            }
            _ => unreachable!("caller should guard with TokenKind::String"),
        }
    }

    fn take_error_token_unchecked(&mut self) -> String {
        match &mut self.tokens[self.pos].kind {
            TokenKind::Error(s) => {
                let out = std::mem::take(s);
                self.pos += 1;
                out
            }
            _ => unreachable!("caller should guard with TokenKind::Error"),
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

    fn looks_like_sheet_prefix(&self, start_pos: usize, allow_incomplete_span: bool) -> bool {
        let kind_at = |pos: usize| self.tokens.get(pos).map(|t| &t.kind);

        let mut pos = start_pos;
        while matches!(kind_at(pos), Some(TokenKind::Whitespace(_))) {
            pos += 1;
        }
        if !matches!(
            kind_at(pos),
            Some(TokenKind::Ident(_) | TokenKind::QuotedIdent(_))
        ) {
            return false;
        }
        pos += 1;
        while matches!(kind_at(pos), Some(TokenKind::Whitespace(_))) {
            pos += 1;
        }

        match kind_at(pos) {
            Some(TokenKind::Bang) => true,
            Some(TokenKind::Colon) => {
                pos += 1;
                while matches!(kind_at(pos), Some(TokenKind::Whitespace(_))) {
                    pos += 1;
                }
                if !matches!(
                    kind_at(pos),
                    Some(TokenKind::Ident(_) | TokenKind::QuotedIdent(_))
                ) {
                    return allow_incomplete_span;
                }
                pos += 1;
                while matches!(kind_at(pos), Some(TokenKind::Whitespace(_))) {
                    pos += 1;
                }
                matches!(kind_at(pos), Some(TokenKind::Bang))
            }
            _ => false,
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

fn parse_field_selector_from_brackets(raw_inner: &str, span: Span) -> Result<String, ParseError> {
    let trimmed = raw_inner.trim();
    if trimmed.starts_with('"') && trimmed.ends_with('"') {
        // Excel string literal escaping: `""` within the quoted string represents a literal `"`.
        return formula_model::unescape_excel_double_quoted_string_literal(trimmed)
            .ok_or_else(|| ParseError::new("Invalid string literal", span));
    }
    Ok(trimmed.to_string())
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

fn coerce_range_operands(left: Expr, right: Expr) -> (Expr, Expr) {
    enum ColCoerce {
        Existing,
        FromName { col: u32, abs: bool },
    }

    enum RowCoerce {
        Existing,
        FromName { row: u32, abs: bool },
        FromNumber { row: u32 },
    }

    fn col_coerce(expr: &Expr) -> Option<ColCoerce> {
        match expr {
            Expr::ColRef(_) => Some(ColCoerce::Existing),
            Expr::NameRef(n) => parse_col_ref_name(&n.name).map(|(col, abs)| ColCoerce::FromName {
                col,
                abs,
            }),
            _ => None,
        }
    }

    fn row_coerce(expr: &Expr) -> Option<RowCoerce> {
        match expr {
            Expr::RowRef(_) => Some(RowCoerce::Existing),
            Expr::NameRef(n) => parse_row_ref_name(&n.name).map(|(row, abs)| RowCoerce::FromName {
                row,
                abs,
            }),
            Expr::Number(raw) => parse_row_number_literal(raw).map(|row| RowCoerce::FromNumber { row }),
            _ => None,
        }
    }

    fn into_col_ref(expr: Expr, coerce: ColCoerce) -> ColRef {
        match (expr, coerce) {
            (Expr::ColRef(r), ColCoerce::Existing) => r,
            (Expr::NameRef(n), ColCoerce::FromName { col, abs }) => ColRef {
                workbook: n.workbook,
                sheet: n.sheet,
                col: Coord::A1 { index: col, abs },
            },
            _ => unreachable!("col_coerce should be checked before calling into_col_ref"),
        }
    }

    fn into_row_ref(expr: Expr, coerce: RowCoerce) -> RowRef {
        match (expr, coerce) {
            (Expr::RowRef(r), RowCoerce::Existing) => r,
            (Expr::NameRef(n), RowCoerce::FromName { row, abs }) => RowRef {
                workbook: n.workbook,
                sheet: n.sheet,
                row: Coord::A1 { index: row, abs },
            },
            (Expr::Number(_), RowCoerce::FromNumber { row }) => RowRef {
                workbook: None,
                sheet: None,
                row: Coord::A1 {
                    index: row,
                    abs: false,
                },
            },
            _ => unreachable!("row_coerce should be checked before calling into_row_ref"),
        }
    }

    if let (Some(left_coerce), Some(right_coerce)) = (col_coerce(&left), col_coerce(&right)) {
        return (
            Expr::ColRef(into_col_ref(left, left_coerce)),
            Expr::ColRef(into_col_ref(right, right_coerce)),
        );
    }

    if let (Some(left_coerce), Some(right_coerce)) = (row_coerce(&left), row_coerce(&right)) {
        return (
            Expr::RowRef(into_row_ref(left, left_coerce)),
            Expr::RowRef(into_row_ref(right, right_coerce)),
        );
    }

    (left, right)
}

fn parse_col_ref_name(raw: &str) -> Option<(u32, bool)> {
    let (abs, letters) = raw
        .strip_prefix('$')
        .map(|rest| (true, rest))
        .unwrap_or((false, raw));
    let col = column_label_to_index_lenient(letters).ok()?;
    Some((col, abs))
}

fn parse_row_ref_name(raw: &str) -> Option<(u32, bool)> {
    let (abs, digits) = raw
        .strip_prefix('$')
        .map(|rest| (true, rest))
        .unwrap_or((false, raw));
    if digits.is_empty() || !digits.chars().all(|c| c.is_ascii_digit()) {
        return None;
    }
    let row: u32 = digits.parse().ok()?;
    if row == 0 {
        return None;
    }
    Some((row - 1, abs))
}

fn parse_row_number_literal(raw: &str) -> Option<u32> {
    let row: u32 = raw.parse().ok()?;
    if row == 0 {
        return None;
    }
    Some(row - 1)
}

fn split_external_sheet_name_parts(name: &str) -> (Option<Cow<'_, str>>, Cow<'_, str>) {
    // Canonical external sheet keys are encoded as `"[{workbook}]{sheet}"`. Workbook ids can
    // include path prefixes from quoted external references (e.g. `'C:\\[foo]\\[Book.xlsx]Sheet1'!A1`)
    // and those prefixes may themselves contain `[` / `]`. To avoid ambiguity, we split workbook
    // ids on the **last** `]` (matching `eval::split_external_sheet_key_parts`).
    //
    // Note: Excel's raw formula syntax escapes literal `]` characters inside workbook names by
    // doubling them (`]]`). The canonical form preserves those characters, so `]]` may appear in
    // the workbook id. We treat it as plain text here.
    //
    // Sheet references can also be path-qualified inside a quoted sheet identifier, e.g.
    // `'C:\path\[Book.xlsx]Sheet1'!A1`.
    //
    // In these cases the raw quoted identifier does not start with `[`, but still contains a
    // `[workbook]sheet` segment. Canonicalize these by folding the path prefix into the workbook
    // id so external sheet keys remain unique:
    // `C:\path\[Book.xlsx]Sheet1` -> workbook `C:\path\Book.xlsx`, sheet `Sheet1`.
    if name.starts_with('[') {
        let Some((workbook, sheet)) = crate::external_refs::split_external_sheet_key_parts(name)
        else {
            return (None, Cow::Borrowed(name));
        };
        return (Some(Cow::Borrowed(workbook)), Cow::Borrowed(sheet));
    }

    // Path-qualified external workbook sheet refs can be lexed as a single quoted sheet name, e.g.
    // `'C:\path\[Book.xlsx]Sheet1'!A1`. Find the bracketed workbook segment and fold any leading
    // path prefix into the workbook id so the resulting key is unambiguous.
    let Some((workbook, sheet)) =
        formula_model::external_refs::parse_path_qualified_external_sheet_key(name)
    else {
        return (None, Cow::Borrowed(name));
    };
    (Some(Cow::Owned(workbook)), Cow::Owned(sheet))
}

fn split_external_sheet_name(name: &str) -> (Option<String>, String) {
    let (workbook, sheet) = split_external_sheet_name_parts(name);
    (workbook.map(Cow::into_owned), sheet.into_owned())
}

fn sheet_ref_from_raw_prefix(raw: &str) -> (Option<String>, SheetRef) {
    let (workbook, sheet) = split_external_sheet_name(raw);
    let sheet_ref = match split_sheet_span_slices(&sheet) {
        Some((start, end)) if sheet_name_eq_case_insensitive(start, end) => {
            SheetRef::Sheet(start.to_string())
        }
        Some((start, end)) => SheetRef::SheetRange {
            start: start.to_string(),
            end: end.to_string(),
        },
        None => SheetRef::Sheet(sheet),
    };
    (workbook, sheet_ref)
}

fn split_sheet_span_slices(name: &str) -> Option<(&str, &str)> {
    let (start, end) = name.split_once(':')?;
    if start.is_empty() || end.is_empty() {
        return None;
    }
    Some((start, end))
}

fn estimate_tokenized_bytes(expr: &Expr) -> usize {
    // Approximate the size of Excel's internal token stream (ptg) for a parsed AST.
    //
    // This is intentionally not a perfect model of Excel's serialized formula format. It is a
    // stable, deterministic estimate used only for enforcing the 16,384-byte limit and protecting
    // against pathological formulas.
    //
    // Note: this function is written iteratively (rather than recursively) to avoid overflowing
    // the Rust stack on formulas that produce deep left-associated ASTs near Excel's size limits.
    let mut total = 0usize;
    let mut stack: Vec<&Expr> = vec![expr];

    while let Some(node) = stack.pop() {
        match node {
            Expr::Number(raw) => total = total.saturating_add(estimate_number_token_bytes(raw)),
            Expr::String(s) => {
                // BIFF8 `ptgStr` is: token byte + cch (1) + flags (1) + character data.
                // Excel stores strings in a compressed/uncompressed unicode form; we conservatively
                // assume 2 bytes per character.
                total = total
                    .saturating_add(3usize.saturating_add(s.chars().count().saturating_mul(2)));
            }
            Expr::Boolean(_) => total = total.saturating_add(2),
            Expr::Error(_) => total = total.saturating_add(2),
            Expr::NameRef(_) => total = total.saturating_add(5),
            Expr::CellRef(r) => {
                // 3D/external refs are larger than local refs; approximate with a small bump.
                total = total.saturating_add(if r.workbook.is_some() || r.sheet.is_some() {
                    7
                } else {
                    5
                });
            }
            Expr::ColRef(r) => {
                // Full-column references are represented as areas.
                total = total.saturating_add(if r.workbook.is_some() || r.sheet.is_some() {
                    11
                } else {
                    9
                });
            }
            Expr::RowRef(r) => {
                // Full-row references are represented as areas.
                total = total.saturating_add(if r.workbook.is_some() || r.sheet.is_some() {
                    11
                } else {
                    9
                });
            }
            Expr::StructuredRef(_) => total = total.saturating_add(5),
            Expr::FieldAccess(access) => {
                // Field access isn't representable in Excel's legacy BIFF token stream, but we still
                // need a stable, conservative estimate to enforce the 16,384 byte limit.
                //
                // It is lowered into a synthetic `_FIELDACCESS(base, "field")` call. Approximate as
                // a small operator/call overhead plus a string-like payload for the field name, and
                // include the base expression size.
                total = total.saturating_add(4);
                total = total.saturating_add(
                    3usize.saturating_add(access.field.chars().count().saturating_mul(2)),
                );
                stack.push(access.base.as_ref());
            }
            Expr::Array(arr) => {
                // Arrays carry inline data; approximate by summing element sizes plus a small header.
                total = total.saturating_add(4);
                for el in arr.rows.iter().flatten() {
                    stack.push(el);
                }
            }
            Expr::FunctionCall(call) => {
                // `ptgFuncVar`: token + argc + func id
                total = total.saturating_add(4);
                for arg in &call.args {
                    stack.push(arg);
                }
            }
            Expr::Call(call) => {
                // Treat anonymous/lambda calls similarly to function calls.
                total = total.saturating_add(4);
                stack.push(call.callee.as_ref());
                for arg in &call.args {
                    stack.push(arg);
                }
            }
            Expr::Unary(u) => {
                total = total.saturating_add(1);
                stack.push(u.expr.as_ref());
            }
            Expr::Postfix(p) => {
                total = total.saturating_add(1);
                stack.push(p.expr.as_ref());
            }
            Expr::Binary(b) => {
                // `:` ranges with static operands can be represented as a single area token.
                if b.op == BinaryOp::Range && can_collapse_range_operands(&b.left, &b.right) {
                    // Area size depends on whether the reference is 3D/external.
                    let has_sheet = match (b.left.as_ref(), b.right.as_ref()) {
                        (Expr::CellRef(l), Expr::CellRef(r)) => {
                            l.workbook.is_some()
                                || l.sheet.is_some()
                                || r.workbook.is_some()
                                || r.sheet.is_some()
                        }
                        (Expr::ColRef(l), Expr::ColRef(r)) => {
                            l.workbook.is_some()
                                || l.sheet.is_some()
                                || r.workbook.is_some()
                                || r.sheet.is_some()
                        }
                        (Expr::RowRef(l), Expr::RowRef(r)) => {
                            l.workbook.is_some()
                                || l.sheet.is_some()
                                || r.workbook.is_some()
                                || r.sheet.is_some()
                        }
                        _ => false,
                    };
                    total = total.saturating_add(if has_sheet { 11 } else { 9 });
                } else {
                    total = total.saturating_add(1);
                    stack.push(b.left.as_ref());
                    stack.push(b.right.as_ref());
                }
            }
            Expr::Missing => total = total.saturating_add(1),
        }
    }

    total
}

fn can_collapse_range_operands(left: &Expr, right: &Expr) -> bool {
    // Only collapse simple reference spans (e.g. `A1:B2`, `A:A`, `1:1`). More complex ranges
    // like `OFFSET(...):A1` must remain as operand + operand + range operator.
    matches!(
        (left, right),
        (Expr::CellRef(_), Expr::CellRef(_))
            | (Expr::ColRef(_), Expr::ColRef(_))
            | (Expr::RowRef(_), Expr::RowRef(_))
    )
}

fn estimate_number_token_bytes(raw: &str) -> usize {
    // Excel may store small integer literals (0..=65535) as `ptgInt` (3 bytes) instead of
    // `ptgNum` (9 bytes). This improves the fidelity of our token-size estimate for formulas that
    // are dense with numeric literals.
    //
    // We only treat the literal as an integer if it consists solely of ASCII digits (no decimal
    // point/exponent), since unary `-` is tokenized separately.
    if raw.as_bytes().iter().all(|b| matches!(b, b'0'..=b'9')) {
        if let Ok(v) = raw.parse::<u32>() {
            if v <= u16::MAX as u32 {
                return 3;
            }
        }
    }
    9
}

#[cfg(test)]
mod tests {
    use super::*;
    use crate::{CellRef, Coord, Expr, FunctionCall, ParseOptions, SerializeOptions, SheetRef};

    #[test]
    fn true_false_lex_as_boolean_literals_when_not_followed_by_paren() {
        let opts = ParseOptions::default();
        let tokens = lex("TRUE", &opts).unwrap();
        let kinds: Vec<TokenKind> = tokens.into_iter().map(|t| t.kind).collect();
        assert_eq!(kinds, vec![TokenKind::Boolean(true), TokenKind::Eof]);

        let tokens = lex("FALSE", &opts).unwrap();
        let kinds: Vec<TokenKind> = tokens.into_iter().map(|t| t.kind).collect();
        assert_eq!(kinds, vec![TokenKind::Boolean(false), TokenKind::Eof]);
    }

    #[test]
    fn true_false_lex_as_idents_when_called_with_parentheses() {
        let opts = ParseOptions::default();

        let tokens = lex("TRUE()", &opts).unwrap();
        let kinds: Vec<TokenKind> = tokens.into_iter().map(|t| t.kind).collect();
        assert_eq!(
            kinds,
            vec![
                TokenKind::Ident("TRUE".to_string()),
                TokenKind::LParen,
                TokenKind::RParen,
                TokenKind::Eof
            ]
        );

        // Whitespace between the name and `(` still counts as a call.
        let tokens = lex("FALSE \t()", &opts).unwrap();
        let kinds: Vec<TokenKind> = tokens.into_iter().map(|t| t.kind).collect();
        assert_eq!(
            kinds,
            vec![
                TokenKind::Ident("FALSE".to_string()),
                TokenKind::Whitespace(" \t".to_string()),
                TokenKind::LParen,
                TokenKind::RParen,
                TokenKind::Eof
            ]
        );
    }

    #[test]
    fn true_false_paren_forms_parse_as_function_calls_not_postfix_calls() {
        let ast = parse_formula("=TRUE()", ParseOptions::default()).unwrap();
        match ast.expr {
            Expr::FunctionCall(FunctionCall { name, args }) => {
                assert_eq!(name.name_upper, "TRUE");
                assert!(args.is_empty());
            }
            other => panic!("expected FunctionCall, got {other:?}"),
        }

        let ast = parse_formula("=FALSE()", ParseOptions::default()).unwrap();
        match ast.expr {
            Expr::FunctionCall(FunctionCall { name, args }) => {
                assert_eq!(name.name_upper, "FALSE");
                assert!(args.is_empty());
            }
            other => panic!("expected FunctionCall, got {other:?}"),
        }
    }

    #[test]
    fn r1c1_cell_ref_followed_by_dot_lexes_as_cell_and_field_access() {
        // Regression test: allow `.` after an R1C1 reference so expressions like `RC[-1].Price`
        // are lexed as a reference token followed by field access, rather than a name-like token.
        let opts = ParseOptions {
            reference_style: ReferenceStyle::R1C1,
            ..ParseOptions::default()
        };

        let tokens = lex("RC[-1].Price", &opts).unwrap();
        let kinds: Vec<TokenKind> = tokens.into_iter().map(|t| t.kind).collect();
        assert_eq!(
            kinds,
            vec![
                TokenKind::R1C1Cell(R1C1CellToken {
                    row: Coord::Offset(0),
                    col: Coord::Offset(-1)
                }),
                TokenKind::Dot,
                TokenKind::Ident("Price".to_string()),
                TokenKind::Eof
            ]
        );
    }

    #[test]
    fn a1_cell_ref_followed_by_ident_char_lexes_as_single_ident() {
        // Regression test: avoid tokenizing `A1FOO` as `A1` + `FOO`, since Excel allows defined
        // names like `A1FOO` (it is not a complete A1 reference).
        let opts = ParseOptions::default();
        let tokens = lex("A1FOO", &opts).unwrap();
        let kinds: Vec<TokenKind> = tokens.into_iter().map(|t| t.kind).collect();
        assert_eq!(kinds, vec![TokenKind::Ident("A1FOO".to_string()), TokenKind::Eof]);
    }

    #[test]
    fn out_of_bounds_a1_cell_ref_lexes_as_cell_token() {
        // Out-of-bounds column labels should still be lexed as cell references so they can
        // later evaluate to `#REF!` (rather than being treated as names).
        let opts = ParseOptions::default();
        let tokens = lex("XFE1", &opts).unwrap();
        let kinds: Vec<TokenKind> = tokens.into_iter().map(|t| t.kind).collect();
        assert_eq!(
            kinds,
            vec![
                TokenKind::Cell(CellToken {
                    col: 16_384,
                    row: 0,
                    col_abs: false,
                    row_abs: false
                }),
                TokenKind::Eof
            ]
        );
    }

    #[test]
    fn too_long_a1_column_label_lexes_as_ident() {
        // Excel A1 column labels are at most 3 letters. Longer "A1-looking" prefixes should be
        // treated as identifiers.
        let opts = ParseOptions::default();
        let tokens = lex("AAAA1", &opts).unwrap();
        let kinds: Vec<TokenKind> = tokens.into_iter().map(|t| t.kind).collect();
        assert_eq!(kinds, vec![TokenKind::Ident("AAAA1".to_string()), TokenKind::Eof]);
    }

    #[test]
    fn unicode_sheet_span_collapses_with_excel_like_case_insensitive_matching() {
        // German sharp s: Unicode uppercasing expands `ß` -> `SS`.
        //
        // Excel compares sheet names case-insensitively across Unicode, so this 3D span should
        // collapse to a single-sheet reference.
        let formula = "='ß':'SS'!A1";
        let ast = parse_formula(formula, ParseOptions::default()).unwrap();

        // Parser normalization: `'<name>':'<casefold-equivalent>'!A1` should become a single-sheet ref.
        match &ast.expr {
            Expr::CellRef(r) => {
                assert_eq!(r.workbook, None);
                assert_eq!(
                    r.sheet,
                    Some(SheetRef::Sheet("ß".to_string())),
                    "expected sheet span to collapse during parsing"
                );
            }
            other => panic!("expected CellRef, got {other:?}"),
        }

        // Stringification should not reintroduce the 3D `start:end` form.
        let rendered = ast.to_string(SerializeOptions::default()).unwrap();
        assert_eq!(rendered, "='ß'!A1");

        // Compiler normalization: even if a SheetRef::SheetRange reaches compilation, it should
        // collapse to a single sheet id using Unicode-aware matching.
        let range_expr = Expr::CellRef(CellRef {
            workbook: None,
            sheet: Some(SheetRef::SheetRange {
                start: "ß".to_string(),
                end: "SS".to_string(),
            }),
            col: Coord::A1 {
                index: 0,
                abs: false,
            },
            row: Coord::A1 {
                index: 0,
                abs: false,
            },
        });

        let ast_range = crate::Ast::new(true, range_expr.clone());
        let rendered_range = ast_range.to_string(SerializeOptions::default()).unwrap();
        assert_eq!(
            rendered_range, "='ß'!A1",
            "expected sheet span to collapse during serialization"
        );

        let mut resolve_sheet = |name: &str| {
            formula_model::sheet_name_eq_case_insensitive(name, "ß")
                .then_some(0usize)
        };
        let mut sheet_dims =
            |_id: usize| (formula_model::EXCEL_MAX_ROWS, formula_model::EXCEL_MAX_COLS);

        let compiled = crate::eval::compile_canonical_expr(
            &range_expr,
            0,
            crate::eval::CellAddr { row: 0, col: 0 },
            &mut resolve_sheet,
            &mut sheet_dims,
        );
        match compiled {
            crate::eval::Expr::CellRef(r) => {
                assert_eq!(
                    r.sheet,
                    crate::eval::SheetReference::Sheet(0),
                    "expected sheet span to collapse during compilation"
                );
            }
            other => panic!("expected compiled CellRef, got {other:?}"),
        }
    }

    #[test]
    fn quoted_external_workbook_name_ref_splits_workbook_prefix() {
        let ast = parse_formula("='[Book.xlsx]MyName'", ParseOptions::default()).unwrap();
        match &ast.expr {
            Expr::NameRef(r) => {
                assert_eq!(r.workbook.as_deref(), Some("Book.xlsx"));
                assert_eq!(r.sheet, None);
                assert_eq!(r.name, "MyName");
            }
            other => panic!("expected NameRef, got {other:?}"),
        }

        // Add-ins emit non-function extern names (constants / macros) via NameX with a workbook-ish
        // prefix like `'[AddIn]ConstName'`.
        let ast = parse_formula("='[AddIn]MyAddinConst'", ParseOptions::default()).unwrap();
        match &ast.expr {
            Expr::NameRef(r) => {
                assert_eq!(r.workbook.as_deref(), Some("AddIn"));
                assert_eq!(r.sheet, None);
                assert_eq!(r.name, "MyAddinConst");
            }
            other => panic!("expected NameRef, got {other:?}"),
        }
    }

    #[test]
    fn partial_parse_splits_external_workbook_prefix_in_quoted_name_refs() {
        let parsed = parse_formula_partial("='[Book.xlsx]MyName'", ParseOptions::default());
        assert!(
            parsed.error.is_none(),
            "unexpected parse error: {:?}",
            parsed.error
        );
        match &parsed.ast.expr {
            Expr::NameRef(r) => {
                assert_eq!(r.workbook.as_deref(), Some("Book.xlsx"));
                assert_eq!(r.sheet, None);
                assert_eq!(r.name, "MyName");
            }
            other => panic!("expected NameRef, got {other:?}"),
        }
    }

    #[test]
    fn unquoted_external_workbook_name_ref_parses_as_name_ref() {
        let ast = parse_formula("=[Book.xlsx]MyName", ParseOptions::default()).unwrap();
        match &ast.expr {
            Expr::NameRef(r) => {
                assert_eq!(r.workbook.as_deref(), Some("Book.xlsx"));
                assert_eq!(r.sheet, None);
                assert_eq!(r.name, "MyName");
            }
            other => panic!("expected NameRef, got {other:?}"),
        }

        // The serializer prefers the fully-quoted token form for workbook-scoped external names.
        let rendered = ast.to_string(SerializeOptions::default()).unwrap();
        assert_eq!(rendered, "='[Book.xlsx]MyName'");
    }

    #[test]
    fn unquoted_workbook_name_with_open_bracket_does_not_swallow_trailing_ops() {
        // Workbook ids can contain `[` characters (Excel treats them as plain text within the
        // `[workbook]` prefix). Ensure we don't treat this as a nested bracket expression and
        // accidentally swallow trailing operators like `+1` into a single identifier token.
        let ast = parse_formula("=[A1[Name.xlsx]MyName+1", ParseOptions::default()).unwrap();
        match &ast.expr {
            Expr::Binary(b) => {
                assert_eq!(b.op, BinaryOp::Add);
                assert_eq!(
                    b.left.as_ref(),
                    &Expr::NameRef(NameRef {
                        workbook: Some("A1[Name.xlsx".to_string()),
                        sheet: None,
                        name: "MyName".to_string(),
                    })
                );
                assert_eq!(b.right.as_ref(), &Expr::Number("1".to_string()));
            }
            other => panic!("expected Binary(Add), got {other:?}"),
        }

        let rendered = ast.to_string(SerializeOptions::default()).unwrap();
        assert_eq!(rendered, "='[A1[Name.xlsx]MyName'+1");
    }

    #[test]
    fn external_workbook_prefix_parses_when_workbook_contains_brackets() {
        // Regression test: workbook ids may include `]` (e.g. `C:\[foo]\Book.xlsx`), and our
        // serializer can emit bracketed workbook prefixes like `[C:\[foo]\Book.xlsx]Sheet1!A1`.
        //
        // Ensure the parser does not treat the `]` from `[foo]` as the workbook delimiter.
        let formula = r"=[C:\[foo]\Book.xlsx]Sheet1!A1";
        let ast = parse_formula(formula, ParseOptions::default()).unwrap();

        match &ast.expr {
            Expr::CellRef(r) => {
                assert_eq!(r.workbook.as_deref(), Some(r"C:\[foo]\Book.xlsx"));
                assert_eq!(
                    r.sheet,
                    Some(SheetRef::Sheet("Sheet1".to_string())),
                    "expected external workbook prefix to be parsed as a sheet reference"
                );
            }
            other => panic!("expected CellRef, got {other:?}"),
        }
    }

    #[test]
    fn external_workbook_prefix_parses_when_quoted_and_workbook_contains_brackets() {
        // When the sheet name needs quoting, Excel-style external references quote the combined
        // `[workbook]sheet` prefix. Ensure we still split on the *last* `]` when the workbook id
        // contains brackets.
        let formula = r"='[C:\[foo]\Book.xlsx]My Sheet'!A1";
        let ast = parse_formula(formula, ParseOptions::default()).unwrap();

        match &ast.expr {
            Expr::CellRef(r) => {
                assert_eq!(r.workbook.as_deref(), Some(r"C:\[foo]\Book.xlsx"));
                assert_eq!(
                    r.sheet,
                    Some(SheetRef::Sheet("My Sheet".to_string())),
                    "expected quoted external workbook prefix to be parsed correctly"
                );
            }
            other => panic!("expected CellRef, got {other:?}"),
        }
    }

    #[test]
    fn external_workbook_path_qualified_reference_round_trips_through_serializer() {
        // Excel can include a path prefix before the `[Book.xlsx]Sheet` portion of an external
        // reference. We canonicalize these into a single workbook id during parsing.
        let formula = r"='C:\[foo]\[Book.xlsx]Sheet1'!A1";
        let ast = parse_formula(formula, ParseOptions::default()).unwrap();

        match &ast.expr {
            Expr::CellRef(r) => {
                assert_eq!(r.workbook.as_deref(), Some(r"C:\[foo]\Book.xlsx"));
                assert_eq!(r.sheet, Some(SheetRef::Sheet("Sheet1".to_string())));
            }
            other => panic!("expected CellRef, got {other:?}"),
        }

        let rendered = ast.to_string(SerializeOptions::default()).unwrap();
        assert!(
            rendered.contains(r"[C:\[foo]\Book.xlsx]Sheet1")
                || rendered.contains(r"C:\[foo]\[Book.xlsx]Sheet1"),
            "expected serializer to emit a canonical external workbook prefix, got {rendered}"
        );

        // The canonical form should be parseable again.
        let reparsed = parse_formula(&rendered, ParseOptions::default()).unwrap();
        match &reparsed.expr {
            Expr::CellRef(r) => {
                assert_eq!(r.workbook.as_deref(), Some(r"C:\[foo]\Book.xlsx"));
                assert_eq!(r.sheet, Some(SheetRef::Sheet("Sheet1".to_string())));
            }
            other => panic!("expected CellRef, got {other:?}"),
        }
        assert_eq!(
            reparsed.to_string(SerializeOptions::default()).unwrap(),
            rendered
        );
    }

    #[test]
    fn external_workbook_prefix_parses_when_workbook_contains_open_bracket_after_bracketed_path() {
        // Regression test: workbook ids may contain bracketed path components *and* literal `[` in
        // the workbook name itself. Workbook prefixes are not nesting, so we should treat the
        // inner `[` as plain text and still locate the correct closing `]`.
        //
        // Example: a file name like `[Book.xlsx` in a folder `C:\[foo]\`.
        let formula = r"=[C:\[foo]\[Book.xlsx]Sheet1!A1";
        let ast = parse_formula(formula, ParseOptions::default()).unwrap();

        match &ast.expr {
            Expr::CellRef(r) => {
                assert_eq!(r.workbook.as_deref(), Some(r"C:\[foo]\[Book.xlsx"));
                assert_eq!(
                    r.sheet,
                    Some(SheetRef::Sheet("Sheet1".to_string())),
                    "expected external workbook prefix to be parsed as a sheet reference"
                );
            }
            other => panic!("expected CellRef, got {other:?}"),
        }
    }

    #[test]
    fn workbook_only_external_structured_ref_parses() {
        // External structured refs can be workbook-only (no explicit sheet), e.g.
        // `[Book.xlsx]Table1[Col]`.
        let formula = "=[Book.xlsx]Table1[Col]";
        let ast = parse_formula(formula, ParseOptions::default()).unwrap();

        match &ast.expr {
            Expr::StructuredRef(r) => {
                assert_eq!(r.workbook.as_deref(), Some("Book.xlsx"));
                assert_eq!(r.sheet, None);
                assert_eq!(r.table.as_deref(), Some("Table1"));
                assert_eq!(r.spec, "Col");
            }
            other => panic!("expected StructuredRef, got {other:?}"),
        }
    }
}
