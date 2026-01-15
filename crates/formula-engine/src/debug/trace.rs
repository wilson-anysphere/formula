use crate::error::ExcelError;
use crate::eval::{
    parse_a1, CellAddr, CompareOp, EvalContext, FormulaParseError, SheetReference, UnaryOp,
};
use crate::functions::{ArgValue as FnArgValue, FunctionContext, SheetId as FnSheetId};
use crate::value::{Array, ErrorKind, Value};
use formula_model::formula_rewrite::sheet_name_eq_case_insensitive;
use std::cmp::Ordering;

/// Maximum number of cells the debug trace evaluator will materialize when a formula result is a
/// range reference.
///
/// Keep this aligned with `eval::Evaluator` to avoid OOMs for formulas like `=A1:XFD1048576`.
const MAX_REFERENCE_DEREF_CELLS: usize = 5_000_000;

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
    Cell {
        sheet: FnSheetId,
        addr: CellAddr,
    },
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
    StructuredRef,
    FieldAccess { field: String },
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
    StructuredRef(crate::eval::StructuredRefExpr<S>),
    NameRef(crate::eval::NameRef<S>),
    FieldAccess {
        base: Box<SpannedExpr<S>>,
        field: String,
    },
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
            SpannedExprKind::StructuredRef(r) => {
                SpannedExprKind::StructuredRef(crate::eval::StructuredRefExpr {
                    sheet: f(&r.sheet),
                    sref: r.sref.clone(),
                })
            }
            SpannedExprKind::NameRef(n) => SpannedExprKind::NameRef(crate::eval::NameRef {
                sheet: f(&n.sheet),
                name: n.name.clone(),
            }),
            SpannedExprKind::FieldAccess { base, field } => SpannedExprKind::FieldAccess {
                base: Box::new(base.map_sheets(f)),
                field: field.clone(),
            },
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
    recalc_ctx: &crate::eval::RecalcContext,
    date_system: crate::date::ExcelDateSystem,
    value_locale: crate::locale::ValueLocaleConfig,
    expr: &SpannedExpr<usize>,
) -> (Value, TraceNode) {
    let evaluator = TracedEvaluator {
        resolver,
        ctx,
        recalc_ctx,
        date_system,
        value_locale,
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
    StructuredRef(crate::structured_refs::StructuredRef),
    Error(ErrorKind),
    Dot,
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
                '.' => {
                    // Leading-decimal numeric literal (e.g. `.5`) vs field-access operator.
                    let next_is_digit = self.input[self.pos + 1..]
                        .chars()
                        .next()
                        .is_some_and(|c| c.is_ascii_digit());
                    if next_is_digit {
                        self.lex_number()?
                    } else {
                        self.pos += 1;
                        TokenKind::Dot
                    }
                }
                '0'..='9' => self.lex_number()?,
                _ if is_ident_start(ch) => self
                    .try_lex_structured_ref()
                    .unwrap_or_else(|| self.lex_ident()),
                _ => {
                    return Err(FormulaParseError::UnexpectedToken(format!(
                        "unexpected character '{ch}'"
                    )))
                }
            };
            let span = Span::new(start, self.pos);
            let can_spill = matches!(
                &kind,
                TokenKind::Ident(_) | TokenKind::StructuredRef(_) | TokenKind::RParen
            );
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

    fn try_lex_structured_ref(&mut self) -> Option<TokenKind> {
        let start = self.pos;
        let ch = self.peek_char()?;

        // Don't interpret external workbook prefixes (`[Book.xlsx]Sheet1!A1`) as structured refs.
        if ch == '[' && self.looks_like_external_workbook_ref(start) {
            return None;
        }

        // Structured refs can be `[... ]` or `TableName[...]`. We only attempt parsing when the
        // current character could start such a reference to avoid needless work.
        if ch != '[' && !(ch.is_ascii_alphabetic() || ch == '_') {
            return None;
        }

        let (sref, end_pos) = crate::structured_refs::parse_structured_ref(self.input, start)?;
        if end_pos <= start {
            return None;
        }

        // Disambiguate between structured refs (`Table1[Col]`) and record field access using a
        // bracket selector (`A1.["Field"]`).
        //
        // The structured ref parser allows `.` in table names; however Excel's field-access syntax
        // uses the `.` immediately before a bracket selector. When the parsed table name ends with
        // a `.`, treat it as field access and let the identifier lexer split it into `Ident("A1.")`
        // + `StructuredRef(["Field"])`.
        if sref
            .table_name
            .as_ref()
            .is_some_and(|table_name| table_name.ends_with('.'))
        {
            return None;
        }
        self.pos = end_pos;
        Some(TokenKind::StructuredRef(sref))
    }

    fn looks_like_external_workbook_ref(&self, bracket_start: usize) -> bool {
        let bytes = self.input.as_bytes();
        if bytes.get(bracket_start) != Some(&b'[') {
            return false;
        }

        // Find the workbook closing bracket.
        //
        // External workbook prefixes escape literal `]` characters by doubling them: `]]` -> `]`.
        // Treat those escapes as part of the workbook segment (not as the terminator).
        let Some(close) =
            crate::external_refs::find_external_workbook_prefix_end(self.input, bracket_start)
        else {
            return false;
        };

        // Parse sheet name + optional sheet span + '!' after the workbook prefix.
        let mut pos = close;
        while let Some(ch) = self.input[pos..].chars().next() {
            if ch.is_whitespace() {
                pos += ch.len_utf8();
            } else {
                break;
            }
        }
        pos = match self.scan_sheet_name(pos) {
            Some(end) => end,
            None => return false,
        };
        while let Some(ch) = self.input[pos..].chars().next() {
            if ch.is_whitespace() {
                pos += ch.len_utf8();
            } else {
                break;
            }
        }

        if self.input[pos..].starts_with(':') {
            pos += 1;
            while let Some(ch) = self.input[pos..].chars().next() {
                if ch.is_whitespace() {
                    pos += ch.len_utf8();
                } else {
                    break;
                }
            }
            pos = match self.scan_sheet_name(pos) {
                Some(end) => end,
                None => return false,
            };
            while let Some(ch) = self.input[pos..].chars().next() {
                if ch.is_whitespace() {
                    pos += ch.len_utf8();
                } else {
                    break;
                }
            }
        }

        // Sheet reference like `[Book.xlsx]Sheet1!A1`.
        if self.input[pos..].starts_with('!') {
            return true;
        }

        // Workbook-scoped external defined name like `[Book.xlsx]MyName`.
        //
        // The canonical parser treats `[workbook]name` as a name ref (and rewrites it to
        // `='[workbook]name'` on roundtrip) to avoid ambiguity with structured refs. Mirror that
        // behavior so we don't interpret the workbook prefix as a structured ref.
        true
    }

    fn scan_sheet_name(&self, start: usize) -> Option<usize> {
        if start >= self.input.len() {
            return None;
        }

        // Quoted sheet name: `'My Sheet'`
        if self.input[start..].starts_with('\'') {
            let mut pos = start + 1;
            loop {
                match self.input[pos..].chars().next() {
                    Some('\'') => {
                        if self.input[pos..].starts_with("''") {
                            pos += 2;
                            continue;
                        }
                        return Some(pos + 1);
                    }
                    Some(ch) => pos += ch.len_utf8(),
                    None => return None,
                }
            }
        }

        // Unquoted sheet name (identifier-like).
        let mut pos = start;
        let first = self.input[pos..].chars().next()?;
        if !is_ident_start(first) || first == '[' {
            return None;
        }
        pos += first.len_utf8();
        while let Some(ch) = self.input[pos..].chars().next() {
            if ch.is_ascii_alphanumeric() || matches!(ch, '_' | '.' | '$') {
                pos += ch.len_utf8();
            } else {
                break;
            }
        }
        Some(pos)
    }

    fn lex_ident(&mut self) -> TokenKind {
        let start = self.pos;

        // External workbook prefixes (`[Book.xlsx]Sheet1!A1`) are treated as a single identifier
        // token, and the workbook portion inside `[...]` is more permissive than a normal Excel
        // identifier (it may contain spaces, dashes, etc). Mirror the canonical lexer by consuming
        // everything up to the closing `]` before switching back to strict identifier rules for
        // the sheet name portion.
        if self.peek_char() == Some('[') {
            if let Some(end) =
                crate::external_refs::find_external_workbook_prefix_end(self.input, self.pos)
            {
                self.pos = end;
            } else {
                // No closing bracket; treat the rest of the input as part of this identifier.
                self.pos = self.input.len();
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
        debug_assert_eq!(self.peek_char(), Some('#'));
        self.pos += 1; // '#'
        while let Some(ch) = self.peek_char() {
            // Keep the accepted body characters in sync with the canonical lexer:
            // `parser/mod.rs:is_error_body_char`.
            if ch.is_ascii_alphanumeric() || matches!(ch, '_' | '/' | '.') {
                self.pos += ch.len_utf8();
            } else {
                break;
            }
        }
        if matches!(self.peek_char(), Some('!' | '?')) {
            self.pos += 1;
        }
        let s = &self.input[start..self.pos];
        let kind = ErrorKind::from_code(s).unwrap_or(ErrorKind::Value);
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
    // When an external workbook key contains an absolute/relative path, the workbook id can
    // contain `:` (e.g. a Windows drive letter: `[C:\path\Book.xlsx]Sheet1`).
    //
    // Only treat `:` as a 3D sheet-span separator when it appears in the sheet portion after the
    // closing `]`, not inside the workbook id.
    let (start, end) = if let Some(rest) = name.strip_prefix('[') {
        // The workbook id can contain `[` / `]` (e.g. `C:\[foo]\Book.xlsx`), so locate the final
        // closing bracket which terminates the workbook segment.
        let close_rel = rest.rfind(']')?;
        let close = 1 + close_rel;
        let sheet_part = &name[close + 1..];
        let colon_rel = sheet_part.find(':')?;
        let colon = close + 1 + colon_rel;
        (&name[..colon], &name[colon + 1..])
    } else {
        name.split_once(':')?
    };
    if start.is_empty() || end.is_empty() {
        return None;
    }
    Some((start.to_string(), end.to_string()))
}

fn parse_workbook_scoped_external_name_ref(name: &str) -> Option<(String, String)> {
    if !name.starts_with('[') {
        return None;
    }

    let (prefix, remainder) = formula_model::external_refs::split_external_workbook_prefix(name)?;
    let workbook = prefix.strip_prefix('[')?.strip_suffix(']')?;
    if workbook.is_empty() || remainder.is_empty() {
        return None;
    }

    Some((workbook.to_string(), remainder.to_string()))
}

/// Parse an Excel-style path-qualified external workbook prefix that's been lexed as a single
/// quoted sheet name, e.g. `C:\path\[Book.xlsx]Sheet1`.
///
/// Returns `(workbook, sheet)` where `workbook` is the canonical workbook name/path (without
/// brackets) and `sheet` is the sheet portion after the closing `]`.
///
/// This is a best-effort parser used only by the debug trace subsystem.
fn parse_path_qualified_external_sheet_key(name: &str) -> Option<(String, String)> {
    // We only treat this as the path-qualified form when `[` is not the leading character. If the
    // name starts with `[`, it's already in the canonical external key form (or is malformed).
    if name.starts_with('[') {
        return None;
    }

    // Excel external workbook prefixes:
    // - are not nested (workbook names may contain `[` characters)
    // - escape literal `]` characters by doubling them (`]]`)
    //
    // Path-qualified references can also contain bracketed path components, e.g.
    // `C:\[foo]\[Book.xlsx]Sheet1`. Scan for the best `[workbook]sheet` segment by finding the
    // bracketed component whose closing `]` is furthest to the right, skipping over workbook
    // prefixes so we do not misclassify `[` characters inside workbook names.
    let bytes = name.as_bytes();
    let mut i = 0usize;
    let mut best: Option<(usize, usize)> = None; // (open, end) where end is exclusive of the closing `]`

    while i < bytes.len() {
        if bytes[i] == b'[' {
            if let Some(end) = crate::external_refs::find_external_workbook_prefix_end(name, i) {
                // Only treat this as a workbook prefix if there is a remainder (sheet name) after
                // the closing `]`.
                if end < name.len() {
                    best = match best {
                        None => Some((i, end)),
                        Some((best_start, best_end)) => {
                            if end > best_end {
                                Some((i, end))
                            } else if end == best_end && i < best_start {
                                Some((i, end))
                            } else {
                                Some((best_start, best_end))
                            }
                        }
                    };
                }

                // Skip the entire bracketed segment to avoid misclassifying `[` characters inside
                // the workbook name as the start of a new prefix.
                i = end;
                continue;
            }
        }

        // Advance by UTF-8 char boundaries so we don't accidentally interpret `[` / `]` bytes
        // inside multi-byte sequences as actual bracket characters.
        let ch = name[i..].chars().next().expect("i always at char boundary");
        i += ch.len_utf8();
    }

    let Some((open, end)) = best else {
        return None;
    };

    // `end` is exclusive, so `end - 1` is the closing `]`.
    let book = &name[open + 1..end - 1];
    let sheet = &name[end..];
    if book.is_empty() || sheet.is_empty() {
        return None;
    }

    let prefix = &name[..open];
    let mut workbook = String::with_capacity(prefix.len().saturating_add(book.len()));
    workbook.push_str(prefix);
    workbook.push_str(book);
    Some((workbook, sheet.to_string()))
}

fn parse_bracket_quoted_field_name(raw: &str) -> Result<String, ()> {
    let raw = raw.trim();
    let Some(inner) = raw.strip_prefix('"').and_then(|s| s.strip_suffix('"')) else {
        return Err(());
    };

    // Excel string escaping uses doubled quotes: `""` -> `"`.
    let mut out = String::with_capacity(inner.len());
    let mut chars = inner.chars().peekable();
    while let Some(ch) = chars.next() {
        if ch == '"' && chars.peek() == Some(&'"') {
            chars.next();
            out.push('"');
        } else {
            out.push(ch);
        }
    }

    Ok(out)
}

fn parse_col_ref_str(input: &str) -> Option<u32> {
    let input = input.trim();
    let input = input.strip_prefix('$').unwrap_or(input);
    if input.is_empty() {
        return None;
    }

    let mut col: u32 = 0;
    for ch in input.chars() {
        if !ch.is_ascii_alphabetic() {
            return None;
        }
        let up = ch.to_ascii_uppercase();
        let digit = (up as u8 - b'A' + 1) as u32;
        col = col.checked_mul(26)?.checked_add(digit)?;
    }
    if col == 0 || col > formula_model::EXCEL_MAX_COLS {
        return None;
    }
    Some(col - 1)
}

fn parse_row_ref_str(input: &str) -> Option<u32> {
    let input = input.trim();
    let input = input.strip_prefix('$').unwrap_or(input);
    if input.is_empty() {
        return None;
    }

    let mut row: u32 = 0;
    for ch in input.chars() {
        if !ch.is_ascii_digit() {
            return None;
        }
        row = row.checked_mul(10)?.checked_add((ch as u8 - b'0') as u32)?;
    }
    if row == 0 {
        return None;
    }
    Some(row - 1)
}

fn parse_row_ref_number(n: f64) -> Option<u32> {
    if !n.is_finite() || n.fract() != 0.0 || n < 1.0 || n > (u32::MAX as f64) {
        return None;
    }
    let row = n as u32;
    if row == 0 {
        return None;
    }
    Some(row - 1)
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
                // Row range refs like `1:3` are lexed as number literals. Treat them as a
                // reference operator when they appear in a `:<row>` context; otherwise they
                // remain numeric constants.
                if matches!(self.peek_n(1).kind, TokenKind::Colon) {
                    let start_row = parse_row_ref_number(*n);
                    let end_row = match &self.peek_n(2).kind {
                        TokenKind::Number(m) => parse_row_ref_number(*m),
                        TokenKind::Ident(s) => parse_row_ref_str(s),
                        _ => None,
                    };
                    if let (Some(start_row), Some(end_row)) = (start_row, end_row) {
                        let start_tok = self.next();
                        self.next(); // ':'
                        let end_tok = self.next();
                        let span = Span::new(start_tok.span.start, end_tok.span.end);
                        let Some(start) = crate::eval::Ref::from_abs_cell_addr(CellAddr {
                            row: start_row,
                            col: 0,
                        }) else {
                            return Ok(SpannedExpr {
                                span,
                                kind: SpannedExprKind::Error(ErrorKind::Ref),
                            });
                        };
                        let Some(end) = crate::eval::Ref::from_abs_cell_addr(CellAddr {
                            row: end_row,
                            col: CellAddr::SHEET_END,
                        }) else {
                            return Ok(SpannedExpr {
                                span,
                                kind: SpannedExprKind::Error(ErrorKind::Ref),
                            });
                        };
                        return Ok(SpannedExpr {
                            span,
                            kind: SpannedExprKind::RangeRef(crate::eval::RangeRef {
                                sheet: SheetReference::Current,
                                start,
                                end,
                            }),
                        });
                    }
                }

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
            TokenKind::StructuredRef(r) => {
                self.next();
                Ok(SpannedExpr {
                    span: tok.span,
                    kind: SpannedExprKind::StructuredRef(crate::eval::StructuredRefExpr {
                        sheet: SheetReference::Current,
                        sref: r.clone(),
                    }),
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
                    self.parse_ident_or_field_access(tok.span, id)
                }
            }
            TokenKind::SheetName(_name) => {
                if matches!(self.peek_n(1).kind, TokenKind::Bang)
                    || (matches!(self.peek_n(1).kind, TokenKind::Colon)
                        && matches!(self.peek_n(3).kind, TokenKind::Bang))
                {
                    self.parse_sheet_ref()
                } else {
                    let sheet_name_tok = self.next();
                    let name = match sheet_name_tok.kind {
                        TokenKind::SheetName(name) => name,
                        _ => unreachable!("peeked SheetName then consumed different token"),
                    };

                    if let Some((workbook, name)) = parse_workbook_scoped_external_name_ref(&name) {
                        Ok(SpannedExpr {
                            span: sheet_name_tok.span,
                            kind: SpannedExprKind::NameRef(crate::eval::NameRef {
                                sheet: SheetReference::External(format!("[{workbook}]")),
                                name,
                            }),
                        })
                    } else {
                        Ok(SpannedExpr {
                            span: sheet_name_tok.span,
                            kind: SpannedExprKind::Error(ErrorKind::Name),
                        })
                    }
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

        loop {
            match &self.peek().kind {
                TokenKind::Dot => {
                    self.next(); // '.'
                    expr = self.parse_field_access_after_dot(expr)?;
                }
                TokenKind::Hash => {
                    let hash = self.next();
                    expr = SpannedExpr {
                        span: Span::new(expr.span.start, hash.span.end),
                        kind: SpannedExprKind::SpillRange(Box::new(expr)),
                    };
                }
                _ => break,
            }
        }

        Ok(expr)
    }

    fn parse_ident_or_field_access(
        &mut self,
        span: Span,
        id: &str,
    ) -> Result<SpannedExpr<String>, FormulaParseError> {
        if let Some((workbook, name)) = parse_workbook_scoped_external_name_ref(id) {
            return Ok(SpannedExpr {
                span,
                kind: SpannedExprKind::NameRef(crate::eval::NameRef {
                    sheet: SheetReference::External(format!("[{workbook}]")),
                    name,
                }),
            });
        }

        if id.contains('.') {
            return self.parse_dotted_identifier(span, id);
        }

        match id.to_ascii_uppercase().as_str() {
            "TRUE" => Ok(SpannedExpr {
                span,
                kind: SpannedExprKind::Bool(true),
            }),
            "FALSE" => Ok(SpannedExpr {
                span,
                kind: SpannedExprKind::Bool(false),
            }),
            _ => {
                if let Ok(addr) = parse_a1(id) {
                    return self.parse_cell_or_range(
                        SheetReference::Current,
                        span.start,
                        addr,
                        span.end,
                    );
                }

                // Whole-column / whole-row ranges like `A:C` and `1:3` don't fit `parse_a1`.
                // For debug tracing, treat these as range references only when used with the
                // `:<ref>` range operator (otherwise they remain name refs).
                if matches!(self.peek().kind, TokenKind::Colon) {
                    if let Some(start_col) = parse_col_ref_str(id) {
                        self.next(); // ':'
                        let end_tok = self.next();
                        let end_str = match end_tok.kind {
                            TokenKind::Ident(s) => s,
                            other => {
                                return Err(FormulaParseError::Expected {
                                    expected: "column reference".to_string(),
                                    got: format!("{other:?}"),
                                })
                            }
                        };
                        let Some(end_col) = parse_col_ref_str(&end_str) else {
                            return Err(FormulaParseError::InvalidAddress(
                                crate::eval::AddressParseError::ColumnOutOfRange,
                            ));
                        };
                        let span = Span::new(span.start, end_tok.span.end);
                        let Some(start) = crate::eval::Ref::from_abs_cell_addr(CellAddr {
                            row: 0,
                            col: start_col,
                        }) else {
                            return Ok(SpannedExpr {
                                span,
                                kind: SpannedExprKind::Error(ErrorKind::Ref),
                            });
                        };
                        let Some(end) = crate::eval::Ref::from_abs_cell_addr(CellAddr {
                            row: CellAddr::SHEET_END,
                            col: end_col,
                        }) else {
                            return Ok(SpannedExpr {
                                span,
                                kind: SpannedExprKind::Error(ErrorKind::Ref),
                            });
                        };
                        return Ok(SpannedExpr {
                            span,
                            kind: SpannedExprKind::RangeRef(crate::eval::RangeRef {
                                sheet: SheetReference::Current,
                                start,
                                end,
                            }),
                        });
                    }

                    if let Some(start_row) = parse_row_ref_str(id) {
                        self.next(); // ':'
                        let end_tok = self.next();
                        let end_row = match &end_tok.kind {
                            TokenKind::Number(n) => parse_row_ref_number(*n),
                            TokenKind::Ident(s) => parse_row_ref_str(s),
                            other => {
                                return Err(FormulaParseError::Expected {
                                    expected: "row reference".to_string(),
                                    got: format!("{other:?}"),
                                })
                            }
                        };
                        let Some(end_row) = end_row else {
                            return Err(FormulaParseError::InvalidAddress(
                                crate::eval::AddressParseError::RowOutOfRange,
                            ));
                        };
                        let span = Span::new(span.start, end_tok.span.end);
                        let Some(start) = crate::eval::Ref::from_abs_cell_addr(CellAddr {
                            row: start_row,
                            col: 0,
                        }) else {
                            return Ok(SpannedExpr {
                                span,
                                kind: SpannedExprKind::Error(ErrorKind::Ref),
                            });
                        };
                        let Some(end) = crate::eval::Ref::from_abs_cell_addr(CellAddr {
                            row: end_row,
                            col: CellAddr::SHEET_END,
                        }) else {
                            return Ok(SpannedExpr {
                                span,
                                kind: SpannedExprKind::Error(ErrorKind::Ref),
                            });
                        };
                        return Ok(SpannedExpr {
                            span,
                            kind: SpannedExprKind::RangeRef(crate::eval::RangeRef {
                                sheet: SheetReference::Current,
                                start,
                                end,
                            }),
                        });
                    }
                }

                Ok(SpannedExpr {
                    span,
                    kind: SpannedExprKind::NameRef(crate::eval::NameRef {
                        sheet: SheetReference::Current,
                        name: id.to_string(),
                    }),
                })
            }
        }
    }

    fn parse_dotted_identifier(
        &mut self,
        span: Span,
        raw: &str,
    ) -> Result<SpannedExpr<String>, FormulaParseError> {
        // The lexer permits `.` inside identifiers (e.g. `_xlfn.XLOOKUP`). For formula field access
        // expressions, treat a dotted identifier (`A1.Price.Net`) as nested postfix field accesses.
        let mut parts: Vec<&str> = raw.split('.').collect();
        if parts.len() < 2 {
            return self.parse_ident_or_field_access(span, raw);
        }

        let mut pending_selector = false;
        if parts.last().is_some_and(|p| p.is_empty()) {
            // Identifier ended with a '.', typically from field selectors like `A1.["Price"]` where
            // the lexer stops at `[` and leaves the dot attached to the identifier token.
            parts.pop();
            pending_selector = true;
        }

        if parts.iter().any(|p| p.is_empty()) {
            return Err(FormulaParseError::UnexpectedToken(
                "invalid dotted identifier".to_string(),
            ));
        }

        let base_raw = parts
            .first()
            .copied()
            .expect("split('.') always returns at least one element");
        let base_end = span.start + base_raw.len();
        let base_span = Span::new(span.start, base_end);

        let mut expr = match base_raw.to_ascii_uppercase().as_str() {
            "TRUE" => SpannedExpr {
                span: base_span,
                kind: SpannedExprKind::Bool(true),
            },
            "FALSE" => SpannedExpr {
                span: base_span,
                kind: SpannedExprKind::Bool(false),
            },
            _ => match parse_a1(base_raw) {
                Ok(addr) => {
                    let addr = crate::eval::Ref::from_abs_cell_addr(addr);
                    let kind = match addr {
                        Some(addr) => SpannedExprKind::CellRef(crate::eval::CellRef {
                            sheet: SheetReference::Current,
                            addr,
                        }),
                        None => SpannedExprKind::Error(ErrorKind::Ref),
                    };
                    SpannedExpr {
                        span: base_span,
                        kind,
                    }
                }
                Err(_) => SpannedExpr {
                    span: base_span,
                    kind: SpannedExprKind::NameRef(crate::eval::NameRef {
                        sheet: SheetReference::Current,
                        name: base_raw.to_string(),
                    }),
                },
            },
        };

        // Apply all dotted field segments that are part of the same identifier token.
        let mut cursor = base_raw.len();
        for field in parts.iter().skip(1) {
            cursor += 1; // '.'
            cursor += field.len();
            let end = span.start + cursor;
            expr = SpannedExpr {
                span: Span::new(expr.span.start, end),
                kind: SpannedExprKind::FieldAccess {
                    base: Box::new(expr),
                    field: (*field).to_string(),
                },
            };
        }

        if pending_selector {
            expr = self.parse_field_access_after_dot(expr)?;
        }

        Ok(expr)
    }

    fn parse_dotted_identifier_with_sheet(
        &mut self,
        overall_start: usize,
        addr_span: Span,
        raw: &str,
        sheet: &SheetReference<String>,
    ) -> Result<SpannedExpr<String>, FormulaParseError> {
        // When parsing a sheet-qualified reference like `Sheet1!A1.Price`, the lexer will emit the
        // address portion as a single identifier token (`A1.Price`). We need to split it and apply
        // the sheet reference to the base before building postfix field-access nodes.
        let mut parts: Vec<&str> = raw.split('.').collect();
        debug_assert!(
            parts.len() >= 2,
            "parse_dotted_identifier_with_sheet called without a '.' in the identifier"
        );

        let mut pending_selector = false;
        if parts.last().is_some_and(|p| p.is_empty()) {
            parts.pop();
            pending_selector = true;
        }
        if parts.iter().any(|p| p.is_empty()) {
            return Err(FormulaParseError::UnexpectedToken(
                "invalid dotted identifier".to_string(),
            ));
        }

        let base_raw = parts
            .first()
            .copied()
            .expect("split('.') always returns at least one element");
        let base_end = addr_span.start + base_raw.len();
        let base_span = Span::new(overall_start, base_end);

        let mut expr = match parse_a1(base_raw) {
            Ok(addr) => {
                let addr = crate::eval::Ref::from_abs_cell_addr(addr);
                let kind = match addr {
                    Some(addr) => SpannedExprKind::CellRef(crate::eval::CellRef {
                        sheet: sheet.clone(),
                        addr,
                    }),
                    None => SpannedExprKind::Error(ErrorKind::Ref),
                };
                SpannedExpr {
                    span: base_span,
                    kind,
                }
            }
            Err(_) => SpannedExpr {
                span: base_span,
                kind: SpannedExprKind::NameRef(crate::eval::NameRef {
                    sheet: sheet.clone(),
                    name: base_raw.to_string(),
                }),
            },
        };

        let mut cursor = base_raw.len();
        for field in parts.iter().skip(1) {
            cursor += 1; // '.'
            cursor += field.len();
            let end = addr_span.start + cursor;
            expr = SpannedExpr {
                span: Span::new(expr.span.start, end),
                kind: SpannedExprKind::FieldAccess {
                    base: Box::new(expr),
                    field: (*field).to_string(),
                },
            };
        }

        if pending_selector {
            expr = self.parse_field_access_after_dot(expr)?;
        }

        Ok(expr)
    }

    fn parse_field_access_after_dot(
        &mut self,
        mut base: SpannedExpr<String>,
    ) -> Result<SpannedExpr<String>, FormulaParseError> {
        let selector_tok = self.next();
        match selector_tok.kind {
            TokenKind::Ident(raw) => {
                // `.Ident` selector.
                let segments: Vec<&str> = raw.split('.').collect();
                if segments.iter().any(|s| s.is_empty()) {
                    return Err(FormulaParseError::UnexpectedToken(
                        "invalid field selector".to_string(),
                    ));
                }

                let mut offset = 0usize;
                for (i, segment) in segments.iter().enumerate() {
                    offset += segment.len();
                    let end = selector_tok.span.start + offset;
                    base = SpannedExpr {
                        span: Span::new(base.span.start, end),
                        kind: SpannedExprKind::FieldAccess {
                            base: Box::new(base),
                            field: (*segment).to_string(),
                        },
                    };
                    if i + 1 < segments.len() {
                        offset += 1; // '.'
                    }
                }

                Ok(base)
            }
            TokenKind::StructuredRef(sref) => {
                // `.["..."]` selector.
                let field = match sref {
                    crate::structured_refs::StructuredRef {
                        table_name: None,
                        items,
                        columns: crate::structured_refs::StructuredColumns::Single(name),
                    } if items.is_empty() => {
                        parse_bracket_quoted_field_name(&name).map_err(|_| {
                            FormulaParseError::UnexpectedToken(
                                "expected bracket-quoted field selector".to_string(),
                            )
                        })?
                    }
                    _ => {
                        return Err(FormulaParseError::UnexpectedToken(
                            "expected field selector".to_string(),
                        ));
                    }
                };

                Ok(SpannedExpr {
                    span: Span::new(base.span.start, selector_tok.span.end),
                    kind: SpannedExprKind::FieldAccess {
                        base: Box::new(base),
                        field,
                    },
                })
            }
            other => Err(FormulaParseError::Expected {
                expected: "field selector".to_string(),
                got: format!("{other:?}"),
            }),
        }
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
                if args.len() == crate::EXCEL_MAX_ARGS {
                    return Err(FormulaParseError::UnexpectedToken(format!(
                        "Too many arguments (max {})",
                        crate::EXCEL_MAX_ARGS
                    )));
                }
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
        let start_was_quoted = matches!(&sheet_tok.kind, TokenKind::SheetName(_));
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
            && matches!(
                self.peek().kind,
                TokenKind::Ident(_) | TokenKind::SheetName(_)
            )
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

        // Path-qualified external workbook references can wrap the entire prefix in a single quoted
        // string, e.g. `'C:\path\[Book.xlsx]Sheet1'!A1`.
        //
        // Canonical evaluation normalizes these into the bracketed external key form:
        // `[C:\path\Book.xlsx]Sheet1`.
        let path_qualified_external = start_was_quoted
            .then(|| parse_path_qualified_external_sheet_key(&start_name))
            .flatten();

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
            if let Some((_, start_sheet)) = crate::external_refs::parse_external_key(&start_name) {
                // Excel treats `[Book]Sheet1:Sheet3!A1` as an external workbook 3D span where the
                // bracketed workbook prefix applies to both endpoints. We preserve the span in the
                // external sheet key (`[Book]Sheet1:Sheet3`) so evaluation can expand it using the
                // provider's external sheet order.
                //
                // When the endpoint sheet names match, collapse to the single external sheet key
                // (`[Book]Sheet1`) so evaluation can consult the external provider directly.
                if sheet_name_eq_case_insensitive(start_sheet, &end_name) {
                    SheetReference::External(start_name)
                } else {
                    // Preserve the full span in the sheet key so `resolve_sheet_id` reliably
                    // yields `#REF!`.
                    SheetReference::External(format!("{start_name}:{end_name}"))
                }
            } else if let Some((workbook, start_sheet)) = &path_qualified_external {
                // External workbook 3D span with a path-qualified workbook prefix, e.g.:
                // `'C:\path\[Book.xlsx]Sheet1:Sheet3'!A1` / `'C:\path\[Book.xlsx]Sheet1':Sheet3!A1`.
                //
                // Preserve the span in the external sheet key (`[C:\path\Book.xlsx]Sheet1:Sheet3`)
                // so evaluation can expand it using the provider's external sheet order. When the
                // endpoint sheet names match, collapse to the single external sheet key so
                // evaluation can consult the external provider directly.
                if sheet_name_eq_case_insensitive(start_sheet, &end_name) {
                    SheetReference::External(format!("[{workbook}]{start_sheet}"))
                } else {
                    SheetReference::External(format!("[{workbook}]{start_sheet}:{end_name}"))
                }
            } else {
                SheetReference::SheetRange(start_name, end_name)
            }
        } else {
            self.expect(TokenKind::Bang)?;
            if let Some((workbook, sheet_part)) = &path_qualified_external {
                match split_sheet_span_name(sheet_part) {
                    Some((start, end)) => {
                        if sheet_name_eq_case_insensitive(&start, &end) {
                            SheetReference::External(format!("[{workbook}]{start}"))
                        } else {
                            SheetReference::External(format!("[{workbook}]{start}:{end}"))
                        }
                    }
                    None => SheetReference::External(format!("[{workbook}]{sheet_part}")),
                }
            } else if let Some((start, end)) = split_sheet_span_name(&start_name) {
                if let Some((_, start_sheet)) = crate::external_refs::parse_external_key(&start) {
                    if sheet_name_eq_case_insensitive(start_sheet, &end) {
                        SheetReference::External(start)
                    } else {
                        SheetReference::External(format!("{start}:{end}"))
                    }
                } else {
                    SheetReference::SheetRange(start, end)
                }
            } else if crate::external_refs::parse_external_key(&start_name).is_some() {
                SheetReference::External(start_name)
            } else {
                SheetReference::Sheet(start_name)
            }
        };

        if matches!(self.peek().kind, TokenKind::Ident(ref id) if id.starts_with('[')) {
            self.next();
            return Ok(SpannedExpr {
                span: Span::new(sheet_tok.span.start, sheet_tok.span.end),
                kind: SpannedExprKind::Error(ErrorKind::Ref),
            });
        }

        if matches!(self.peek().kind, TokenKind::StructuredRef(_)) {
            let sref_tok = self.next();
            let sref = match sref_tok.kind {
                TokenKind::StructuredRef(sref) => sref,
                _ => unreachable!("peeked structured ref then consumed different token"),
            };

            return Ok(SpannedExpr {
                span: Span::new(sheet_tok.span.start, sref_tok.span.end),
                kind: SpannedExprKind::StructuredRef(crate::eval::StructuredRefExpr {
                    sheet,
                    sref,
                }),
            });
        }

        let first_tok = self.next();
        match first_tok.kind {
            TokenKind::Ident(addr_str) => {
                if addr_str.contains('.') {
                    return self.parse_dotted_identifier_with_sheet(
                        sheet_tok.span.start,
                        first_tok.span,
                        &addr_str,
                        &sheet,
                    );
                }

                if let Ok(addr) = parse_a1(&addr_str) {
                    return self.parse_cell_or_range(
                        sheet,
                        sheet_tok.span.start,
                        addr,
                        first_tok.span.end,
                    );
                }

                if matches!(self.peek().kind, TokenKind::Colon) {
                    if let Some(start_col) = parse_col_ref_str(&addr_str) {
                        self.next(); // ':'
                        let end_tok = self.next();
                        let end_str = match end_tok.kind {
                            TokenKind::Ident(s) => s,
                            other => {
                                return Err(FormulaParseError::Expected {
                                    expected: "column reference".to_string(),
                                    got: format!("{other:?}"),
                                })
                            }
                        };
                        let Some(end_col) = parse_col_ref_str(&end_str) else {
                            return Err(FormulaParseError::InvalidAddress(
                                crate::eval::AddressParseError::ColumnOutOfRange,
                            ));
                        };
                        let span = Span::new(sheet_tok.span.start, end_tok.span.end);
                        let Some(start) = crate::eval::Ref::from_abs_cell_addr(CellAddr {
                            row: 0,
                            col: start_col,
                        }) else {
                            return Ok(SpannedExpr {
                                span,
                                kind: SpannedExprKind::Error(ErrorKind::Ref),
                            });
                        };
                        let Some(end) = crate::eval::Ref::from_abs_cell_addr(CellAddr {
                            row: CellAddr::SHEET_END,
                            col: end_col,
                        }) else {
                            return Ok(SpannedExpr {
                                span,
                                kind: SpannedExprKind::Error(ErrorKind::Ref),
                            });
                        };
                        return Ok(SpannedExpr {
                            span,
                            kind: SpannedExprKind::RangeRef(crate::eval::RangeRef {
                                sheet,
                                start,
                                end,
                            }),
                        });
                    }

                    if let Some(start_row) = parse_row_ref_str(&addr_str) {
                        self.next(); // ':'
                        let end_tok = self.next();
                        let end_row = match &end_tok.kind {
                            TokenKind::Number(n) => parse_row_ref_number(*n),
                            TokenKind::Ident(s) => parse_row_ref_str(s),
                            other => {
                                return Err(FormulaParseError::Expected {
                                    expected: "row reference".to_string(),
                                    got: format!("{other:?}"),
                                })
                            }
                        };
                        let Some(end_row) = end_row else {
                            return Err(FormulaParseError::InvalidAddress(
                                crate::eval::AddressParseError::RowOutOfRange,
                            ));
                        };
                        let span = Span::new(sheet_tok.span.start, end_tok.span.end);
                        let Some(start) = crate::eval::Ref::from_abs_cell_addr(CellAddr {
                            row: start_row,
                            col: 0,
                        }) else {
                            return Ok(SpannedExpr {
                                span,
                                kind: SpannedExprKind::Error(ErrorKind::Ref),
                            });
                        };
                        let Some(end) = crate::eval::Ref::from_abs_cell_addr(CellAddr {
                            row: end_row,
                            col: CellAddr::SHEET_END,
                        }) else {
                            return Ok(SpannedExpr {
                                span,
                                kind: SpannedExprKind::Error(ErrorKind::Ref),
                            });
                        };
                        return Ok(SpannedExpr {
                            span,
                            kind: SpannedExprKind::RangeRef(crate::eval::RangeRef {
                                sheet,
                                start,
                                end,
                            }),
                        });
                    }
                }

                Ok(SpannedExpr {
                    span: Span::new(sheet_tok.span.start, first_tok.span.end),
                    kind: SpannedExprKind::NameRef(crate::eval::NameRef {
                        sheet,
                        name: addr_str,
                    }),
                })
            }
            TokenKind::Number(n) => {
                // Row ranges like `Sheet1!1:3` are lexed as number literals.
                if matches!(self.peek().kind, TokenKind::Colon) {
                    let Some(start_row) = parse_row_ref_number(n) else {
                        return Err(FormulaParseError::InvalidAddress(
                            crate::eval::AddressParseError::RowOutOfRange,
                        ));
                    };
                    self.next(); // ':'
                    let end_tok = self.next();
                    let end_row = match &end_tok.kind {
                        TokenKind::Number(m) => parse_row_ref_number(*m),
                        TokenKind::Ident(s) => parse_row_ref_str(s),
                        other => {
                            return Err(FormulaParseError::Expected {
                                expected: "row reference".to_string(),
                                got: format!("{other:?}"),
                            })
                        }
                    };
                    let Some(end_row) = end_row else {
                        return Err(FormulaParseError::InvalidAddress(
                            crate::eval::AddressParseError::RowOutOfRange,
                        ));
                    };
                    let span = Span::new(sheet_tok.span.start, end_tok.span.end);
                    let Some(start) = crate::eval::Ref::from_abs_cell_addr(CellAddr {
                        row: start_row,
                        col: 0,
                    }) else {
                        return Ok(SpannedExpr {
                            span,
                            kind: SpannedExprKind::Error(ErrorKind::Ref),
                        });
                    };
                    let Some(end) = crate::eval::Ref::from_abs_cell_addr(CellAddr {
                        row: end_row,
                        col: CellAddr::SHEET_END,
                    }) else {
                        return Ok(SpannedExpr {
                            span,
                            kind: SpannedExprKind::Error(ErrorKind::Ref),
                        });
                    };
                    Ok(SpannedExpr {
                        span,
                        kind: SpannedExprKind::RangeRef(crate::eval::RangeRef {
                            sheet,
                            start,
                            end,
                        }),
                    })
                } else {
                    Err(FormulaParseError::UnexpectedToken("number".to_string()))
                }
            }
            other => Err(FormulaParseError::Expected {
                expected: "cell address".to_string(),
                got: format!("{other:?}"),
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
            let span = Span::new(start_span, end_tok.span.end);
            let Some(start_ref) = crate::eval::Ref::from_abs_cell_addr(start) else {
                return Ok(SpannedExpr {
                    span,
                    kind: SpannedExprKind::Error(ErrorKind::Ref),
                });
            };
            let Some(end_ref) = crate::eval::Ref::from_abs_cell_addr(end) else {
                return Ok(SpannedExpr {
                    span,
                    kind: SpannedExprKind::Error(ErrorKind::Ref),
                });
            };
            Ok(SpannedExpr {
                span,
                kind: SpannedExprKind::RangeRef(crate::eval::RangeRef {
                    sheet,
                    start: start_ref,
                    end: end_ref,
                }),
            })
        } else {
            let span = Span::new(start_span, end_span);
            let Some(addr) = crate::eval::Ref::from_abs_cell_addr(start) else {
                return Ok(SpannedExpr {
                    span,
                    kind: SpannedExprKind::Error(ErrorKind::Ref),
                });
            };
            Ok(SpannedExpr {
                span,
                kind: SpannedExprKind::CellRef(crate::eval::CellRef { sheet, addr }),
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
    date_system: crate::date::ExcelDateSystem,
    value_locale: crate::locale::ValueLocaleConfig,
}

impl<'a, R: crate::eval::ValueResolver> TracedEvaluator<'a, R> {
    fn resolve_range_bounds(
        &self,
        sheet_id: &FnSheetId,
        start: CellAddr,
        end: CellAddr,
    ) -> Option<(CellAddr, CellAddr)> {
        let (rows, cols) = match sheet_id {
            FnSheetId::Local(id) => {
                if !self.resolver.sheet_exists(*id) {
                    return None;
                }
                self.resolver.sheet_dimensions(*id)
            }
            // External workbooks do not expose dimensions via the ValueResolver interface, so treat
            // the bounds as unknown and only resolve whole-row/whole-column sentinels using Excel's
            // default grid size.
            FnSheetId::External(_) => {
                (formula_model::EXCEL_MAX_ROWS, formula_model::EXCEL_MAX_COLS)
            }
        };

        let max_row = rows.saturating_sub(1);
        let max_col = cols.saturating_sub(1);

        let start = CellAddr {
            row: if start.row == CellAddr::SHEET_END {
                max_row
            } else {
                start.row
            },
            col: if start.col == CellAddr::SHEET_END {
                max_col
            } else {
                start.col
            },
        };
        let end = CellAddr {
            row: if end.row == CellAddr::SHEET_END {
                max_row
            } else {
                end.row
            },
            col: if end.col == CellAddr::SHEET_END {
                max_col
            } else {
                end.col
            },
        };

        if matches!(sheet_id, FnSheetId::Local(_))
            && (start.row >= rows || end.row >= rows || start.col >= cols || end.col >= cols)
        {
            return None;
        }

        Some((start, end))
    }

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
        let rows_u64 = u64::from(range.end.row).saturating_sub(u64::from(range.start.row)) + 1;
        let cols_u64 = u64::from(range.end.col).saturating_sub(u64::from(range.start.col)) + 1;
        let cell_count = rows_u64.saturating_mul(cols_u64);
        if cell_count > (MAX_REFERENCE_DEREF_CELLS as u64) {
            return Value::Error(ErrorKind::Spill);
        }
        let rows = rows_u64 as usize;
        let cols = cols_u64 as usize;
        let mut values = Vec::with_capacity(cell_count as usize);
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
                    let Some(addr) = r.addr.resolve(self.ctx.current_cell) else {
                        let value = Value::Error(ErrorKind::Ref);
                        return (
                            EvalValue::Scalar(value.clone()),
                            TraceNode {
                                kind: TraceKind::CellRef,
                                span: expr.span,
                                value,
                                reference: None,
                                children: Vec::new(),
                            },
                        );
                    };
                    let mut ranges = Vec::with_capacity(sheet_ids.len());
                    for sheet_id in sheet_ids {
                        if matches!(&sheet_id, FnSheetId::Local(id) if !self.resolver.sheet_exists(*id))
                        {
                            let value = Value::Error(ErrorKind::Ref);
                            return (
                                EvalValue::Scalar(value.clone()),
                                TraceNode {
                                    kind: TraceKind::CellRef,
                                    span: expr.span,
                                    value,
                                    reference: None,
                                    children: Vec::new(),
                                },
                            );
                        }
                        let Some((start, end)) =
                            self.resolve_range_bounds(&sheet_id, addr, addr)
                        else {
                            let value = Value::Error(ErrorKind::Ref);
                            return (
                                EvalValue::Scalar(value.clone()),
                                TraceNode {
                                    kind: TraceKind::CellRef,
                                    span: expr.span,
                                    value,
                                    reference: None,
                                    children: Vec::new(),
                                },
                            );
                        };
                        ranges.push(ResolvedRange {
                            sheet_id,
                            start,
                            end,
                        });
                    }

                    let reference = if ranges.len() == 1 {
                        Some(TraceRef::Cell {
                            sheet: ranges[0].sheet_id.clone(),
                            addr: ranges[0].start,
                        })
                    } else {
                        None
                    };

                    (
                        EvalValue::Reference(ranges),
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
                    let Some(resolved_start) = r.start.resolve(self.ctx.current_cell) else {
                        let value = Value::Error(ErrorKind::Ref);
                        return (
                            EvalValue::Scalar(value.clone()),
                            TraceNode {
                                kind: TraceKind::RangeRef,
                                span: expr.span,
                                value,
                                reference: None,
                                children: Vec::new(),
                            },
                        );
                    };
                    let Some(resolved_end) = r.end.resolve(self.ctx.current_cell) else {
                        let value = Value::Error(ErrorKind::Ref);
                        return (
                            EvalValue::Scalar(value.clone()),
                            TraceNode {
                                kind: TraceKind::RangeRef,
                                span: expr.span,
                                value,
                                reference: None,
                                children: Vec::new(),
                            },
                        );
                    };
                    let mut ranges = Vec::with_capacity(sheet_ids.len());
                    for sheet_id in sheet_ids {
                        if matches!(&sheet_id, FnSheetId::Local(id) if !self.resolver.sheet_exists(*id))
                        {
                            let value = Value::Error(ErrorKind::Ref);
                            return (
                                EvalValue::Scalar(value.clone()),
                                TraceNode {
                                    kind: TraceKind::RangeRef,
                                    span: expr.span,
                                    value,
                                    reference: None,
                                    children: Vec::new(),
                                },
                            );
                        }
                        let Some((start, end)) =
                            self.resolve_range_bounds(&sheet_id, resolved_start, resolved_end)
                        else {
                            let value = Value::Error(ErrorKind::Ref);
                            return (
                                EvalValue::Scalar(value.clone()),
                                TraceNode {
                                    kind: TraceKind::RangeRef,
                                    span: expr.span,
                                    value,
                                    reference: None,
                                    children: Vec::new(),
                                },
                            );
                        };
                        ranges.push(ResolvedRange {
                            sheet_id,
                            start,
                            end,
                        });
                    }

                    let reference = if ranges.len() == 1 {
                        Some(TraceRef::Range {
                            sheet: ranges[0].sheet_id.clone(),
                            start: ranges[0].start,
                            end: ranges[0].end,
                        })
                    } else {
                        None
                    };

                    (
                        EvalValue::Reference(ranges),
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
            SpannedExprKind::StructuredRef(sref_expr) => {
                // External workbook structured references are resolved dynamically using
                // provider-supplied table metadata.
                if let SheetReference::External(key) = &sref_expr.sheet {
                    if !key.starts_with('[') {
                        let value = Value::Error(ErrorKind::Ref);
                        return (
                            EvalValue::Scalar(value.clone()),
                            TraceNode {
                                kind: TraceKind::StructuredRef,
                                span: expr.span,
                                value,
                                reference: None,
                                children: Vec::new(),
                            },
                        );
                    }

                    let (workbook, explicit_sheet_key) =
                        if let Some((workbook, _sheet)) = crate::external_refs::parse_external_key(key)
                        {
                            (workbook, Some(key.as_str()))
                        } else if crate::external_refs::parse_external_span_key(key).is_some() {
                            let value = Value::Error(ErrorKind::Ref);
                            return (
                                EvalValue::Scalar(value.clone()),
                                TraceNode {
                                    kind: TraceKind::StructuredRef,
                                    span: expr.span,
                                    value,
                                    reference: None,
                                    children: Vec::new(),
                                },
                            );
                        } else {
                            let Some(workbook) = crate::external_refs::parse_external_workbook_key(key)
                            else {
                                let value = Value::Error(ErrorKind::Ref);
                                return (
                                    EvalValue::Scalar(value.clone()),
                                    TraceNode {
                                        kind: TraceKind::StructuredRef,
                                        span: expr.span,
                                        value,
                                        reference: None,
                                        children: Vec::new(),
                                    },
                                );
                            };
                            (workbook, None)
                        };

                    let Some(table_name) = sref_expr.sref.table_name.as_deref() else {
                        let value = Value::Error(ErrorKind::Ref);
                        return (
                            EvalValue::Scalar(value.clone()),
                            TraceNode {
                                kind: TraceKind::StructuredRef,
                                span: expr.span,
                                value,
                                reference: None,
                                children: Vec::new(),
                            },
                        );
                    };

                    // We do not currently support `[@ThisRow]` semantics for external workbooks.
                    if sref_expr
                        .sref
                        .items
                        .iter()
                        .any(|item| matches!(item, crate::structured_refs::StructuredRefItem::ThisRow))
                    {
                        let value = Value::Error(ErrorKind::Ref);
                        return (
                            EvalValue::Scalar(value.clone()),
                            TraceNode {
                                kind: TraceKind::StructuredRef,
                                span: expr.span,
                                value,
                                reference: None,
                                children: Vec::new(),
                            },
                        );
                    }

                    let Some((table_sheet, table)) =
                        self.resolver.external_workbook_table(workbook, table_name)
                    else {
                        let value = Value::Error(ErrorKind::Ref);
                        return (
                            EvalValue::Scalar(value.clone()),
                            TraceNode {
                                kind: TraceKind::StructuredRef,
                                span: expr.span,
                                value,
                                reference: None,
                                children: Vec::new(),
                            },
                        );
                    };

                    let sheet_key = explicit_sheet_key
                        .map(|s| s.to_string())
                        .unwrap_or_else(|| format!("[{workbook}]{table_sheet}"));

                    let ranges = match crate::structured_refs::resolve_structured_ref_in_table(
                        &table,
                        self.ctx.current_cell,
                        &sref_expr.sref,
                    ) {
                        Ok(ranges) => ranges,
                        Err(_) => {
                            let value = Value::Error(ErrorKind::Ref);
                            return (
                                EvalValue::Scalar(value.clone()),
                                TraceNode {
                                    kind: TraceKind::StructuredRef,
                                    span: expr.span,
                                    value,
                                    reference: None,
                                    children: Vec::new(),
                                },
                            );
                        }
                    };

                    let resolved: Vec<ResolvedRange> = ranges
                        .into_iter()
                        .map(|(start, end)| ResolvedRange {
                            sheet_id: FnSheetId::External(sheet_key.clone()),
                            start,
                            end,
                        })
                        .collect();

                    let reference = match resolved.as_slice() {
                        [only] if only.is_single_cell() => Some(TraceRef::Cell {
                            sheet: only.sheet_id.clone(),
                            addr: only.start,
                        }),
                        [only] => Some(TraceRef::Range {
                            sheet: only.sheet_id.clone(),
                            start: only.start,
                            end: only.end,
                        }),
                        _ => None,
                    };

                    return (
                        EvalValue::Reference(resolved),
                        TraceNode {
                            kind: TraceKind::StructuredRef,
                            span: expr.span,
                            value: Value::Blank,
                            reference,
                            children: Vec::new(),
                        },
                    );
                }

                // Local structured refs resolve via workbook table metadata.
                match self.resolver.resolve_structured_ref(self.ctx, &sref_expr.sref) {
                    Ok(ranges)
                        if !ranges.is_empty()
                            && ranges
                                .iter()
                                .all(|(sheet_id, _, _)| self.resolver.sheet_exists(*sheet_id)) =>
                    {
                        let resolved: Vec<ResolvedRange> = ranges
                            .into_iter()
                            .map(|(sheet_id, start, end)| ResolvedRange {
                                sheet_id: FnSheetId::Local(sheet_id),
                                start,
                                end,
                            })
                            .collect();

                        let reference = match resolved.as_slice() {
                            [only] if only.is_single_cell() => Some(TraceRef::Cell {
                                sheet: only.sheet_id.clone(),
                                addr: only.start,
                            }),
                            [only] => Some(TraceRef::Range {
                                sheet: only.sheet_id.clone(),
                                start: only.start,
                                end: only.end,
                            }),
                            _ => None,
                        };

                        (
                            EvalValue::Reference(resolved),
                            TraceNode {
                                kind: TraceKind::StructuredRef,
                                span: expr.span,
                                value: Value::Blank,
                                reference,
                                children: Vec::new(),
                            },
                        )
                    }
                    Ok(_) => {
                        let value = Value::Error(ErrorKind::Ref);
                        (
                            EvalValue::Scalar(value.clone()),
                            TraceNode {
                                kind: TraceKind::StructuredRef,
                                span: expr.span,
                                value,
                                reference: None,
                                children: Vec::new(),
                            },
                        )
                    }
                    Err(e) => {
                        let value = Value::Error(e);
                        (
                            EvalValue::Scalar(value.clone()),
                            TraceNode {
                                kind: TraceKind::StructuredRef,
                                span: expr.span,
                                value,
                                reference: None,
                                children: Vec::new(),
                            },
                        )
                    }
                }
            }
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
            SpannedExprKind::FieldAccess { base, field } => {
                let (ev, child) = self.eval_value(base);
                let base_value = self.deref_eval_value_dynamic(ev);
                let field_name = field.clone();

                let out = elementwise_unary(&base_value, |elem| match elem {
                    Value::Error(e) => Value::Error(*e),
                    Value::Record(r) => r
                        .fields
                        .iter()
                        .find(|(k, _)| {
                            crate::value::cmp_case_insensitive(k, &field_name) == Ordering::Equal
                        })
                        .map(|(_, v)| v.clone())
                        .unwrap_or(Value::Error(ErrorKind::Field)),
                    Value::Entity(e) => e
                        .fields
                        .iter()
                        .find(|(k, _)| {
                            crate::value::cmp_case_insensitive(k, &field_name) == Ordering::Equal
                        })
                        .map(|(_, v)| v.clone())
                        .unwrap_or(Value::Error(ErrorKind::Field)),
                    // Field access on a non-rich value yields `#VALUE!` (type mismatch). `#FIELD!`
                    // is reserved for missing fields on rich values.
                    _ => Value::Error(ErrorKind::Value),
                });

                (
                    EvalValue::Scalar(out.clone()),
                    TraceNode {
                        kind: TraceKind::FieldAccess {
                            field: field.clone(),
                        },
                        span: expr.span,
                        value: out,
                        reference: None,
                        children: vec![child],
                    },
                )
            }
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
                let out = elementwise_unary(&value, |elem| self.numeric_unary(*op, elem));
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

                let out = match op {
                    crate::eval::BinaryOp::Add
                    | crate::eval::BinaryOp::Sub
                    | crate::eval::BinaryOp::Mul
                    | crate::eval::BinaryOp::Div
                    | crate::eval::BinaryOp::Pow => {
                        elementwise_binary(&l, &r, |a, b| self.numeric_binary(*op, a, b))
                    }
                    crate::eval::BinaryOp::Concat => {
                        elementwise_binary(&l, &r, |a, b| self.concat_binary(a, b))
                    }
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
            SheetReference::External(key) => crate::eval::is_valid_external_single_sheet_key(key)
                .then(|| FnSheetId::External(key.clone())),
        }
    }

    fn resolve_sheet_ids(&self, sheet: &SheetReference<usize>) -> Option<Vec<FnSheetId>> {
        match sheet {
            SheetReference::Current => Some(vec![FnSheetId::Local(self.ctx.current_sheet)]),
            SheetReference::Sheet(id) => Some(vec![FnSheetId::Local(*id)]),
            SheetReference::SheetRange(a, b) => self
                .resolver
                .expand_sheet_span(*a, *b)
                .map(|ids| ids.into_iter().map(FnSheetId::Local).collect()),
            SheetReference::External(key) => {
                if crate::eval::is_valid_external_single_sheet_key(key) {
                    return Some(vec![FnSheetId::External(key.clone())]);
                }

                let (workbook, start, end) = crate::external_refs::parse_external_span_key(key)?;
                let order = self.resolver.workbook_sheet_names(workbook)?;

                let start_key = formula_model::sheet_name_casefold(start);
                let end_key = formula_model::sheet_name_casefold(end);
                let mut start_idx: Option<usize> = None;
                let mut end_idx: Option<usize> = None;
                for (idx, name) in order.iter().enumerate() {
                    let name_key = formula_model::sheet_name_casefold(name);
                    if start_idx.is_none() && name_key == start_key {
                        start_idx = Some(idx);
                    }
                    if end_idx.is_none() && name_key == end_key {
                        end_idx = Some(idx);
                    }
                    if start_idx.is_some() && end_idx.is_some() {
                        break;
                    }
                }
                let start_idx = start_idx?;
                let end_idx = end_idx?;
                let (start_idx, end_idx) = if start_idx <= end_idx {
                    (start_idx, end_idx)
                } else {
                    (end_idx, start_idx)
                };

                Some(
                    order[start_idx..=end_idx]
                        .iter()
                        .map(|sheet| FnSheetId::External(format!("[{workbook}]{sheet}")))
                        .collect(),
                )
            }
        }
    }

    fn deref_reference_scalar(&self, ranges: &[ResolvedRange]) -> Value {
        match ranges {
            [only] if only.is_single_cell() => {
                self.get_sheet_cell_value(&only.sheet_id, only.start)
            }
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
            "RTD" => self.fn_rtd(args),
            "SUM" => self.fn_sum(args),
            "CUBEVALUE" => self.fn_cubevalue(args),
            "CUBEMEMBER" => self.fn_cubemember(args),
            "CUBEMEMBERPROPERTY" => self.fn_cubememberproperty(args),
            "CUBERANKEDMEMBER" => self.fn_cuberankedmember(args),
            "CUBESET" => self.fn_cubeset(args),
            "CUBESETCOUNT" => self.fn_cubesetcount(args),
            "CUBEKPIMEMBER" => self.fn_cubekpimember(args),
            "VLOOKUP" => self.fn_vlookup(args),
            _ => (Value::Error(ErrorKind::Name), Vec::new()),
        }
    }

    fn fn_rtd(&self, args: &[SpannedExpr<usize>]) -> (Value, Vec<TraceNode>) {
        if args.len() < 3 {
            return (Value::Error(ErrorKind::Value), Vec::new());
        }

        let Some(provider) = self.resolver.external_data_provider() else {
            return (Value::Error(ErrorKind::NA), Vec::new());
        };

        let mut traces = Vec::with_capacity(args.len());

        let (prog_id, trace) = self.eval_scalar(&args[0]);
        traces.push(trace);
        if let Value::Error(e) = prog_id {
            return (Value::Error(e), traces);
        }
        let prog_id = match prog_id.coerce_to_string() {
            Ok(v) => v,
            Err(e) => return (Value::Error(e), traces),
        };

        let (server, trace) = self.eval_scalar(&args[1]);
        traces.push(trace);
        if let Value::Error(e) = server {
            return (Value::Error(e), traces);
        }
        let server = match server.coerce_to_string() {
            Ok(v) => v,
            Err(e) => return (Value::Error(e), traces),
        };

        let mut topics = Vec::with_capacity(args.len().saturating_sub(2));
        for arg in &args[2..] {
            let (topic, trace) = self.eval_scalar(arg);
            traces.push(trace);
            if let Value::Error(e) = topic {
                return (Value::Error(e), traces);
            }
            match topic.coerce_to_string() {
                Ok(v) => topics.push(v),
                Err(e) => return (Value::Error(e), traces),
            }
        }

        let out = provider.rtd(&prog_id, &server, &topics);
        (out, traces)
    }

    fn fn_cubevalue(&self, args: &[SpannedExpr<usize>]) -> (Value, Vec<TraceNode>) {
        if args.len() < 2 {
            return (Value::Error(ErrorKind::Value), Vec::new());
        }

        let Some(provider) = self.resolver.external_data_provider() else {
            return (Value::Error(ErrorKind::NA), Vec::new());
        };

        let mut traces = Vec::with_capacity(args.len());

        let (connection, trace) = self.eval_scalar(&args[0]);
        traces.push(trace);
        if let Value::Error(e) = connection {
            return (Value::Error(e), traces);
        }
        let connection = match connection.coerce_to_string() {
            Ok(v) => v,
            Err(e) => return (Value::Error(e), traces),
        };

        let mut tuples = Vec::with_capacity(args.len().saturating_sub(1));
        for arg in &args[1..] {
            let (tuple, trace) = self.eval_scalar(arg);
            traces.push(trace);
            if let Value::Error(e) = tuple {
                return (Value::Error(e), traces);
            }
            match tuple.coerce_to_string() {
                Ok(v) => tuples.push(v),
                Err(e) => return (Value::Error(e), traces),
            }
        }

        let out = provider.cube_value(&connection, &tuples);
        (out, traces)
    }

    fn fn_cubemember(&self, args: &[SpannedExpr<usize>]) -> (Value, Vec<TraceNode>) {
        if args.len() < 2 || args.len() > 3 {
            return (Value::Error(ErrorKind::Value), Vec::new());
        }

        let Some(provider) = self.resolver.external_data_provider() else {
            return (Value::Error(ErrorKind::NA), Vec::new());
        };

        let mut traces = Vec::with_capacity(args.len());

        let (connection, trace) = self.eval_scalar(&args[0]);
        traces.push(trace);
        if let Value::Error(e) = connection {
            return (Value::Error(e), traces);
        }
        let connection = match connection.coerce_to_string() {
            Ok(v) => v,
            Err(e) => return (Value::Error(e), traces),
        };

        let (member_expression, trace) = self.eval_scalar(&args[1]);
        traces.push(trace);
        if let Value::Error(e) = member_expression {
            return (Value::Error(e), traces);
        }
        let member_expression = match member_expression.coerce_to_string() {
            Ok(v) => v,
            Err(e) => return (Value::Error(e), traces),
        };

        let caption = if args.len() >= 3 {
            let (caption, trace) = self.eval_scalar(&args[2]);
            traces.push(trace);
            if let Value::Error(e) = caption {
                return (Value::Error(e), traces);
            }
            match caption.coerce_to_string() {
                Ok(v) => Some(v),
                Err(e) => return (Value::Error(e), traces),
            }
        } else {
            None
        };

        let out = provider.cube_member(&connection, &member_expression, caption.as_deref());
        (out, traces)
    }

    fn fn_cubememberproperty(&self, args: &[SpannedExpr<usize>]) -> (Value, Vec<TraceNode>) {
        if args.len() != 3 {
            return (Value::Error(ErrorKind::Value), Vec::new());
        }

        let Some(provider) = self.resolver.external_data_provider() else {
            return (Value::Error(ErrorKind::NA), Vec::new());
        };

        let mut traces = Vec::with_capacity(args.len());

        let (connection, trace) = self.eval_scalar(&args[0]);
        traces.push(trace);
        if let Value::Error(e) = connection {
            return (Value::Error(e), traces);
        }
        let connection = match connection.coerce_to_string() {
            Ok(v) => v,
            Err(e) => return (Value::Error(e), traces),
        };

        let (member_expression_or_handle, trace) = self.eval_scalar(&args[1]);
        traces.push(trace);
        if let Value::Error(e) = member_expression_or_handle {
            return (Value::Error(e), traces);
        }
        let member_expression_or_handle = match member_expression_or_handle.coerce_to_string() {
            Ok(v) => v,
            Err(e) => return (Value::Error(e), traces),
        };

        let (property, trace) = self.eval_scalar(&args[2]);
        traces.push(trace);
        if let Value::Error(e) = property {
            return (Value::Error(e), traces);
        }
        let property = match property.coerce_to_string() {
            Ok(v) => v,
            Err(e) => return (Value::Error(e), traces),
        };

        let out =
            provider.cube_member_property(&connection, &member_expression_or_handle, &property);
        (out, traces)
    }

    fn fn_cuberankedmember(&self, args: &[SpannedExpr<usize>]) -> (Value, Vec<TraceNode>) {
        if args.len() < 3 || args.len() > 4 {
            return (Value::Error(ErrorKind::Value), Vec::new());
        }

        let Some(provider) = self.resolver.external_data_provider() else {
            return (Value::Error(ErrorKind::NA), Vec::new());
        };

        let mut traces = Vec::with_capacity(args.len());

        // connection
        let (conn, trace) = self.eval_scalar(&args[0]);
        traces.push(trace);
        if let Value::Error(e) = conn {
            return (Value::Error(e), traces);
        }
        let conn = match conn.coerce_to_string() {
            Ok(s) => s,
            Err(e) => return (Value::Error(e), traces),
        };

        // set_expression
        let (set_expr, trace) = self.eval_scalar(&args[1]);
        traces.push(trace);
        if let Value::Error(e) = set_expr {
            return (Value::Error(e), traces);
        }
        let set_expr = match set_expr.coerce_to_string() {
            Ok(s) => s,
            Err(e) => return (Value::Error(e), traces),
        };

        // rank (numeric)
        let (rank, trace) = self.eval_scalar(&args[2]);
        traces.push(trace);
        if let Value::Error(e) = rank {
            return (Value::Error(e), traces);
        }
        let rank = match self.coerce_to_number_with_ctx(&rank) {
            Ok(n) => n.trunc() as i64,
            Err(e) => return (Value::Error(e), traces),
        };

        // optional caption
        let caption = if args.len() >= 4 {
            let (caption, trace) = self.eval_scalar(&args[3]);
            traces.push(trace);
            if let Value::Error(e) = caption {
                return (Value::Error(e), traces);
            }
            let caption = match caption.coerce_to_string() {
                Ok(s) => s,
                Err(e) => return (Value::Error(e), traces),
            };
            Some(caption)
        } else {
            None
        };

        let out = provider.cube_ranked_member(&conn, &set_expr, rank, caption.as_deref());
        (out, traces)
    }

    fn fn_cubeset(&self, args: &[SpannedExpr<usize>]) -> (Value, Vec<TraceNode>) {
        if args.len() < 2 || args.len() > 5 {
            return (Value::Error(ErrorKind::Value), Vec::new());
        }

        let Some(provider) = self.resolver.external_data_provider() else {
            return (Value::Error(ErrorKind::NA), Vec::new());
        };

        let mut traces = Vec::with_capacity(args.len());

        // connection
        let (conn, trace) = self.eval_scalar(&args[0]);
        traces.push(trace);
        if let Value::Error(e) = conn {
            return (Value::Error(e), traces);
        }
        let conn = match conn.coerce_to_string() {
            Ok(s) => s,
            Err(e) => return (Value::Error(e), traces),
        };

        // set_expression
        let (set_expr, trace) = self.eval_scalar(&args[1]);
        traces.push(trace);
        if let Value::Error(e) = set_expr {
            return (Value::Error(e), traces);
        }
        let set_expr = match set_expr.coerce_to_string() {
            Ok(s) => s,
            Err(e) => return (Value::Error(e), traces),
        };

        // optional caption (string)
        let caption = if args.len() >= 3 {
            let (caption, trace) = self.eval_scalar(&args[2]);
            traces.push(trace);
            if let Value::Error(e) = caption {
                return (Value::Error(e), traces);
            }
            let caption = match caption.coerce_to_string() {
                Ok(s) => s,
                Err(e) => return (Value::Error(e), traces),
            };
            Some(caption)
        } else {
            None
        };

        // optional sort_order (numeric)
        let sort_order = if args.len() >= 4 {
            let (order, trace) = self.eval_scalar(&args[3]);
            traces.push(trace);
            if let Value::Error(e) = order {
                return (Value::Error(e), traces);
            }
            let order = match self.coerce_to_number_with_ctx(&order) {
                Ok(n) => n.trunc() as i64,
                Err(e) => return (Value::Error(e), traces),
            };
            Some(order)
        } else {
            None
        };

        // optional sort_by (string)
        let sort_by = if args.len() >= 5 {
            let (sort_by, trace) = self.eval_scalar(&args[4]);
            traces.push(trace);
            if let Value::Error(e) = sort_by {
                return (Value::Error(e), traces);
            }
            let sort_by = match sort_by.coerce_to_string() {
                Ok(s) => s,
                Err(e) => return (Value::Error(e), traces),
            };
            Some(sort_by)
        } else {
            None
        };

        let out = provider.cube_set(
            &conn,
            &set_expr,
            caption.as_deref(),
            sort_order,
            sort_by.as_deref(),
        );
        (out, traces)
    }

    fn fn_cubesetcount(&self, args: &[SpannedExpr<usize>]) -> (Value, Vec<TraceNode>) {
        if args.len() != 1 {
            return (Value::Error(ErrorKind::Value), Vec::new());
        }

        let Some(provider) = self.resolver.external_data_provider() else {
            return (Value::Error(ErrorKind::NA), Vec::new());
        };

        let (set, trace) = self.eval_scalar(&args[0]);
        let traces = vec![trace];
        if let Value::Error(e) = set {
            return (Value::Error(e), traces);
        }
        let set = match set.coerce_to_string() {
            Ok(s) => s,
            Err(e) => return (Value::Error(e), traces),
        };

        let out = provider.cube_set_count(&set);
        (out, traces)
    }

    fn fn_cubekpimember(&self, args: &[SpannedExpr<usize>]) -> (Value, Vec<TraceNode>) {
        if args.len() < 3 || args.len() > 4 {
            return (Value::Error(ErrorKind::Value), Vec::new());
        }

        let Some(provider) = self.resolver.external_data_provider() else {
            return (Value::Error(ErrorKind::NA), Vec::new());
        };

        let mut traces = Vec::with_capacity(args.len());

        let (connection, trace) = self.eval_scalar(&args[0]);
        traces.push(trace);
        if let Value::Error(e) = connection {
            return (Value::Error(e), traces);
        }
        let connection = match connection.coerce_to_string() {
            Ok(v) => v,
            Err(e) => return (Value::Error(e), traces),
        };

        let (kpi_name, trace) = self.eval_scalar(&args[1]);
        traces.push(trace);
        if let Value::Error(e) = kpi_name {
            return (Value::Error(e), traces);
        }
        let kpi_name = match kpi_name.coerce_to_string() {
            Ok(v) => v,
            Err(e) => return (Value::Error(e), traces),
        };

        let (kpi_property, trace) = self.eval_scalar(&args[2]);
        traces.push(trace);
        if let Value::Error(e) = kpi_property {
            return (Value::Error(e), traces);
        }
        let kpi_property = match kpi_property.coerce_to_string() {
            Ok(v) => v,
            Err(e) => return (Value::Error(e), traces),
        };

        let caption = if args.len() == 4 {
            let (caption, trace) = self.eval_scalar(&args[3]);
            traces.push(trace);
            if let Value::Error(e) = caption {
                return (Value::Error(e), traces);
            }
            match caption.coerce_to_string() {
                Ok(v) => Some(v),
                Err(e) => return (Value::Error(e), traces),
            }
        } else {
            None
        };

        let out =
            provider.cube_kpi_member(&connection, &kpi_name, &kpi_property, caption.as_deref());
        (out, traces)
    }

    fn fn_if(&self, args: &[SpannedExpr<usize>]) -> (Value, Vec<TraceNode>) {
        if args.is_empty() {
            return (Value::Error(ErrorKind::Value), Vec::new());
        }
        let (cond_val, cond_trace) = self.eval_scalar(&args[0]);
        if let Value::Error(e) = cond_val {
            return (Value::Error(e), vec![cond_trace]);
        }
        let cond = match self.coerce_to_bool_with_ctx(&cond_val) {
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
                        let n = match self.coerce_text_to_number(&s) {
                            Ok(n) => n,
                            Err(e) => return (Value::Error(e), traces),
                        };
                        acc += n;
                    }
                    Value::Record(_) | Value::Entity(_) => {
                        return (Value::Error(ErrorKind::Value), traces)
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
                                | Value::Record(_)
                                | Value::Entity(_)
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
                                | Value::Record(_)
                                | Value::Entity(_)
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
        let col_index_num = match self.coerce_to_number_with_ctx(&col_index_val) {
            Ok(n) => n,
            Err(e) => return (Value::Error(e), traces),
        };
        let col_index = col_index_num.trunc() as i64;
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
            match self.coerce_to_bool_with_ctx(&v) {
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

    fn coerce_text_to_number(&self, text: &str) -> Result<f64, ErrorKind> {
        let trimmed = text.trim();
        if trimmed.is_empty() {
            return Ok(0.0);
        }

        crate::coercion::datetime::parse_value_text(
            trimmed,
            self.value_locale,
            self.recalc_ctx.now_utc,
            self.date_system,
        )
        .map_err(map_excel_error)
    }

    fn coerce_to_number_with_ctx(&self, value: &Value) -> Result<f64, ErrorKind> {
        match value {
            Value::Number(n) => Ok(*n),
            Value::Bool(b) => Ok(if *b { 1.0 } else { 0.0 }),
            Value::Blank => Ok(0.0),
            Value::Text(s) => self.coerce_text_to_number(s),
            Value::Entity(_) | Value::Record(_) => Err(ErrorKind::Value),
            Value::Error(e) => Err(*e),
            Value::Reference(_)
            | Value::ReferenceUnion(_)
            | Value::Array(_)
            | Value::Lambda(_)
            | Value::Spill { .. } => Err(ErrorKind::Value),
        }
    }

    fn coerce_to_bool_with_ctx(&self, value: &Value) -> Result<bool, ErrorKind> {
        match value {
            Value::Bool(b) => Ok(*b),
            Value::Number(n) => Ok(*n != 0.0),
            Value::Blank => Ok(false),
            Value::Text(s) => {
                let t = s.trim();
                if t.is_empty() {
                    return Ok(false);
                }
                if t.eq_ignore_ascii_case("TRUE") {
                    return Ok(true);
                }
                if t.eq_ignore_ascii_case("FALSE") {
                    return Ok(false);
                }
                Ok(self.coerce_text_to_number(t)? != 0.0)
            }
            Value::Entity(_) | Value::Record(_) => Err(ErrorKind::Value),
            Value::Error(e) => Err(*e),
            Value::Reference(_)
            | Value::ReferenceUnion(_)
            | Value::Array(_)
            | Value::Lambda(_)
            | Value::Spill { .. } => Err(ErrorKind::Value),
        }
    }

    fn coerce_to_string_with_ctx(&self, value: &Value) -> Result<String, ErrorKind> {
        match value {
            Value::Text(s) => Ok(s.clone()),
            Value::Entity(v) => Ok(v.display.clone()),
            Value::Record(v) => Ok(v.display.clone()),
            Value::Number(n) => {
                let options = formula_format::FormatOptions {
                    locale: self.value_locale.separators,
                    date_system: match self.date_system {
                        crate::date::ExcelDateSystem::Excel1900 { .. } => {
                            formula_format::DateSystem::Excel1900
                        }
                        crate::date::ExcelDateSystem::Excel1904 => {
                            formula_format::DateSystem::Excel1904
                        }
                    },
                };
                Ok(
                    formula_format::format_value(formula_format::Value::Number(*n), None, &options)
                        .text,
                )
            }
            Value::Bool(b) => Ok(if *b { "TRUE" } else { "FALSE" }.to_string()),
            Value::Blank => Ok(String::new()),
            Value::Error(e) => Err(*e),
            Value::Reference(_)
            | Value::ReferenceUnion(_)
            | Value::Array(_)
            | Value::Lambda(_)
            | Value::Spill { .. } => Err(ErrorKind::Value),
        }
    }

    fn numeric_unary(&self, op: UnaryOp, value: &Value) -> Value {
        match value {
            Value::Error(e) => Value::Error(*e),
            other => match self.coerce_to_number_with_ctx(other) {
                Ok(n) => match op {
                    UnaryOp::Plus => Value::Number(n),
                    UnaryOp::Minus => Value::Number(-n),
                },
                Err(e) => Value::Error(e),
            },
        }
    }

    fn concat_binary(&self, left: &Value, right: &Value) -> Value {
        if let Value::Error(e) = left {
            return Value::Error(*e);
        }
        if let Value::Error(e) = right {
            return Value::Error(*e);
        }

        let ls = match self.coerce_to_string_with_ctx(left) {
            Ok(s) => s,
            Err(e) => return Value::Error(e),
        };
        let rs = match self.coerce_to_string_with_ctx(right) {
            Ok(s) => s,
            Err(e) => return Value::Error(e),
        };
        Value::Text(format!("{ls}{rs}"))
    }

    fn numeric_binary(&self, op: crate::eval::BinaryOp, left: &Value, right: &Value) -> Value {
        if let Value::Error(e) = left {
            return Value::Error(*e);
        }
        if let Value::Error(e) = right {
            return Value::Error(*e);
        }

        let ln = match self.coerce_to_number_with_ctx(left) {
            Ok(n) => n,
            Err(e) => return Value::Error(e),
        };
        let rn = match self.coerce_to_number_with_ctx(right) {
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
                Err(e) => Value::Error(map_excel_error(e)),
            },
            _ => Value::Error(ErrorKind::Value),
        }
    }
}

fn elementwise_unary(value: &Value, f: impl Fn(&Value) -> Value) -> Value {
    match value {
        Value::Array(arr) => {
            Value::Array(Array::new(arr.rows, arr.cols, arr.iter().map(f).collect()))
        }
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

fn map_excel_error(error: ExcelError) -> ErrorKind {
    match error {
        ExcelError::Div0 => ErrorKind::Div0,
        ExcelError::Value => ErrorKind::Value,
        ExcelError::Num => ErrorKind::Num,
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

    // Treat rich values as text for comparison semantics.
    let left = match left.clone() {
        Value::Entity(v) => Value::Text(v.display),
        Value::Record(v) => Value::Text(v.display),
        other => other,
    };
    let right = match right.clone() {
        Value::Entity(v) => Value::Text(v.display),
        Value::Record(v) => Value::Text(v.display),
        other => other,
    };
    if matches!(
        &left,
        Value::Array(_)
            | Value::Lambda(_)
            | Value::Spill { .. }
            | Value::Reference(_)
            | Value::ReferenceUnion(_)
    ) || matches!(
        &right,
        Value::Array(_)
            | Value::Lambda(_)
            | Value::Spill { .. }
            | Value::Reference(_)
            | Value::ReferenceUnion(_)
    ) {
        return Err(ErrorKind::Value);
    }

    let (l, r) = match (left, right) {
        (Value::Blank, Value::Number(b)) => (Value::Number(0.0), Value::Number(b)),
        (Value::Number(a), Value::Blank) => (Value::Number(a), Value::Number(0.0)),
        (Value::Blank, Value::Bool(b)) => (Value::Bool(false), Value::Bool(b)),
        (Value::Bool(a), Value::Blank) => (Value::Bool(a), Value::Bool(false)),
        (Value::Blank, Value::Text(b)) => (Value::Text(String::new()), Value::Text(b)),
        (Value::Text(a), Value::Blank) => (Value::Text(a), Value::Text(String::new())),
        (l, r) => (l, r),
    };

    Ok(match (&l, &r) {
        (Value::Number(a), Value::Number(b)) => a.partial_cmp(b).unwrap_or(Ordering::Equal),
        (Value::Text(a), Value::Text(b)) => crate::value::cmp_case_insensitive(a, b),
        (Value::Bool(a), Value::Bool(b)) => a.cmp(b),
        (Value::Number(_), Value::Text(_) | Value::Bool(_)) => Ordering::Less,
        (Value::Text(_), Value::Bool(_)) => Ordering::Less,
        (Value::Text(_), Value::Number(_)) => Ordering::Greater,
        (Value::Bool(_), Value::Number(_) | Value::Text(_)) => Ordering::Greater,
        (Value::Blank, Value::Blank) => Ordering::Equal,
        (Value::Blank, _) => Ordering::Less,
        (_, Value::Blank) => Ordering::Greater,
        _ => Ordering::Equal,
    })
}

#[cfg(test)]
mod tests {
    use super::*;

    #[test]
    fn parse_spanned_formula_rejects_function_calls_with_more_than_255_args() {
        let args = std::iter::repeat("1")
            .take(crate::EXCEL_MAX_ARGS + 1)
            .collect::<Vec<_>>()
            .join(",");
        let formula = format!("=SUM({args})");

        let err = parse_spanned_formula(&formula).unwrap_err();
        assert!(
            err.to_string().contains("Too many arguments"),
            "unexpected error: {err}"
        );
    }

    #[test]
    fn lexes_getting_data_error_literal_as_single_token() {
        let formula = "=#GETTING_DATA";
        let mut lexer = Lexer::new(formula);
        let tokens = lexer.tokenize().expect("tokenize should succeed");

        // One error literal + implicit End token.
        assert_eq!(tokens.len(), 2);
        assert_eq!(tokens[0].kind, TokenKind::Error(ErrorKind::GettingData));
        assert_eq!(
            &formula[tokens[0].span.start..tokens[0].span.end],
            "#GETTING_DATA"
        );
    }

    #[test]
    fn lexes_getting_data_error_literal_case_insensitive() {
        let formula = "=#getting_data";
        let mut lexer = Lexer::new(formula);
        let tokens = lexer.tokenize().expect("tokenize should succeed");

        assert_eq!(tokens.len(), 2);
        assert_eq!(tokens[0].kind, TokenKind::Error(ErrorKind::GettingData));
        assert_eq!(
            &formula[tokens[0].span.start..tokens[0].span.end],
            "#getting_data"
        );
    }

    #[test]
    fn parse_spanned_formula_accepts_getting_data_error_literal() {
        let formula = "=#GETTING_DATA";
        let expr = parse_spanned_formula(formula).expect("parse should succeed");
        assert_eq!(expr.span, Span::new(1, 14));
        assert_eq!(expr.kind, SpannedExprKind::Error(ErrorKind::GettingData));
    }

    #[test]
    fn parse_spanned_formula_supports_field_access() {
        let expr = parse_spanned_formula("=A1.Price").expect("field access should parse");

        match expr.kind {
            SpannedExprKind::FieldAccess { base, field } => {
                assert_eq!(field, "Price");
                assert!(matches!(base.kind, SpannedExprKind::CellRef(_)));
            }
            other => panic!("expected FieldAccess, got {other:?}"),
        }
    }

    #[test]
    fn parse_spanned_formula_supports_bracket_field_access() {
        let expr =
            parse_spanned_formula(r#"=A1.["Change%"]"#).expect("bracket field access should parse");

        match expr.kind {
            SpannedExprKind::FieldAccess { base, field } => {
                assert_eq!(field, "Change%");
                assert!(matches!(base.kind, SpannedExprKind::CellRef(_)));
            }
            other => panic!("expected FieldAccess, got {other:?}"),
        }
    }

    #[test]
    fn parse_spanned_formula_accepts_all_excel_error_literals() {
        let cases = [
            ("#NULL!", ErrorKind::Null),
            ("#DIV/0!", ErrorKind::Div0),
            ("#VALUE!", ErrorKind::Value),
            ("#REF!", ErrorKind::Ref),
            ("#NAME?", ErrorKind::Name),
            ("#NUM!", ErrorKind::Num),
            ("#N/A", ErrorKind::NA),
            ("#GETTING_DATA", ErrorKind::GettingData),
            ("#SPILL!", ErrorKind::Spill),
            ("#CALC!", ErrorKind::Calc),
            ("#FIELD!", ErrorKind::Field),
            ("#CONNECT!", ErrorKind::Connect),
            ("#BLOCKED!", ErrorKind::Blocked),
            ("#UNKNOWN!", ErrorKind::Unknown),
        ];

        for (lit, kind) in cases {
            let formula = format!("={lit}");
            let expr = parse_spanned_formula(&formula).expect("parse should succeed");
            assert_eq!(
                expr.span,
                Span::new(1, 1 + lit.len()),
                "unexpected span for {lit}"
            );
            assert_eq!(
                expr.kind,
                SpannedExprKind::Error(kind),
                "unexpected kind for {lit}"
            );
        }
    }

    #[test]
    fn parse_spanned_formula_supports_sheet_field_access() {
        let expr = parse_spanned_formula("=Sheet1!A1.Price")
            .expect("sheet-qualified field access should parse");

        match expr.kind {
            SpannedExprKind::FieldAccess { base, field } => {
                assert_eq!(field, "Price");
                match base.kind {
                    SpannedExprKind::CellRef(cell) => {
                        assert_eq!(cell.sheet, SheetReference::Sheet("Sheet1".to_string()));
                        assert_eq!(
                            cell.addr,
                            crate::eval::Ref::from_abs_cell_addr(parse_a1("A1").unwrap()).unwrap()
                        );
                    }
                    other => panic!("expected base CellRef, got {other:?}"),
                }
            }
            other => panic!("expected FieldAccess, got {other:?}"),
        }
    }

    #[test]
    fn parse_spanned_formula_supports_sheet_bracket_field_access() {
        let expr = parse_spanned_formula(r#"=Sheet1!A1.["Change%"]"#)
            .expect("sheet-qualified bracket field access should parse");

        match expr.kind {
            SpannedExprKind::FieldAccess { base, field } => {
                assert_eq!(field, "Change%");
                match base.kind {
                    SpannedExprKind::CellRef(cell) => {
                        assert_eq!(cell.sheet, SheetReference::Sheet("Sheet1".to_string()));
                        assert_eq!(
                            cell.addr,
                            crate::eval::Ref::from_abs_cell_addr(parse_a1("A1").unwrap()).unwrap()
                        );
                    }
                    other => panic!("expected base CellRef, got {other:?}"),
                }
            }
            other => panic!("expected FieldAccess, got {other:?}"),
        }
    }

    #[test]
    fn parse_path_qualified_external_sheet_key_extracts_workbook_and_sheet() {
        let (workbook, sheet) =
            parse_path_qualified_external_sheet_key(r"C:\path\[Book.xlsx]Sheet1")
                .expect("expected path-qualified external ref");
        assert_eq!(workbook, r"C:\path\Book.xlsx");
        assert_eq!(sheet, "Sheet1");
    }

    #[test]
    fn parse_path_qualified_external_sheet_key_supports_workbook_names_containing_lbracket() {
        let (workbook, sheet) =
            parse_path_qualified_external_sheet_key(r"C:\path\[A1[Name.xlsx]Sheet1")
                .expect("expected path-qualified external ref");
        assert_eq!(workbook, r"C:\path\A1[Name.xlsx");
        assert_eq!(sheet, "Sheet1");
    }

    #[test]
    fn parse_path_qualified_external_sheet_key_supports_workbook_names_with_escaped_rbracket() {
        let (workbook, sheet) =
            parse_path_qualified_external_sheet_key(r"C:\path\[Book[Name]].xlsx]Sheet1")
                .expect("expected path-qualified external ref");
        assert_eq!(workbook, r"C:\path\Book[Name]].xlsx");
        assert_eq!(sheet, "Sheet1");
    }

    #[test]
    fn split_sheet_span_name_ignores_drive_colons_in_external_sheet_keys() {
        assert_eq!(split_sheet_span_name(r"[C:\path\Book.xlsx]Sheet1"), None);
        assert_eq!(
            split_sheet_span_name(r"[C:\path\Book.xlsx]Sheet1:Sheet3"),
            Some((
                r"[C:\path\Book.xlsx]Sheet1".to_string(),
                "Sheet3".to_string()
            ))
        );

        // The workbook id itself can contain `[` / `]` when it represents a path with bracketed
        // components (e.g. `C:\[foo]\Book.xlsx`).
        assert_eq!(split_sheet_span_name(r"[C:\[foo]\Book.xlsx]Sheet1"), None);
        assert_eq!(
            split_sheet_span_name(r"[C:\[foo]\Book.xlsx]Sheet1:Sheet3"),
            Some((
                r"[C:\[foo]\Book.xlsx]Sheet1".to_string(),
                "Sheet3".to_string()
            ))
        );
    }

    #[test]
    fn parse_workbook_scoped_external_name_ref_accepts_bracketed_paths() {
        assert_eq!(
            parse_workbook_scoped_external_name_ref(r"[C:\[foo]\Book.xlsx]MyName"),
            Some((r"C:\[foo]\Book.xlsx".to_string(), "MyName".to_string()))
        );
    }

    #[test]
    fn parse_spanned_formula_supports_path_qualified_external_workbook_refs() {
        let expr = parse_spanned_formula(r#"='C:\path\[Book.xlsx]Sheet1'!A1"#)
            .expect("parse should succeed");

        match expr.kind {
            SpannedExprKind::CellRef(cell) => {
                assert_eq!(
                    cell.sheet,
                    SheetReference::External(r"[C:\path\Book.xlsx]Sheet1".to_string())
                );
                assert_eq!(
                    cell.addr,
                    crate::eval::Ref::from_abs_cell_addr(parse_a1("A1").unwrap()).unwrap()
                );
            }
            other => panic!("expected CellRef, got {other:?}"),
        }
    }

    #[test]
    fn parse_spanned_formula_supports_path_qualified_external_refs_with_bracketed_path_components()
    {
        let expr = parse_spanned_formula(r#"='C:\[foo]\[Book.xlsx]Sheet1'!A1"#)
            .expect("parse should succeed");

        match expr.kind {
            SpannedExprKind::CellRef(cell) => {
                assert_eq!(
                    cell.sheet,
                    SheetReference::External(r"[C:\[foo]\Book.xlsx]Sheet1".to_string())
                );
                assert_eq!(
                    cell.addr,
                    crate::eval::Ref::from_abs_cell_addr(parse_a1("A1").unwrap()).unwrap()
                );
            }
            other => panic!("expected CellRef, got {other:?}"),
        }
    }

    #[test]
    fn parse_spanned_formula_supports_path_qualified_external_refs_with_lbracket_in_workbook_name()
    {
        let expr = parse_spanned_formula(r#"='C:\path\[A1[Name.xlsx]Sheet1'!A1"#)
            .expect("parse should succeed");

        match expr.kind {
            SpannedExprKind::CellRef(cell) => {
                assert_eq!(
                    cell.sheet,
                    SheetReference::External(r"[C:\path\A1[Name.xlsx]Sheet1".to_string())
                );
                assert_eq!(
                    cell.addr,
                    crate::eval::Ref::from_abs_cell_addr(parse_a1("A1").unwrap()).unwrap()
                );
            }
            other => panic!("expected CellRef, got {other:?}"),
        }
    }

    #[test]
    fn parse_spanned_formula_supports_path_qualified_external_refs_with_escaped_rbracket_in_workbook_name(
    ) {
        let expr = parse_spanned_formula(r#"='C:\path\[Book[Name]].xlsx]Sheet1'!A1"#)
            .expect("parse should succeed");

        match expr.kind {
            SpannedExprKind::CellRef(cell) => {
                assert_eq!(
                    cell.sheet,
                    SheetReference::External(r"[C:\path\Book[Name]].xlsx]Sheet1".to_string())
                );
                assert_eq!(
                    cell.addr,
                    crate::eval::Ref::from_abs_cell_addr(parse_a1("A1").unwrap()).unwrap()
                );
            }
            other => panic!("expected CellRef, got {other:?}"),
        }
    }

    #[test]
    fn parse_spanned_formula_supports_external_refs_with_escaped_rbracket_in_workbook_name() {
        let expr =
            parse_spanned_formula("=[Book]]Name.xlsx]Sheet1!A1").expect("parse should succeed");

        match expr.kind {
            SpannedExprKind::CellRef(cell) => {
                let key = match &cell.sheet {
                    SheetReference::Sheet(s) | SheetReference::External(s) => s,
                    other => panic!("expected sheet or external ref, got {other:?}"),
                };
                assert_eq!(key, "[Book]]Name.xlsx]Sheet1");
                assert_eq!(
                    cell.addr,
                    crate::eval::Ref::from_abs_cell_addr(parse_a1("A1").unwrap()).unwrap()
                );
            }
            other => panic!("expected CellRef, got {other:?}"),
        }
    }

    #[test]
    fn parse_spanned_formula_supports_external_refs_with_literal_brackets_in_workbook_name() {
        // Workbook name: `[Book]` -> workbook id: `[Book]]` (escaped `]`)
        let expr = parse_spanned_formula("=[[Book]]]Sheet1!A1").expect("parse should succeed");

        match expr.kind {
            SpannedExprKind::CellRef(cell) => {
                let key = match &cell.sheet {
                    SheetReference::Sheet(s) | SheetReference::External(s) => s,
                    other => panic!("expected sheet or external ref, got {other:?}"),
                };
                assert_eq!(key, "[[Book]]]Sheet1");
                assert_eq!(
                    cell.addr,
                    crate::eval::Ref::from_abs_cell_addr(parse_a1("A1").unwrap()).unwrap()
                );
            }
            other => panic!("expected CellRef, got {other:?}"),
        }
    }

    #[test]
    fn parse_spanned_formula_supports_unquoted_external_workbook_name_refs() {
        let expr = parse_spanned_formula("=[Book.xlsx]MyName").expect("parse should succeed");

        match expr.kind {
            SpannedExprKind::NameRef(nref) => {
                assert_eq!(
                    nref.sheet,
                    SheetReference::External("[Book.xlsx]".to_string())
                );
                assert_eq!(nref.name, "MyName");
            }
            other => panic!("expected NameRef, got {other:?}"),
        }
    }

    #[test]
    fn parse_spanned_formula_supports_quoted_external_workbook_name_refs() {
        let expr = parse_spanned_formula("='[Book.xlsx]MyName'").expect("parse should succeed");

        match expr.kind {
            SpannedExprKind::NameRef(nref) => {
                assert_eq!(
                    nref.sheet,
                    SheetReference::External("[Book.xlsx]".to_string())
                );
                assert_eq!(nref.name, "MyName");
            }
            other => panic!("expected NameRef, got {other:?}"),
        }
    }

    #[test]
    fn parse_spanned_formula_supports_external_workbook_name_refs_with_lbracket_in_workbook_name() {
        let expr = parse_spanned_formula("=[A1[Name.xlsx]MyName+1").expect("parse should succeed");

        match expr.kind {
            SpannedExprKind::Binary { op, left, right } => {
                assert_eq!(op, crate::eval::BinaryOp::Add);

                match left.kind {
                    SpannedExprKind::NameRef(nref) => {
                        assert_eq!(
                            nref.sheet,
                            SheetReference::External("[A1[Name.xlsx]".to_string())
                        );
                        assert_eq!(nref.name, "MyName");
                    }
                    other => panic!("expected NameRef, got {other:?}"),
                }

                assert_eq!(right.kind, SpannedExprKind::Number(1.0));
            }
            other => panic!("expected Binary(Add), got {other:?}"),
        }
    }

    #[test]
    fn parse_spanned_formula_supports_external_workbook_name_refs_with_escaped_rbracket_in_workbook_name(
    ) {
        let expr = parse_spanned_formula("=[Book]]Name.xlsx]MyName").expect("parse should succeed");

        match expr.kind {
            SpannedExprKind::NameRef(nref) => {
                assert_eq!(
                    nref.sheet,
                    SheetReference::External("[Book]]Name.xlsx]".to_string())
                );
                assert_eq!(nref.name, "MyName");
            }
            other => panic!("expected NameRef, got {other:?}"),
        }
    }

    #[test]
    fn parse_spanned_formula_supports_external_workbook_name_refs_with_literal_brackets_in_workbook_name(
    ) {
        // Workbook name: `[Book]` -> workbook id: `[Book]]` (escaped `]`)
        let expr = parse_spanned_formula("=[[Book]]]MyName").expect("parse should succeed");

        match expr.kind {
            SpannedExprKind::NameRef(nref) => {
                assert_eq!(
                    nref.sheet,
                    SheetReference::External("[[Book]]]".to_string())
                );
                assert_eq!(nref.name, "MyName");
            }
            other => panic!("expected NameRef, got {other:?}"),
        }
    }
}
