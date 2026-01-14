use formula_model::formula_rewrite::sheet_name_eq_case_insensitive;
use serde::{Deserialize, Serialize};

/// 0-indexed cell address.
#[derive(Debug, Clone, Copy, PartialEq, Eq, Serialize, Deserialize)]
pub struct CellAddr {
    pub row: u32,
    pub col: u32,
}

impl CellAddr {
    #[must_use]
    pub fn new(row: u32, col: u32) -> Self {
        Self { row, col }
    }
}

#[derive(Debug, Clone, Copy, PartialEq, Eq, Serialize, Deserialize)]
pub struct Span {
    pub start: usize,
    pub end: usize,
}

impl Span {
    #[must_use]
    pub fn new(start: usize, end: usize) -> Self {
        Self { start, end }
    }

    #[must_use]
    pub fn add_offset(self, delta: usize) -> Self {
        Self {
            start: self.start.saturating_add(delta),
            end: self.end.saturating_add(delta),
        }
    }
}

#[derive(Debug, Clone, PartialEq, Eq, Serialize, Deserialize)]
pub struct ParseError {
    pub message: String,
    pub span: Span,
}

impl std::fmt::Display for ParseError {
    fn fmt(&self, f: &mut std::fmt::Formatter<'_>) -> std::fmt::Result {
        write!(
            f,
            "{} (at {}..{})",
            self.message, self.span.start, self.span.end
        )
    }
}

impl std::error::Error for ParseError {}

impl ParseError {
    #[must_use]
    pub fn new(message: impl Into<String>, span: Span) -> Self {
        Self {
            message: message.into(),
            span,
        }
    }

    #[must_use]
    pub fn add_offset(self, delta: usize) -> Self {
        Self {
            message: self.message,
            span: self.span.add_offset(delta),
        }
    }
}

#[derive(Debug, Clone, PartialEq, Eq, Serialize, Deserialize)]
pub struct SerializeError {
    pub message: String,
}

impl std::fmt::Display for SerializeError {
    fn fmt(&self, f: &mut std::fmt::Formatter<'_>) -> std::fmt::Result {
        write!(f, "{}", self.message)
    }
}

impl std::error::Error for SerializeError {}

impl SerializeError {
    #[must_use]
    pub fn new(message: impl Into<String>) -> Self {
        Self {
            message: message.into(),
        }
    }
}

/// Locale-specific configuration for tokenization and serialization.
#[derive(Debug, Clone, PartialEq, Eq, Serialize, Deserialize)]
pub struct LocaleConfig {
    pub decimal_separator: char,
    pub arg_separator: char,
    pub array_col_separator: char,
    pub array_row_separator: char,
    /// Thousands/grouping separator that may appear in numeric literals (e.g. `1.234,56` in
    /// `de-DE`). This is only used by the lexer to *accept* localized input; the parser does not
    /// preserve grouping separators in the token/AST representation.
    pub thousands_separator: Option<char>,
}

impl LocaleConfig {
    #[must_use]
    pub const fn en_us() -> Self {
        Self {
            decimal_separator: '.',
            arg_separator: ',',
            array_col_separator: ',',
            array_row_separator: ';',
            // The en-US thousands separator (`,`) collides with the argument separator, so we
            // disable it during lexing to avoid ambiguity.
            thousands_separator: None,
        }
    }

    #[must_use]
    pub const fn de_de() -> Self {
        Self {
            decimal_separator: ',',
            arg_separator: ';',
            array_col_separator: '\\',
            array_row_separator: ';',
            thousands_separator: Some('.'),
        }
    }

    #[must_use]
    pub const fn fr_fr() -> Self {
        // French (France) uses the same separators as `de-DE` for formulas, but commonly uses a
        // non-breaking space for thousands grouping (e.g. `1 234,56`).
        //
        // Unlike an ASCII space, NBSP does not collide with the range intersection operator
        // (which is represented by a normal space in the formula language).
        Self {
            thousands_separator: Some('\u{00A0}'),
            ..Self::de_de()
        }
    }

    #[must_use]
    pub const fn es_es() -> Self {
        // Spanish (Spain) matches the same punctuation settings as `de-DE`.
        Self::de_de()
    }

    /// Parse a number from a locale-aware string in a deterministic, Excel-like way.
    ///
    /// This is primarily intended for parsing numbers that appear in string literals, such
    /// as criteria arguments (e.g. `">1,5"` in `de-DE`).
    ///
    /// Rules:
    /// - Accept both the locale decimal separator *and* the canonical `.` decimal separator.
    /// - Strip the locale thousands separator when present.
    /// - Return `None` if the input cannot be interpreted as a number after normalization.
    pub fn parse_number(&self, raw: &str) -> Option<f64> {
        let raw = raw.trim();
        if raw.is_empty() {
            return None;
        }

        let (mantissa, exponent) = split_numeric_exponent(raw);

        let (sign, mantissa) = match mantissa.as_bytes().first().copied() {
            Some(b'+' | b'-') => (Some(mantissa.as_bytes()[0] as char), &mantissa[1..]),
            _ => (None, mantissa),
        };

        if mantissa.is_empty() {
            return None;
        }

        let mut decimal = if mantissa.contains(self.decimal_separator) {
            Some(self.decimal_separator)
        } else if mantissa.contains('.') {
            Some('.')
        } else {
            None
        };

        // Disambiguate locales where the thousands separator collides with the canonical decimal
        // separator (e.g. `de-DE` uses `.` for grouping and `,` for decimals).
        //
        // If the input only contains `.` and it matches a typical thousands grouping pattern
        // (`1.234.567`), treat `.` as thousands separators rather than a decimal point.
        if decimal == Some('.')
            && self.decimal_separator != '.'
            && self.thousands_separator == Some('.')
            && looks_like_thousands_grouping(mantissa, '.')
        {
            decimal = None;
        }

        let mut out = String::with_capacity(raw.len());
        if let Some(sign) = sign {
            out.push(sign);
        }

        let mut decimal_used = false;
        for ch in mantissa.chars() {
            if ch.is_ascii_digit() {
                out.push(ch);
                continue;
            }

            if Some(ch) == decimal {
                if decimal_used {
                    return None;
                }
                out.push('.');
                decimal_used = true;
                continue;
            }

            // Some locales (notably fr-FR) commonly use NBSP (U+00A0) for thousands grouping, but
            // narrow NBSP (U+202F) also appears in spreadsheets. When configured for either,
            // accept both.
            let is_thousands_sep = Some(ch) == self.thousands_separator
                || (self.thousands_separator == Some('\u{00A0}') && ch == '\u{202F}')
                || (self.thousands_separator == Some('\u{202F}') && ch == '\u{00A0}');
            if is_thousands_sep && Some(ch) != decimal {
                // Strip locale grouping separators.
                continue;
            }

            // Reject any other character (including locale separators that don't apply to this input).
            return None;
        }

        out.push_str(exponent);
        out.parse::<f64>().ok()
    }
}

fn split_numeric_exponent(raw: &str) -> (&str, &str) {
    // Fast path: no exponent marker.
    if !raw.as_bytes().iter().any(|b| matches!(b, b'e' | b'E')) {
        return (raw, "");
    }

    // Find the first valid exponent marker (`e`/`E` followed by `[+-]?digits`).
    for (idx, ch) in raw.char_indices() {
        if !matches!(ch, 'e' | 'E') {
            continue;
        }
        let rest = &raw[idx + ch.len_utf8()..];
        let rest = rest
            .strip_prefix('+')
            .or_else(|| rest.strip_prefix('-'))
            .unwrap_or(rest);
        if rest.is_empty() || !rest.chars().all(|c| c.is_ascii_digit()) {
            continue;
        }
        return (&raw[..idx], &raw[idx..]);
    }

    (raw, "")
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

#[derive(Debug, Clone, Copy, PartialEq, Eq, Serialize, Deserialize)]
pub enum ReferenceStyle {
    A1,
    R1C1,
}

#[derive(Debug, Clone, PartialEq, Eq, Serialize, Deserialize)]
pub struct ParseOptions {
    pub locale: LocaleConfig,
    pub reference_style: ReferenceStyle,
    /// If provided, normalize relative A1 references into row/col offsets.
    pub normalize_relative_to: Option<CellAddr>,
}

impl Default for ParseOptions {
    fn default() -> Self {
        Self {
            locale: LocaleConfig::en_us(),
            reference_style: ReferenceStyle::A1,
            normalize_relative_to: None,
        }
    }
}

#[derive(Debug, Clone, PartialEq, Eq, Serialize, Deserialize)]
pub struct SerializeOptions {
    pub locale: LocaleConfig,
    pub reference_style: ReferenceStyle,
    /// When `true`, serialize functions with `_xlfn.` prefix if present.
    pub include_xlfn_prefix: bool,
    /// Origin cell used to render relative offsets.
    pub origin: Option<CellAddr>,
    /// When `true`, omit a leading `=` in the output.
    pub omit_equals: bool,
}

impl Default for SerializeOptions {
    fn default() -> Self {
        Self {
            locale: LocaleConfig::en_us(),
            reference_style: ReferenceStyle::A1,
            include_xlfn_prefix: false,
            origin: None,
            omit_equals: false,
        }
    }
}

#[derive(Debug, Clone, PartialEq, Eq, Serialize, Deserialize)]
pub struct Ast {
    pub has_equals: bool,
    pub expr: Expr,
}

impl Ast {
    #[must_use]
    pub fn new(has_equals: bool, expr: Expr) -> Self {
        Self { has_equals, expr }
    }

    /// Normalize relative A1 references into row/col offsets relative to `origin`.
    #[must_use]
    pub fn normalize_relative(&self, origin: CellAddr) -> Self {
        Self {
            has_equals: self.has_equals,
            expr: self.expr.normalize_relative(origin),
        }
    }

    /// Serialize the AST back into a formula string.
    pub fn to_string(&self, opts: SerializeOptions) -> Result<String, SerializeError> {
        let mut out = String::new();
        if !opts.omit_equals && self.has_equals {
            out.push('=');
        }
        self.expr.fmt(&mut out, &opts, None)?;
        Ok(out)
    }

    /// Stable JSON serialization useful for debugging/tests.
    #[must_use]
    pub fn to_json(&self) -> String {
        serde_json::to_string(self).expect("Ast should be JSON-serializable")
    }
}

#[derive(Debug, Clone, PartialEq, Eq, Serialize, Deserialize)]
pub enum Expr {
    Number(String),
    String(String),
    Boolean(bool),
    Error(String),
    NameRef(NameRef),
    CellRef(CellRef),
    ColRef(ColRef),
    RowRef(RowRef),
    StructuredRef(StructuredRef),
    /// Field access on a base expression, e.g. `A1.Price` or `A1.["Market Cap"]`.
    ///
    /// The `field` string is the unescaped field name (it does not include the leading `.` or any
    /// surrounding brackets/quotes).
    FieldAccess(FieldAccessExpr),
    Array(ArrayLiteral),
    FunctionCall(FunctionCall),
    Call(CallExpr),
    Unary(UnaryExpr),
    Postfix(PostfixExpr),
    Binary(BinaryExpr),
    /// Missing expression/argument (used by partial parsing and empty args like `IF(,1,2)`).
    Missing,
}

impl Expr {
    fn normalize_relative(&self, origin: CellAddr) -> Self {
        match self {
            Expr::CellRef(r) => Expr::CellRef(r.normalize_relative(origin)),
            Expr::ColRef(r) => Expr::ColRef(r.normalize_relative(origin)),
            Expr::RowRef(r) => Expr::RowRef(r.normalize_relative(origin)),
            Expr::StructuredRef(r) => Expr::StructuredRef(r.clone()),
            Expr::FieldAccess(access) => Expr::FieldAccess(FieldAccessExpr {
                base: Box::new(access.base.normalize_relative(origin)),
                field: access.field.clone(),
            }),
            Expr::Number(v) => Expr::Number(v.clone()),
            Expr::String(v) => Expr::String(v.clone()),
            Expr::Boolean(v) => Expr::Boolean(*v),
            Expr::Error(v) => Expr::Error(v.clone()),
            Expr::NameRef(v) => Expr::NameRef(v.clone()),
            Expr::Array(arr) => Expr::Array(arr.clone()),
            Expr::FunctionCall(call) => Expr::FunctionCall(FunctionCall {
                name: call.name.clone(),
                args: call
                    .args
                    .iter()
                    .map(|e| e.normalize_relative(origin))
                    .collect(),
            }),
            Expr::Call(call) => Expr::Call(CallExpr {
                callee: Box::new(call.callee.normalize_relative(origin)),
                args: call
                    .args
                    .iter()
                    .map(|e| e.normalize_relative(origin))
                    .collect(),
            }),
            Expr::Unary(u) => Expr::Unary(UnaryExpr {
                op: u.op,
                expr: Box::new(u.expr.normalize_relative(origin)),
            }),
            Expr::Postfix(p) => Expr::Postfix(PostfixExpr {
                op: p.op,
                expr: Box::new(p.expr.normalize_relative(origin)),
            }),
            Expr::Binary(b) => Expr::Binary(BinaryExpr {
                op: b.op,
                left: Box::new(b.left.normalize_relative(origin)),
                right: Box::new(b.right.normalize_relative(origin)),
            }),
            Expr::Missing => Expr::Missing,
        }
    }

    fn precedence(&self) -> u8 {
        match self {
            Expr::Binary(b) => b.op.precedence(),
            Expr::Unary(_) => 70,
            Expr::Postfix(_) => 60,
            Expr::Call(_) => 90,
            Expr::FieldAccess(_) => 90,
            _ => 100,
        }
    }

    fn fmt(
        &self,
        out: &mut String,
        opts: &SerializeOptions,
        parent_prec: Option<u8>,
    ) -> Result<(), SerializeError> {
        let my_prec = self.precedence();
        let needs_parens = parent_prec.is_some_and(|p| my_prec < p);
        if needs_parens {
            out.push('(');
        }
        match self {
            Expr::Number(raw) => out.push_str(raw),
            Expr::String(value) => {
                out.push('"');
                for ch in value.chars() {
                    if ch == '"' {
                        out.push('"');
                        out.push('"');
                    } else {
                        out.push(ch);
                    }
                }
                out.push('"');
            }
            Expr::Boolean(v) => out.push_str(if *v { "TRUE" } else { "FALSE" }),
            Expr::Error(v) => out.push_str(v),
            Expr::NameRef(name) => name.fmt(out, opts),
            Expr::CellRef(r) => r.fmt(out, opts)?,
            Expr::ColRef(r) => r.fmt(out, opts)?,
            Expr::RowRef(r) => r.fmt(out, opts)?,
            Expr::StructuredRef(r) => r.fmt(out, opts)?,
            Expr::Array(arr) => arr.fmt(out, opts)?,
            Expr::FunctionCall(call) => call.fmt(out, opts)?,
            Expr::Call(call) => {
                call.callee.fmt(out, opts, Some(my_prec))?;
                out.push('(');
                for (i, arg) in call.args.iter().enumerate() {
                    if i > 0 {
                        out.push(opts.locale.arg_separator);
                    }
                    // See `FunctionCall::fmt` for rationale: union uses the locale list separator,
                    // which must be disambiguated from argument separators with parentheses.
                    if arg.contains_union() {
                        out.push('(');
                        arg.fmt(out, opts, None)?;
                        out.push(')');
                    } else {
                        arg.fmt(out, opts, None)?;
                    }
                }
                out.push(')');
            }
            Expr::FieldAccess(fa) => {
                fa.base.fmt(out, opts, Some(my_prec))?;
                out.push('.');
                if is_field_ident_safe(&fa.field) {
                    out.push_str(&fa.field);
                } else {
                    out.push('[');
                    out.push('"');
                    for ch in fa.field.chars() {
                        if ch == '"' {
                            out.push('"');
                            out.push('"');
                        } else {
                            out.push(ch);
                        }
                    }
                    out.push('"');
                    out.push(']');
                }
            }
            Expr::Unary(u) => {
                out.push_str(u.op.as_str());
                u.expr.fmt(out, opts, Some(my_prec))?;
            }
            Expr::Postfix(p) => {
                p.expr.fmt(out, opts, Some(my_prec))?;
                out.push_str(p.op.as_str());
            }
            Expr::Binary(b) => {
                b.left.fmt(out, opts, Some(my_prec))?;
                out.push_str(&b.op.as_str(opts));
                b.right.fmt(out, opts, Some(my_prec))?;
            }
            Expr::Missing => {}
        }
        if needs_parens {
            out.push(')');
        }
        Ok(())
    }

    fn contains_union(&self) -> bool {
        match self {
            Expr::Binary(b) => {
                b.op == BinaryOp::Union || b.left.contains_union() || b.right.contains_union()
            }
            Expr::Unary(u) => u.expr.contains_union(),
            Expr::Postfix(p) => p.expr.contains_union(),
            Expr::FunctionCall(call) => call.args.iter().any(Expr::contains_union),
            Expr::Call(call) => {
                call.callee.contains_union() || call.args.iter().any(Expr::contains_union)
            }
            Expr::FieldAccess(fa) => fa.base.contains_union(),
            Expr::Array(arr) => arr.rows.iter().flatten().any(Expr::contains_union),
            _ => false,
        }
    }
}

#[derive(Debug, Clone, Copy, PartialEq, Eq, Serialize, Deserialize)]
pub enum UnaryOp {
    Plus,
    Minus,
    ImplicitIntersection,
}

impl UnaryOp {
    fn as_str(self) -> &'static str {
        match self {
            UnaryOp::Plus => "+",
            UnaryOp::Minus => "-",
            UnaryOp::ImplicitIntersection => "@",
        }
    }
}

#[derive(Debug, Clone, PartialEq, Eq, Serialize, Deserialize)]
pub struct UnaryExpr {
    pub op: UnaryOp,
    pub expr: Box<Expr>,
}

#[derive(Debug, Clone, Copy, PartialEq, Eq, Serialize, Deserialize)]
pub enum PostfixOp {
    Percent,
    /// Excel spill-range reference operator (`#`), e.g. `A1#`.
    SpillRange,
}

impl PostfixOp {
    fn as_str(self) -> &'static str {
        match self {
            PostfixOp::Percent => "%",
            PostfixOp::SpillRange => "#",
        }
    }
}

#[derive(Debug, Clone, PartialEq, Eq, Serialize, Deserialize)]
pub struct PostfixExpr {
    pub op: PostfixOp,
    pub expr: Box<Expr>,
}

#[derive(Debug, Clone, Copy, PartialEq, Eq, Serialize, Deserialize)]
pub enum BinaryOp {
    Range,
    Intersect,
    Union,
    Pow,
    Mul,
    Div,
    Add,
    Sub,
    Concat,
    Eq,
    Ne,
    Lt,
    Gt,
    Le,
    Ge,
}

impl BinaryOp {
    fn precedence(self) -> u8 {
        match self {
            BinaryOp::Range => 82,
            BinaryOp::Intersect => 81,
            BinaryOp::Union => 80,
            BinaryOp::Pow => 50,
            BinaryOp::Mul | BinaryOp::Div => 40,
            BinaryOp::Add | BinaryOp::Sub => 30,
            BinaryOp::Concat => 20,
            BinaryOp::Eq
            | BinaryOp::Ne
            | BinaryOp::Lt
            | BinaryOp::Gt
            | BinaryOp::Le
            | BinaryOp::Ge => 10,
        }
    }

    fn as_str(self, opts: &SerializeOptions) -> String {
        match self {
            BinaryOp::Range => ":".to_string(),
            BinaryOp::Intersect => " ".to_string(),
            BinaryOp::Union => opts.locale.arg_separator.to_string(),
            BinaryOp::Pow => "^".to_string(),
            BinaryOp::Mul => "*".to_string(),
            BinaryOp::Div => "/".to_string(),
            BinaryOp::Add => "+".to_string(),
            BinaryOp::Sub => "-".to_string(),
            BinaryOp::Concat => "&".to_string(),
            BinaryOp::Eq => "=".to_string(),
            BinaryOp::Ne => "<>".to_string(),
            BinaryOp::Lt => "<".to_string(),
            BinaryOp::Gt => ">".to_string(),
            BinaryOp::Le => "<=".to_string(),
            BinaryOp::Ge => ">=".to_string(),
        }
    }
}

#[derive(Debug, Clone, PartialEq, Eq, Serialize, Deserialize)]
pub struct BinaryExpr {
    pub op: BinaryOp,
    pub left: Box<Expr>,
    pub right: Box<Expr>,
}

#[derive(Debug, Clone, PartialEq, Eq, Serialize, Deserialize)]
pub struct FunctionName {
    pub original: String,
    pub name_upper: String,
    pub has_xlfn_prefix: bool,
}

impl FunctionName {
    #[must_use]
    pub fn new(original: String) -> Self {
        let upper = original.to_ascii_uppercase();
        let (has_prefix, base_upper) = upper
            .strip_prefix("_XLFN.")
            .map(|s| (true, s.to_string()))
            .unwrap_or((false, upper.clone()));

        Self {
            original,
            name_upper: base_upper,
            has_xlfn_prefix: has_prefix,
        }
    }

    fn fmt(&self, out: &mut String, opts: &SerializeOptions) {
        if self.has_xlfn_prefix && opts.include_xlfn_prefix {
            out.push_str("_xlfn.");
        }
        out.push_str(&self.name_upper);
    }
}

#[derive(Debug, Clone, PartialEq, Eq, Serialize, Deserialize)]
pub struct FunctionCall {
    pub name: FunctionName,
    pub args: Vec<Expr>,
}

impl FunctionCall {
    fn fmt(&self, out: &mut String, opts: &SerializeOptions) -> Result<(), SerializeError> {
        self.name.fmt(out, opts);
        out.push('(');
        for (i, arg) in self.args.iter().enumerate() {
            if i > 0 {
                out.push(opts.locale.arg_separator);
            }
            // The union operator uses the locale list separator, which is also the function
            // argument separator. Excel disambiguates union inside arguments by requiring
            // explicit parentheses (e.g. `SUM((A1,B1))`).
            //
            // For round-trip safety, wrap any argument expression that contains a union
            // operator in parentheses so it re-parses as an expression rather than being
            // split into multiple args.
            if arg.contains_union() {
                out.push('(');
                arg.fmt(out, opts, None)?;
                out.push(')');
            } else {
                arg.fmt(out, opts, None)?;
            }
        }
        out.push(')');
        Ok(())
    }
}

#[derive(Debug, Clone, PartialEq, Eq, Serialize, Deserialize)]
pub struct CallExpr {
    pub callee: Box<Expr>,
    pub args: Vec<Expr>,
}

#[derive(Debug, Clone, PartialEq, Eq, Serialize, Deserialize)]
pub struct FieldAccessExpr {
    pub base: Box<Expr>,
    pub field: String,
}

#[derive(Debug, Clone, PartialEq, Eq, Serialize, Deserialize)]
pub enum Coord {
    /// A coordinate as written in A1 notation (0-indexed) with a `$` marker flag.
    A1 { index: u32, abs: bool },
    /// A relative offset from an origin cell.
    Offset(i32),
}

impl Coord {
    fn normalize(index: u32, abs: bool, origin: u32) -> Self {
        if abs {
            Coord::A1 { index, abs }
        } else {
            Coord::Offset(index as i32 - origin as i32)
        }
    }

    fn to_a1_index(&self, origin: Option<u32>) -> Result<(u32, bool), SerializeError> {
        match self {
            Coord::A1 { index, abs } => Ok((*index, *abs)),
            Coord::Offset(delta) => {
                let origin = origin.ok_or_else(|| {
                    SerializeError::new("Cannot render relative offset without an origin cell")
                })?;
                let idx = origin as i32 + *delta;
                if idx < 0 {
                    return Err(SerializeError::new("Relative reference moved before A1"));
                }
                Ok((idx as u32, false))
            }
        }
    }

    fn fmt_r1c1(
        &self,
        out: &mut String,
        axis: char,
        origin: Option<u32>,
    ) -> Result<(), SerializeError> {
        match self {
            Coord::A1 { index, abs: true } => {
                out.push(axis);
                out.push_str(&(u64::from(*index) + 1).to_string());
            }
            Coord::A1 { index, abs: false } => {
                let origin = origin.ok_or_else(|| {
                    SerializeError::new("Cannot render relative reference without an origin cell")
                })?;
                let delta = *index as i32 - origin as i32;
                out.push(axis);
                if delta != 0 {
                    out.push('[');
                    out.push_str(&delta.to_string());
                    out.push(']');
                }
            }
            Coord::Offset(delta) => {
                out.push(axis);
                if *delta != 0 {
                    out.push('[');
                    out.push_str(&delta.to_string());
                    out.push(']');
                }
            }
        }
        Ok(())
    }
}

#[derive(Debug, Clone, PartialEq, Eq, Serialize, Deserialize)]
pub enum SheetRef {
    Sheet(String),
    /// 3D sheet span reference like `Sheet1:Sheet3!A1`.
    SheetRange {
        start: String,
        end: String,
    },
}

impl SheetRef {
    pub fn as_single_sheet(&self) -> Option<&str> {
        match self {
            SheetRef::Sheet(name) => Some(name),
            SheetRef::SheetRange { start, end } if sheet_name_eq_case_insensitive(start, end) => {
                Some(start)
            }
            SheetRef::SheetRange { .. } => None,
        }
    }
}

#[derive(Debug, Clone, PartialEq, Eq, Serialize, Deserialize)]
pub struct CellRef {
    pub workbook: Option<String>,
    pub sheet: Option<SheetRef>,
    pub col: Coord,
    pub row: Coord,
}

impl CellRef {
    fn normalize_relative(&self, origin: CellAddr) -> Self {
        let (col_index, col_abs) = match self.col {
            Coord::A1 { index, abs } => (index, abs),
            Coord::Offset(_) => return self.clone(),
        };
        let (row_index, row_abs) = match self.row {
            Coord::A1 { index, abs } => (index, abs),
            Coord::Offset(_) => return self.clone(),
        };

        Self {
            workbook: self.workbook.clone(),
            sheet: self.sheet.clone(),
            col: Coord::normalize(col_index, col_abs, origin.col),
            row: Coord::normalize(row_index, row_abs, origin.row),
        }
    }

    fn fmt(&self, out: &mut String, opts: &SerializeOptions) -> Result<(), SerializeError> {
        fmt_ref_prefix(out, &self.workbook, &self.sheet, opts.reference_style);

        match opts.reference_style {
            ReferenceStyle::A1 => {
                let origin = opts.origin;
                let (col_idx, col_abs) = self.col.to_a1_index(origin.map(|o| o.col))?;
                let (row_idx, row_abs) = self.row.to_a1_index(origin.map(|o| o.row))?;

                if col_abs {
                    out.push('$');
                }
                out.push_str(&col_to_a1(col_idx));
                if row_abs {
                    out.push('$');
                }
                out.push_str(&(u64::from(row_idx) + 1).to_string());
            }
            ReferenceStyle::R1C1 => {
                self.row.fmt_r1c1(out, 'R', opts.origin.map(|o| o.row))?;
                self.col.fmt_r1c1(out, 'C', opts.origin.map(|o| o.col))?;
            }
        }
        Ok(())
    }
}

#[derive(Debug, Clone, PartialEq, Eq, Serialize, Deserialize)]
pub struct NameRef {
    pub workbook: Option<String>,
    pub sheet: Option<SheetRef>,
    pub name: String,
}

impl NameRef {
    fn fmt(&self, out: &mut String, opts: &SerializeOptions) {
        match (self.workbook.as_ref(), self.sheet.as_ref()) {
            (Some(book), None) => {
                // Workbook-scoped external names are written in Excel as `[Book.xlsx]MyName`, but
                // this is ambiguous with structured references for our parser/lexer. Emit the
                // combined token as a quoted identifier (`'[Book.xlsx]MyName'`) so it round-trips.
                //
                // If the workbook id is path-qualified (e.g. `C:\path\Book.xlsx`), prefer the
                // Excel-canonical form `'C:\path\[Book.xlsx]MyName'`.
                if let Some(sep) = book.rfind(['\\', '/']) {
                    let (prefix, base) = book.split_at(sep + 1);
                    if !base.is_empty() {
                        out.push('\'');
                        fmt_sheet_name_escaped(out, prefix);
                        out.push('[');
                        fmt_sheet_name_escaped(out, base);
                        out.push(']');
                        fmt_sheet_name_escaped(out, &self.name);
                        out.push('\'');
                        return;
                    }
                }

                out.push('\'');
                out.push('[');
                fmt_sheet_name_escaped(out, book);
                out.push(']');
                fmt_sheet_name_escaped(out, &self.name);
                out.push('\'');
            }
            _ => {
                fmt_ref_prefix(out, &self.workbook, &self.sheet, opts.reference_style);
                out.push_str(&self.name);
            }
        }
    }
}

#[derive(Debug, Clone, PartialEq, Eq, Serialize, Deserialize)]
pub struct ColRef {
    pub workbook: Option<String>,
    pub sheet: Option<SheetRef>,
    pub col: Coord,
}

impl ColRef {
    fn normalize_relative(&self, origin: CellAddr) -> Self {
        let (col_index, col_abs) = match self.col {
            Coord::A1 { index, abs } => (index, abs),
            Coord::Offset(_) => return self.clone(),
        };
        Self {
            workbook: self.workbook.clone(),
            sheet: self.sheet.clone(),
            col: Coord::normalize(col_index, col_abs, origin.col),
        }
    }

    fn fmt(&self, out: &mut String, opts: &SerializeOptions) -> Result<(), SerializeError> {
        fmt_ref_prefix(out, &self.workbook, &self.sheet, opts.reference_style);
        match opts.reference_style {
            ReferenceStyle::A1 => {
                let (col_idx, col_abs) = self.col.to_a1_index(opts.origin.map(|o| o.col))?;
                if col_abs {
                    out.push('$');
                }
                out.push_str(&col_to_a1(col_idx));
            }
            ReferenceStyle::R1C1 => {
                self.col.fmt_r1c1(out, 'C', opts.origin.map(|o| o.col))?;
            }
        }
        Ok(())
    }
}

#[derive(Debug, Clone, PartialEq, Eq, Serialize, Deserialize)]
pub struct RowRef {
    pub workbook: Option<String>,
    pub sheet: Option<SheetRef>,
    pub row: Coord,
}

impl RowRef {
    fn normalize_relative(&self, origin: CellAddr) -> Self {
        let (row_index, row_abs) = match self.row {
            Coord::A1 { index, abs } => (index, abs),
            Coord::Offset(_) => return self.clone(),
        };
        Self {
            workbook: self.workbook.clone(),
            sheet: self.sheet.clone(),
            row: Coord::normalize(row_index, row_abs, origin.row),
        }
    }

    fn fmt(&self, out: &mut String, opts: &SerializeOptions) -> Result<(), SerializeError> {
        fmt_ref_prefix(out, &self.workbook, &self.sheet, opts.reference_style);
        match opts.reference_style {
            ReferenceStyle::A1 => {
                let (row_idx, row_abs) = self.row.to_a1_index(opts.origin.map(|o| o.row))?;
                if row_abs {
                    out.push('$');
                }
                out.push_str(&(u64::from(row_idx) + 1).to_string());
            }
            ReferenceStyle::R1C1 => {
                self.row.fmt_r1c1(out, 'R', opts.origin.map(|o| o.row))?;
            }
        }
        Ok(())
    }
}

#[derive(Debug, Clone, PartialEq, Eq, Serialize, Deserialize)]
pub struct StructuredRef {
    pub workbook: Option<String>,
    pub sheet: Option<SheetRef>,
    pub table: Option<String>,
    /// The raw specifier inside `[...]` (without the brackets).
    pub spec: String,
}

impl StructuredRef {
    fn fmt(&self, out: &mut String, opts: &SerializeOptions) -> Result<(), SerializeError> {
        fmt_ref_prefix(out, &self.workbook, &self.sheet, opts.reference_style);
        if let Some(table) = &self.table {
            out.push_str(table);
        }
        out.push('[');
        out.push_str(&self.spec);
        out.push(']');
        Ok(())
    }
}

#[derive(Debug, Clone, PartialEq, Eq, Serialize, Deserialize)]
pub struct ArrayLiteral {
    pub rows: Vec<Vec<Expr>>,
}

impl ArrayLiteral {
    fn fmt(&self, out: &mut String, opts: &SerializeOptions) -> Result<(), SerializeError> {
        out.push('{');
        for (r, row) in self.rows.iter().enumerate() {
            if r > 0 {
                out.push(opts.locale.array_row_separator);
            }
            for (c, el) in row.iter().enumerate() {
                if c > 0 {
                    out.push(opts.locale.array_col_separator);
                }
                el.fmt(out, opts, None)?;
            }
        }
        out.push('}');
        Ok(())
    }
}

fn fmt_sheet_name(out: &mut String, sheet: &str, reference_style: ReferenceStyle) {
    // Be conservative: quoting is always accepted by Excel, but the unquoted form is only valid
    // for a subset of identifier-like sheet names. This avoids producing formulas that the
    // canonical lexer cannot re-parse (e.g. non-ASCII names) and matches Excel's requirement to
    // quote sheet names that look like cell references (e.g. `A1` or `R1C1`).
    let needs_quotes = sheet_name_needs_quotes(sheet, reference_style);
    if needs_quotes {
        out.push('\'');
        for ch in sheet.chars() {
            if ch == '\'' {
                out.push('\'');
                out.push('\'');
            } else {
                out.push(ch);
            }
        }
        out.push('\'');
    } else {
        out.push_str(sheet);
    }
}

fn fmt_sheet_name_escaped(out: &mut String, sheet: &str) {
    for ch in sheet.chars() {
        if ch == '\'' {
            out.push('\'');
            out.push('\'');
        } else {
            out.push(ch);
        }
    }
}

fn fmt_sheet_range_name(out: &mut String, start: &str, end: &str, reference_style: ReferenceStyle) {
    let needs_quotes = sheet_name_needs_quotes(start, reference_style)
        || sheet_name_needs_quotes(end, reference_style);
    if needs_quotes {
        out.push('\'');
        fmt_sheet_name_escaped(out, start);
        out.push(':');
        fmt_sheet_name_escaped(out, end);
        out.push('\'');
    } else {
        out.push_str(start);
        out.push(':');
        out.push_str(end);
    }
}

fn fmt_ref_prefix(
    out: &mut String,
    workbook: &Option<String>,
    sheet: &Option<SheetRef>,
    reference_style: ReferenceStyle,
) {
    fn split_path_prefix(book: &str) -> Option<(&str, &str)> {
        // Path-qualified external workbook ids arise when parsing formulas like
        // `'C:\path\[Book.xlsx]Sheet1'!A1`: the parser folds the path prefix (`C:\path\`) into the
        // workbook id, producing `C:\path\Book.xlsx`.
        //
        // When serializing back to formula text, keep the path prefix *outside* the `[workbook]`
        // bracket so any `[` / `]` characters in directory names (e.g. `C:\[foo]\`) do not need to
        // be escaped as workbook-prefix `]]` sequences.
        let sep = book.rfind(['\\', '/'])?;
        let (prefix, base) = book.split_at(sep + 1);
        if base.is_empty() {
            None
        } else {
            Some((prefix, base))
        }
    }

    match (workbook.as_ref(), sheet.as_ref()) {
        (Some(book), Some(sheet_ref)) => {
            // External references are written as `[Book.xlsx]Sheet1!A1`.
            //
            // If `book` is a path-qualified workbook id (e.g. `C:\path\Book.xlsx`), prefer the
            // Excel-canonical form where the path prefix appears *outside* the workbook brackets:
            // `'C:\path\[Book.xlsx]Sheet1'!A1`.
            //
            // This is important for paths containing bracket characters in directory names (e.g.
            // `C:\[foo]\`): embedding the full path inside `[ ... ]` would introduce an unescaped
            // `]` inside the workbook prefix, making the serialized formula unparseable.
            if let Some((path_prefix, base)) = split_path_prefix(book) {
                out.push('\'');
                fmt_sheet_name_escaped(out, path_prefix);
                out.push('[');
                fmt_sheet_name_escaped(out, base);
                out.push(']');
                match sheet_ref {
                    SheetRef::Sheet(sheet) => {
                        fmt_sheet_name_escaped(out, sheet);
                    }
                    SheetRef::SheetRange { start, end } => {
                        if sheet_name_eq_case_insensitive(start, end) {
                            fmt_sheet_name_escaped(out, start);
                        } else {
                            fmt_sheet_name_escaped(out, start);
                            out.push(':');
                            fmt_sheet_name_escaped(out, end);
                        }
                    }
                }
                out.push('\'');
                out.push('!');
                return;
            }

            // Excel uses a single quoted string for the combined `[book]sheet` prefix when the
            // sheet name requires quoting (e.g. spaces / characters that aren't valid identifiers).
            match sheet_ref {
                SheetRef::Sheet(sheet) => {
                    // Workbook names inside `[...]` are permissive (Excel allows spaces, dashes,
                    // etc), but the sheet name portion follows normal quoting rules.
                    //
                    // However, when the workbook id contains `[`/`]` (e.g. from a path prefix that
                    // includes brackets), we must also force a quoted combined token so the lexer
                    // doesn't interpret nested brackets as structured-ref syntax.
                    let needs_quotes = sheet_name_needs_quotes(sheet, reference_style)
                        || book.contains('[')
                        || book.contains(']');
                    if needs_quotes {
                        out.push('\'');
                        out.push('[');
                        fmt_sheet_name_escaped(out, book);
                        out.push(']');
                        fmt_sheet_name_escaped(out, sheet);
                        out.push('\'');
                    } else {
                        out.push('[');
                        out.push_str(book);
                        out.push(']');
                        out.push_str(sheet);
                    }
                    out.push('!');
                }
                SheetRef::SheetRange { start, end } => {
                    if sheet_name_eq_case_insensitive(start, end) {
                        // Degenerate 3D span within an external workbook.
                        let needs_quotes = sheet_name_needs_quotes(start, reference_style)
                            || book.contains('[')
                            || book.contains(']');
                        if needs_quotes {
                            out.push('\'');
                            out.push('[');
                            fmt_sheet_name_escaped(out, book);
                            out.push(']');
                            fmt_sheet_name_escaped(out, start);
                            out.push('\'');
                        } else {
                            out.push('[');
                            out.push_str(book);
                            out.push(']');
                            out.push_str(start);
                        }
                        out.push('!');
                    } else {
                        let needs_quotes = book.contains('[')
                            || book.contains(']')
                            || sheet_name_needs_quotes(start, reference_style)
                            || sheet_name_needs_quotes(end, reference_style);
                        if needs_quotes {
                            out.push('\'');
                            out.push('[');
                            fmt_sheet_name_escaped(out, book);
                            out.push(']');
                            fmt_sheet_name_escaped(out, start);
                            out.push(':');
                            fmt_sheet_name_escaped(out, end);
                            out.push('\'');
                        } else {
                            out.push('[');
                            out.push_str(book);
                            out.push(']');
                            out.push_str(start);
                            out.push(':');
                            out.push_str(end);
                        }
                        out.push('!');
                    }
                }
            }
        }
        (None, Some(sheet_ref)) => match sheet_ref {
            SheetRef::Sheet(sheet) => {
                fmt_sheet_name(out, sheet, reference_style);
                out.push('!');
            }
            SheetRef::SheetRange { start, end } => {
                if sheet_name_eq_case_insensitive(start, end) {
                    fmt_sheet_name(out, start, reference_style);
                } else {
                    fmt_sheet_range_name(out, start, end, reference_style);
                }
                out.push('!');
            }
        },
        (Some(book), None) => {
            out.push('[');
            out.push_str(book);
            out.push(']');
        }
        (None, None) => {}
    };
}

fn sheet_name_needs_quotes(sheet: &str, reference_style: ReferenceStyle) -> bool {
    if sheet.is_empty() {
        return true;
    }
    if sheet
        .chars()
        .any(|c| c.is_whitespace() || matches!(c, '!' | '\''))
    {
        return true;
    }

    sheet_part_needs_quotes(sheet, reference_style)
}

fn sheet_part_needs_quotes(sheet: &str, reference_style: ReferenceStyle) -> bool {
    debug_assert!(!sheet.is_empty());

    if sheet.eq_ignore_ascii_case("TRUE") || sheet.eq_ignore_ascii_case("FALSE") {
        return true;
    }

    // A1-style cell references are tokenized as `Cell(...)` even when followed by additional
    // identifier characters, so treat any sheet name starting with a valid cell reference as
    // requiring quotes (e.g. `'A1B'!C1`).
    if starts_like_a1_cell_ref(sheet) {
        return true;
    }

    if reference_style == ReferenceStyle::R1C1 && starts_like_r1c1_ref(sheet) {
        return true;
    }

    if !is_valid_ident(sheet) {
        return true;
    }

    false
}

fn is_valid_ident(ident: &str) -> bool {
    let mut chars = ident.chars();
    let Some(first) = chars.next() else {
        return false;
    };
    if !matches!(first, '$' | '_' | '\\' | 'A'..='Z' | 'a'..='z') {
        return false;
    }
    chars.all(is_ident_cont_char)
}

fn is_field_ident_safe(field: &str) -> bool {
    let mut chars = field.chars();
    let Some(first) = chars.next() else {
        return false;
    };

    // The lexer treats TRUE/FALSE as booleans rather than identifiers.
    if field.eq_ignore_ascii_case("TRUE") || field.eq_ignore_ascii_case("FALSE") {
        return false;
    }

    let valid_start = matches!(first, '$' | '_' | '\\' | 'A'..='Z' | 'a'..='z')
        || (!first.is_ascii() && first.is_alphabetic());
    if !valid_start {
        return false;
    }

    chars.all(|c| {
        c != '.'
            && (matches!(c, '$' | '_' | '\\' | 'A'..='Z' | 'a'..='z' | '0'..='9')
                || (!c.is_ascii() && c.is_alphanumeric()))
    })
}

fn is_ident_cont_char(c: char) -> bool {
    matches!(
        c,
        '$' | '_' | '\\' | '.' | 'A'..='Z' | 'a'..='z' | '0'..='9'
    )
}

fn starts_like_a1_cell_ref(s: &str) -> bool {
    let mut chars = s.chars().peekable();
    if chars.peek() == Some(&'$') {
        chars.next();
    }

    let mut col_letters = String::new();
    while let Some(&ch) = chars.peek() {
        if ch.is_ascii_alphabetic() {
            col_letters.push(ch);
            chars.next();
        } else {
            break;
        }
    }
    if col_letters.is_empty() {
        return false;
    }

    if chars.peek() == Some(&'$') {
        chars.next();
    }

    let mut row_digits = String::new();
    while let Some(&ch) = chars.peek() {
        if ch.is_ascii_digit() {
            row_digits.push(ch);
            chars.next();
        } else {
            break;
        }
    }
    if row_digits.is_empty() {
        return false;
    }

    if col_from_a1(&col_letters).is_none() {
        return false;
    }
    matches!(row_digits.parse::<u32>(), Ok(v) if v != 0)
}

fn starts_like_r1c1_ref(s: &str) -> bool {
    starts_like_r1c1_cell_ref(s) || starts_like_r1c1_row_ref(s) || starts_like_r1c1_col_ref(s)
}

fn starts_like_r1c1_cell_ref(s: &str) -> bool {
    let mut chars = s.chars().peekable();
    let Some(first) = chars.next() else {
        return false;
    };
    if !matches!(first, 'R' | 'r') {
        return false;
    }

    // Row part (optional).
    if matches!(chars.peek(), Some(c) if c.is_ascii_digit()) {
        let digits: String = chars.by_ref().take_while(|c| c.is_ascii_digit()).collect();
        if matches!(digits.parse::<u32>(), Ok(v) if v == 0) || digits.parse::<u32>().is_err() {
            return false;
        }
    }

    let Some(ch) = chars.next() else {
        return false;
    };
    if !matches!(ch, 'C' | 'c') {
        return false;
    }

    // Col part (optional).
    if matches!(chars.peek(), Some(c) if c.is_ascii_digit()) {
        let digits: String = chars.by_ref().take_while(|c| c.is_ascii_digit()).collect();
        if matches!(digits.parse::<u32>(), Ok(v) if v == 0) || digits.parse::<u32>().is_err() {
            return false;
        }
    }

    // A valid R1C1 cell reference is accepted as a prefix; any remaining characters would form
    // additional tokens, so treat this as requiring quotes.
    true
}

fn starts_like_r1c1_row_ref(s: &str) -> bool {
    let mut chars = s.chars().peekable();
    let Some(first) = chars.next() else {
        return false;
    };
    if !matches!(first, 'R' | 'r') {
        return false;
    }

    if matches!(chars.peek(), Some(c) if c.is_ascii_digit()) {
        let digits: String = chars.by_ref().take_while(|c| c.is_ascii_digit()).collect();
        if matches!(digits.parse::<u32>(), Ok(v) if v == 0) || digits.parse::<u32>().is_err() {
            return false;
        }
    }

    // Matches the lexer guard: treat as a row ref only if the next character does *not* continue
    // an identifier.
    !matches!(chars.peek(), Some(c) if is_ident_cont_char(*c) || *c == '(')
}

fn starts_like_r1c1_col_ref(s: &str) -> bool {
    let mut chars = s.chars().peekable();
    let Some(first) = chars.next() else {
        return false;
    };
    if !matches!(first, 'C' | 'c') {
        return false;
    }

    if matches!(chars.peek(), Some(c) if c.is_ascii_digit()) {
        let digits: String = chars.by_ref().take_while(|c| c.is_ascii_digit()).collect();
        if matches!(digits.parse::<u32>(), Ok(v) if v == 0) || digits.parse::<u32>().is_err() {
            return false;
        }
    }

    !matches!(chars.peek(), Some(c) if is_ident_cont_char(*c) || *c == '(')
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

fn col_to_a1(mut col: u32) -> String {
    // Excel-style base-26 letters.
    let mut out = String::new();
    loop {
        let rem = (col % 26) as u8;
        out.push((b'A' + rem) as char);
        col /= 26;
        if col == 0 {
            break;
        }
        col -= 1;
    }
    out.chars().rev().collect()
}
