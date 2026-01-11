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
        write!(f, "{} (at {}..{})", self.message, self.span.start, self.span.end)
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
    Array(ArrayLiteral),
    FunctionCall(FunctionCall),
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
            Expr::NameRef(name) => name.fmt(out),
            Expr::CellRef(r) => r.fmt(out, opts)?,
            Expr::ColRef(r) => r.fmt(out, opts)?,
            Expr::RowRef(r) => r.fmt(out, opts)?,
            Expr::StructuredRef(r) => r.fmt(out)?,
            Expr::Array(arr) => arr.fmt(out, opts)?,
            Expr::FunctionCall(call) => call.fmt(out, opts)?,
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
                out.push_str(&((*index + 1).to_string()));
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
pub struct CellRef {
    pub workbook: Option<String>,
    pub sheet: Option<String>,
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
        fmt_ref_prefix(out, &self.workbook, &self.sheet);

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
                out.push_str(&(row_idx + 1).to_string());
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
    pub sheet: Option<String>,
    pub name: String,
}

impl NameRef {
    fn fmt(&self, out: &mut String) {
        fmt_ref_prefix(out, &self.workbook, &self.sheet);
        out.push_str(&self.name);
    }
}

#[derive(Debug, Clone, PartialEq, Eq, Serialize, Deserialize)]
pub struct ColRef {
    pub workbook: Option<String>,
    pub sheet: Option<String>,
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
        fmt_ref_prefix(out, &self.workbook, &self.sheet);
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
    pub sheet: Option<String>,
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
        fmt_ref_prefix(out, &self.workbook, &self.sheet);
        match opts.reference_style {
            ReferenceStyle::A1 => {
                let (row_idx, row_abs) = self.row.to_a1_index(opts.origin.map(|o| o.row))?;
                if row_abs {
                    out.push('$');
                }
                out.push_str(&(row_idx + 1).to_string());
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
    pub sheet: Option<String>,
    pub table: Option<String>,
    /// The raw specifier inside `[...]` (without the brackets).
    pub spec: String,
}

impl StructuredRef {
    fn fmt(&self, out: &mut String) -> Result<(), SerializeError> {
        fmt_ref_prefix(out, &self.workbook, &self.sheet);
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

fn fmt_sheet_name(out: &mut String, sheet: &str) {
    // Be conservative: quoting is always accepted by Excel, but the unquoted form is only valid
    // for a subset of ASCII identifier-like sheet names. This avoids producing formulas that the
    // canonical lexer cannot re-parse (e.g. non-ASCII names) and matches Excel's requirement to
    // quote sheet names that look like cell references (e.g. `A1` or `R1C1`).
    let needs_quotes = needs_sheet_quotes(sheet);
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

fn needs_sheet_quotes(sheet: &str) -> bool {
    if sheet.is_empty() {
        return true;
    }

    if is_reserved_unquoted_sheet_name(sheet) {
        return true;
    }

    if looks_like_a1_cell_reference(sheet) || looks_like_r1c1_cell_reference(sheet) {
        return true;
    }

    let mut chars = sheet.chars();
    let Some(first) = chars.next() else {
        return true;
    };
    if first.is_ascii_digit() {
        return true;
    }
    if !(first == '_' || first.is_ascii_alphabetic()) {
        return true;
    }
    if !chars.all(|ch| ch == '_' || ch == '.' || ch.is_ascii_alphanumeric()) {
        return true;
    }

    false
}

fn is_reserved_unquoted_sheet_name(sheet: &str) -> bool {
    sheet.eq_ignore_ascii_case("TRUE") || sheet.eq_ignore_ascii_case("FALSE")
}

fn looks_like_a1_cell_reference(name: &str) -> bool {
    let bytes = name.as_bytes();
    if bytes.is_empty() {
        return false;
    }

    let mut i = 0;
    while i < bytes.len() && bytes[i].is_ascii_alphabetic() {
        i += 1;
        if i > 3 {
            return false;
        }
    }

    if i == 0 || i >= bytes.len() {
        return false;
    }

    let digit_start = i;
    while i < bytes.len() && bytes[i].is_ascii_digit() {
        i += 1;
    }

    if digit_start == i || i != bytes.len() {
        return false;
    }

    // Reject impossible columns (beyond XFD). This mirrors the cheap check in
    // `formula-model` to avoid quoting names like `SHEET1` where the letters segment
    // is longer than 3 and already returned false above.
    let col = name[..digit_start]
        .chars()
        .fold(0u32, |acc, c| acc * 26 + (c.to_ascii_uppercase() as u32 - 'A' as u32 + 1));
    col <= 16_384
}

fn looks_like_r1c1_cell_reference(name: &str) -> bool {
    if name.eq_ignore_ascii_case("r") || name.eq_ignore_ascii_case("c") {
        return true;
    }

    let bytes = name.as_bytes();
    if bytes.first().copied().map(|b| b.to_ascii_uppercase()) != Some(b'R') {
        return false;
    }

    let mut i = 1;
    while i < bytes.len() && bytes[i].is_ascii_digit() {
        i += 1;
    }

    if i >= bytes.len() || bytes[i].to_ascii_uppercase() != b'C' {
        return false;
    }

    i += 1;
    while i < bytes.len() && bytes[i].is_ascii_digit() {
        i += 1;
    }

    i == bytes.len()
}

fn fmt_ref_prefix(out: &mut String, workbook: &Option<String>, sheet: &Option<String>) {
    match (workbook.as_ref(), sheet.as_ref()) {
        (Some(book), Some(sheet)) => {
            // External references are written as `[Book.xlsx]Sheet1!A1`.
            // Excel uses a single quoted string for the combined `[book]sheet` prefix when it
            // contains spaces/special characters.
            let combined = format!("[{book}]{sheet}");
            fmt_sheet_name(out, &combined);
            out.push('!');
        }
        (None, Some(sheet)) => {
            fmt_sheet_name(out, sheet);
            out.push('!');
        }
        (Some(book), None) => {
            out.push('[');
            out.push_str(book);
            out.push(']');
        }
        (None, None) => {}
    }
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
