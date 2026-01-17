use super::value::{CellCoord, ErrorKind, MultiRangeRef, RangeRef, Ref, Value};
use formula_model::column_label_to_index_lenient;
use std::sync::Arc;

#[derive(Clone, Copy, Debug, PartialEq, Eq, Hash)]
pub enum UnaryOp {
    Plus,
    Neg,
    ImplicitIntersection,
}

#[derive(Clone, Copy, Debug, PartialEq, Eq, Hash)]
pub enum BinaryOp {
    Add,
    Sub,
    Mul,
    Div,
    Pow,
    /// Reference union (`,`).
    Union,
    /// Reference intersection (whitespace).
    Intersect,
    Eq,
    Ne,
    Lt,
    Le,
    Gt,
    Ge,
}

#[derive(Clone, Debug, PartialEq, Eq, Hash)]
pub enum Function {
    /// Internal synthetic function used by lowering field-access expressions (`A1.Price`).
    ///
    /// This is not an Excel worksheet function; it is emitted by the canonical AST lowering
    /// pipeline as an implementation detail of the `.` operator.
    FieldAccess,
    Let,
    IsOmitted,
    True,
    False,
    If,
    Choose,
    Ifs,
    And,
    Or,
    Xor,
    IfError,
    IfNa,
    IsError,
    IsNa,
    IsBlank,
    IsNumber,
    IsText,
    IsLogical,
    IsErr,
    Type,
    ErrorType,
    N,
    T,
    Na,
    Switch,
    Sum,
    SumIf,
    SumIfs,
    Average,
    AverageIf,
    AverageIfs,
    Min,
    MinIfs,
    Max,
    MaxIfs,
    Count,
    CountA,
    CountBlank,
    CountIf,
    CountIfs,
    SumProduct,
    VLookup,
    HLookup,
    Match,
    Abs,
    Int,
    Round,
    RoundUp,
    RoundDown,
    Mod,
    Sign,
    Db,
    Vdb,
    CoupDayBs,
    CoupDays,
    CoupDaysNc,
    CoupNcd,
    CoupNum,
    CoupPcd,
    Price,
    Yield,
    Duration,
    MDuration,
    Accrint,
    Accrintm,
    Disc,
    PriceDisc,
    YieldDisc,
    Intrate,
    Received,
    PriceMat,
    YieldMat,
    TbillEq,
    TbillPrice,
    TbillYield,
    OddFPrice,
    OddFYield,
    OddLPrice,
    OddLYield,
    /// Variadic concatenation operator (`&`) lowered by the engine parser/lowerer.
    ///
    /// This uses Excel's elementwise/broadcasting semantics and can spill arrays.
    ConcatOp,
    /// `CONCAT(...)` function (flattens ranges/arrays into a single string).
    Concat,
    /// `CONCATENATE(...)` function (scalar-only; applies implicit intersection on references).
    Concatenate,
    Rand,
    RandBetween,
    Not,
    Now,
    Today,
    Row,
    Column,
    Rows,
    Columns,
    Address,
    Offset,
    Indirect,
    XLookup,
    XMatch,
    Unknown(Arc<str>),
}

impl Function {
    pub fn from_name(name: &str) -> Self {
        crate::value::with_ascii_uppercased_key(name, |upper| Self::from_uppercase_name(upper))
    }

    fn from_uppercase_name(upper: &str) -> Self {
        let base = upper.strip_prefix("_XLFN.").unwrap_or(upper);
        match base {
            "_FIELDACCESS" => Function::FieldAccess,
            "LET" => Function::Let,
            "ISOMITTED" => Function::IsOmitted,
            "TRUE" => Function::True,
            "FALSE" => Function::False,
            "IF" => Function::If,
            "CHOOSE" => Function::Choose,
            "IFS" => Function::Ifs,
            "AND" => Function::And,
            "OR" => Function::Or,
            "XOR" => Function::Xor,
            "IFERROR" => Function::IfError,
            "IFNA" => Function::IfNa,
            "ISERROR" => Function::IsError,
            "ISNA" => Function::IsNa,
            "ISBLANK" => Function::IsBlank,
            "ISNUMBER" => Function::IsNumber,
            "ISTEXT" => Function::IsText,
            "ISLOGICAL" => Function::IsLogical,
            "ISERR" => Function::IsErr,
            "TYPE" => Function::Type,
            "ERROR.TYPE" => Function::ErrorType,
            "N" => Function::N,
            "T" => Function::T,
            "NA" => Function::Na,
            "SWITCH" => Function::Switch,
            "SUM" => Function::Sum,
            "SUMIF" => Function::SumIf,
            "SUMIFS" => Function::SumIfs,
            "AVERAGE" => Function::Average,
            "AVERAGEIF" => Function::AverageIf,
            "AVERAGEIFS" => Function::AverageIfs,
            "MIN" => Function::Min,
            "MINIFS" => Function::MinIfs,
            "MAX" => Function::Max,
            "MAXIFS" => Function::MaxIfs,
            "COUNT" => Function::Count,
            "COUNTA" => Function::CountA,
            "COUNTBLANK" => Function::CountBlank,
            "COUNTIF" => Function::CountIf,
            "COUNTIFS" => Function::CountIfs,
            "SUMPRODUCT" => Function::SumProduct,
            "VLOOKUP" => Function::VLookup,
            "HLOOKUP" => Function::HLookup,
            "MATCH" => Function::Match,
            "XLOOKUP" => Function::XLookup,
            "XMATCH" => Function::XMatch,
            "ABS" => Function::Abs,
            "INT" => Function::Int,
            "ROUND" => Function::Round,
            "ROUNDUP" => Function::RoundUp,
            "ROUNDDOWN" => Function::RoundDown,
            "MOD" => Function::Mod,
            "SIGN" => Function::Sign,
            "DB" => Function::Db,
            "VDB" => Function::Vdb,
            "COUPDAYBS" => Function::CoupDayBs,
            "COUPDAYS" => Function::CoupDays,
            "COUPDAYSNC" => Function::CoupDaysNc,
            "COUPNCD" => Function::CoupNcd,
            "COUPNUM" => Function::CoupNum,
            "COUPPCD" => Function::CoupPcd,
            "PRICE" => Function::Price,
            "YIELD" => Function::Yield,
            "DURATION" => Function::Duration,
            "MDURATION" => Function::MDuration,
            "ACCRINT" => Function::Accrint,
            "ACCRINTM" => Function::Accrintm,
            "DISC" => Function::Disc,
            "PRICEDISC" => Function::PriceDisc,
            "YIELDDISC" => Function::YieldDisc,
            "INTRATE" => Function::Intrate,
            "RECEIVED" => Function::Received,
            "PRICEMAT" => Function::PriceMat,
            "YIELDMAT" => Function::YieldMat,
            "TBILLEQ" => Function::TbillEq,
            "TBILLPRICE" => Function::TbillPrice,
            "TBILLYIELD" => Function::TbillYield,
            "ODDFPRICE" => Function::OddFPrice,
            "ODDFYIELD" => Function::OddFYield,
            "ODDLPRICE" => Function::OddLPrice,
            "ODDLYIELD" => Function::OddLYield,
            "CONCAT" => Function::Concat,
            "CONCATENATE" => Function::Concatenate,
            "RAND" => Function::Rand,
            "RANDBETWEEN" => Function::RandBetween,
            "NOT" => Function::Not,
            "NOW" => Function::Now,
            "TODAY" => Function::Today,
            "ROW" => Function::Row,
            "COLUMN" => Function::Column,
            "ROWS" => Function::Rows,
            "COLUMNS" => Function::Columns,
            "ADDRESS" => Function::Address,
            "OFFSET" => Function::Offset,
            "INDIRECT" => Function::Indirect,
            other => Function::Unknown(Arc::from(other)),
        }
    }

    pub fn name(&self) -> &str {
        match self {
            Function::FieldAccess => "_FIELDACCESS",
            Function::Let => "LET",
            Function::IsOmitted => "ISOMITTED",
            Function::True => "TRUE",
            Function::False => "FALSE",
            Function::If => "IF",
            Function::Choose => "CHOOSE",
            Function::Ifs => "IFS",
            Function::And => "AND",
            Function::Or => "OR",
            Function::Xor => "XOR",
            Function::IfError => "IFERROR",
            Function::IfNa => "IFNA",
            Function::IsError => "ISERROR",
            Function::IsNa => "ISNA",
            Function::Na => "NA",
            Function::Switch => "SWITCH",
            Function::Sum => "SUM",
            Function::SumIf => "SUMIF",
            Function::SumIfs => "SUMIFS",
            Function::Average => "AVERAGE",
            Function::AverageIf => "AVERAGEIF",
            Function::AverageIfs => "AVERAGEIFS",
            Function::Min => "MIN",
            Function::MinIfs => "MINIFS",
            Function::Max => "MAX",
            Function::MaxIfs => "MAXIFS",
            Function::Count => "COUNT",
            Function::CountA => "COUNTA",
            Function::CountBlank => "COUNTBLANK",
            Function::CountIf => "COUNTIF",
            Function::CountIfs => "COUNTIFS",
            Function::SumProduct => "SUMPRODUCT",
            Function::VLookup => "VLOOKUP",
            Function::HLookup => "HLOOKUP",
            Function::Match => "MATCH",
            Function::Abs => "ABS",
            Function::Int => "INT",
            Function::Round => "ROUND",
            Function::RoundUp => "ROUNDUP",
            Function::RoundDown => "ROUNDDOWN",
            Function::Mod => "MOD",
            Function::Sign => "SIGN",
            Function::Db => "DB",
            Function::Vdb => "VDB",
            Function::CoupDayBs => "COUPDAYBS",
            Function::CoupDays => "COUPDAYS",
            Function::CoupDaysNc => "COUPDAYSNC",
            Function::CoupNcd => "COUPNCD",
            Function::CoupNum => "COUPNUM",
            Function::CoupPcd => "COUPPCD",
            Function::Price => "PRICE",
            Function::Yield => "YIELD",
            Function::Duration => "DURATION",
            Function::MDuration => "MDURATION",
            Function::Accrint => "ACCRINT",
            Function::Accrintm => "ACCRINTM",
            Function::Disc => "DISC",
            Function::PriceDisc => "PRICEDISC",
            Function::YieldDisc => "YIELDDISC",
            Function::Intrate => "INTRATE",
            Function::Received => "RECEIVED",
            Function::PriceMat => "PRICEMAT",
            Function::YieldMat => "YIELDMAT",
            Function::TbillEq => "TBILLEQ",
            Function::TbillPrice => "TBILLPRICE",
            Function::TbillYield => "TBILLYIELD",
            Function::OddFPrice => "ODDFPRICE",
            Function::OddFYield => "ODDFYIELD",
            Function::OddLPrice => "ODDLPRICE",
            Function::OddLYield => "ODDLYIELD",
            Function::ConcatOp => "CONCAT_OP",
            Function::Concat => "CONCAT",
            Function::Concatenate => "CONCATENATE",
            Function::Rand => "RAND",
            Function::RandBetween => "RANDBETWEEN",
            Function::Not => "NOT",
            Function::IsBlank => "ISBLANK",
            Function::IsNumber => "ISNUMBER",
            Function::IsText => "ISTEXT",
            Function::IsLogical => "ISLOGICAL",
            Function::IsErr => "ISERR",
            Function::Type => "TYPE",
            Function::ErrorType => "ERROR.TYPE",
            Function::N => "N",
            Function::T => "T",
            Function::Now => "NOW",
            Function::Today => "TODAY",
            Function::Row => "ROW",
            Function::Column => "COLUMN",
            Function::Rows => "ROWS",
            Function::Columns => "COLUMNS",
            Function::Address => "ADDRESS",
            Function::Offset => "OFFSET",
            Function::Indirect => "INDIRECT",
            Function::XLookup => "XLOOKUP",
            Function::XMatch => "XMATCH",
            Function::Unknown(s) => s,
        }
    }
}

#[derive(Clone, Debug, PartialEq, Eq)]
pub enum Expr {
    Literal(Value),
    CellRef(Ref),
    RangeRef(RangeRef),
    MultiRangeRef(MultiRangeRef),
    /// Lexical name reference (LET/LAMBDA bindings). Names are stored case-folded.
    NameRef(Arc<str>),
    /// Excel spill-range reference operator (`#`), e.g. `A1#`.
    ///
    /// The inner expression is evaluated in a "reference context" (cell references are lowered
    /// to single-cell ranges). At runtime this resolves the spill origin and expands to the full
    /// spill footprint as a `Value::Range`.
    SpillRange(Box<Expr>),
    Unary {
        op: UnaryOp,
        expr: Box<Expr>,
    },
    Binary {
        op: BinaryOp,
        left: Box<Expr>,
        right: Box<Expr>,
    },
    FuncCall {
        func: Function,
        args: Vec<Expr>,
    },
    /// LAMBDA(param1, ..., body) lowered into explicit params + body.
    Lambda {
        params: Arc<[Arc<str>]>,
        body: Box<Expr>,
    },
    /// Postfix call expression: `callee(args...)`.
    Call {
        callee: Box<Expr>,
        args: Vec<Expr>,
    },
}

#[derive(thiserror::Error, Debug, Clone, PartialEq, Eq)]
pub enum ParseError {
    #[error("unexpected end of input")]
    UnexpectedEof,
    #[error("unexpected token at byte {0}")]
    UnexpectedToken(usize),
    #[error("Too many arguments (max {0})")]
    TooManyArguments(usize),
    #[error("invalid number")]
    InvalidNumber,
    #[error("invalid cell reference")]
    InvalidCellRef,
    #[error("unterminated string literal")]
    UnterminatedString,
}

pub fn parse_formula(formula: &str, origin: CellCoord) -> Result<Expr, ParseError> {
    Parser::new(formula, origin).parse()
}

#[derive(Clone, Copy, Debug, PartialEq, Eq)]
enum InfixOp {
    Binary(BinaryOp),
    Concat,
}

impl InfixOp {
    fn token_len(self) -> usize {
        match self {
            InfixOp::Binary(op) => op.token_len(),
            InfixOp::Concat => 1,
        }
    }
}

struct Parser<'a> {
    input: &'a [u8],
    pos: usize,
    origin: CellCoord,
}

impl<'a> Parser<'a> {
    fn new(formula: &'a str, origin: CellCoord) -> Self {
        Self {
            input: formula.as_bytes(),
            pos: 0,
            origin,
        }
    }

    fn parse(mut self) -> Result<Expr, ParseError> {
        self.skip_ws();
        if self.peek_byte() == Some(b'=') {
            self.pos += 1;
        }
        let expr = self.parse_bp(0, false)?;
        self.skip_ws();
        if self.pos != self.input.len() {
            return Err(ParseError::UnexpectedToken(self.pos));
        }
        Ok(expr)
    }

    fn parse_bp(&mut self, min_bp: u8, stop_on_comma: bool) -> Result<Expr, ParseError> {
        self.skip_ws();
        let mut lhs = self.parse_prefix(stop_on_comma)?;
        loop {
            let had_ws = self.skip_ws_report();

            // Postfix percent operator (`expr%`).
            //
            // This lowers directly to `expr / 100` so we can reuse existing numeric coercion
            // semantics in `BinaryOp::Div` (including error propagation and spill behavior for
            // range/array-as-scalar cases).
            // Postfix operators bind tighter than any infix operator.
            let postfix_bp = 20;
            if self.peek_byte() == Some(b'%') && postfix_bp >= min_bp {
                self.pos += 1;
                lhs = Expr::Binary {
                    op: BinaryOp::Div,
                    left: Box::new(lhs),
                    right: Box::new(Expr::Literal(Value::Number(100.0))),
                };
                continue;
            }

            // Postfix spill-range operator (`expr#`).
            //
            // Like the full engine parser, treat this as a postfix operator that binds tighter
            // than exponentiation. The bytecode runtime expects the operand to be evaluated in a
            // "reference context", so lower a direct cell reference to a single-cell range here.
            if self.peek_byte() == Some(b'#') && postfix_bp >= min_bp {
                self.pos += 1;
                lhs = match lhs {
                    Expr::CellRef(r) => {
                        Expr::SpillRange(Box::new(Expr::RangeRef(RangeRef::new(r, r))))
                    }
                    other => Expr::SpillRange(Box::new(other)),
                };
                continue;
            }

            // Postfix call expression: `callee(args...)`.
            //
            // This is used for lambda invocation syntax (e.g. `LAMBDA(x, x+1)(3)` or `f(3)` where
            // `f` is a LET-bound lambda).
            if self.peek_byte() == Some(b'(') && postfix_bp >= min_bp {
                self.pos += 1;
                let args = self.parse_parenthesized_args()?;
                lhs = Expr::Call {
                    callee: Box::new(lhs),
                    args,
                };
                continue;
            }

            let op_pos = self.pos;
            let (op, l_bp, r_bp) = match self.peek_infix_op(stop_on_comma) {
                Some(v) => v,
                None => {
                    // Reference intersection operator (whitespace) is only syntactically visible
                    // after whitespace has been consumed. If there was whitespace between two
                    // expressions and we're not immediately followed by an infix operator or
                    // delimiter, interpret it as the intersection operator.
                    let Some(next) = self.peek_byte() else {
                        break;
                    };
                    let starts_expr = matches!(
                        next,
                        b'+' | b'-' | b'@' | b'(' | b'"' | b'#' | b'.' | b'0'..=b'9'
                            | b'A'..=b'Z'
                            | b'a'..=b'z'
                            | b'_' | b'$'
                    );
                    if had_ws && starts_expr {
                        let (l_bp, r_bp) = (12, 13);
                        if l_bp < min_bp {
                            break;
                        }
                        let rhs = self.parse_bp(r_bp, stop_on_comma)?;
                        lhs = Expr::Binary {
                            op: BinaryOp::Intersect,
                            left: Box::new(lhs),
                            right: Box::new(rhs),
                        };
                        continue;
                    }
                    break;
                }
            };
            if l_bp < min_bp {
                break;
            }
            self.pos = op_pos + op.token_len();
            let rhs = self.parse_bp(r_bp, stop_on_comma)?;
            lhs = match op {
                InfixOp::Binary(op) => Expr::Binary {
                    op,
                    left: Box::new(lhs),
                    right: Box::new(rhs),
                },
                InfixOp::Concat => {
                    // Flatten concat chains (`a&b&c`) into a single CONCAT_OP call.
                    let mut args = match lhs {
                        Expr::FuncCall {
                            func: Function::ConcatOp,
                            args,
                        } => args,
                        other => vec![other],
                    };
                    match rhs {
                        Expr::FuncCall {
                            func: Function::ConcatOp,
                            args: rhs_args,
                        } => args.extend(rhs_args),
                        other => args.push(other),
                    }
                    Expr::FuncCall {
                        func: Function::ConcatOp,
                        args,
                    }
                }
            };
        }
        Ok(lhs)
    }

    fn parse_prefix(&mut self, stop_on_comma: bool) -> Result<Expr, ParseError> {
        self.skip_ws();
        match self.peek_byte() {
            Some(b'+') => {
                self.pos += 1;
                Ok(Expr::Unary {
                    op: UnaryOp::Plus,
                    expr: Box::new(self.parse_bp(9, stop_on_comma)?),
                })
            }
            Some(b'-') => {
                self.pos += 1;
                Ok(Expr::Unary {
                    op: UnaryOp::Neg,
                    expr: Box::new(self.parse_bp(9, stop_on_comma)?),
                })
            }
            Some(b'@') => {
                self.pos += 1;
                Ok(Expr::Unary {
                    op: UnaryOp::ImplicitIntersection,
                    expr: Box::new(self.parse_bp(9, stop_on_comma)?),
                })
            }
            Some(b'(') => {
                self.pos += 1;
                // Parenthesized expressions are not function-argument lists, so commas act as the
                // reference union operator (Excel requires parentheses to disambiguate e.g.
                // `SUM((A1,B1))`).
                let expr = self.parse_bp(0, false)?;
                self.skip_ws();
                if self.peek_byte() != Some(b')') {
                    return Err(ParseError::UnexpectedToken(self.pos));
                }
                self.pos += 1;
                Ok(expr)
            }
            Some(b'"') => self.parse_string(),
            Some(b'0'..=b'9') | Some(b'.') => self.parse_number(),
            Some(b'#') => self.parse_error_literal(),
            Some(_) => self.parse_ident_like(),
            None => Err(ParseError::UnexpectedEof),
        }
    }

    fn parse_error_literal(&mut self) -> Result<Expr, ParseError> {
        debug_assert_eq!(self.peek_byte(), Some(b'#'));
        let start = self.pos;
        self.pos += 1; // '#'
        while let Some(b) = self.peek_byte() {
            if matches!(b, b'_' | b'/' | b'.' | b'0'..=b'9' | b'A'..=b'Z' | b'a'..=b'z') {
                self.pos += 1;
            } else {
                break;
            }
        }

        // Optional `!` / `?` suffix (e.g. `#REF!`, `#NAME?`).
        if matches!(self.peek_byte(), Some(b'!' | b'?')) {
            self.pos += 1;
        }

        if self.pos == start + 1 {
            return Err(ParseError::UnexpectedToken(start));
        }

        let s = std::str::from_utf8(&self.input[start..self.pos])
            .map_err(|_| ParseError::UnexpectedToken(start))?;
        let kind = ErrorKind::from_code(s).ok_or(ParseError::UnexpectedToken(start))?;
        Ok(Expr::Literal(Value::Error(kind)))
    }

    fn parse_number(&mut self) -> Result<Expr, ParseError> {
        let start = self.pos;
        let mut end = self.pos;
        let mut seen_dot = false;
        let mut seen_exp = false;

        while let Some(b) = self.input.get(end).copied() {
            match b {
                b'0'..=b'9' => end += 1,
                b'.' if !seen_dot && !seen_exp => {
                    seen_dot = true;
                    end += 1;
                }
                b'e' | b'E' if !seen_exp => {
                    seen_exp = true;
                    end += 1;
                    if let Some(sign) = self.input.get(end).copied() {
                        if sign == b'+' || sign == b'-' {
                            end += 1;
                        }
                    }
                }
                _ => break,
            }
        }

        self.pos = end;
        let s =
            std::str::from_utf8(&self.input[start..end]).map_err(|_| ParseError::InvalidNumber)?;
        let v: f64 = s.parse().map_err(|_| ParseError::InvalidNumber)?;
        Ok(Expr::Literal(Value::Number(v)))
    }

    fn parse_string(&mut self) -> Result<Expr, ParseError> {
        debug_assert_eq!(self.peek_byte(), Some(b'"'));
        self.pos += 1;
        // Keep bytes intact so UTF-8 content is preserved. We only interpret `"` (ASCII) for
        // termination / `""` escaping.
        let mut out: Vec<u8> = Vec::new();
        while let Some(b) = self.peek_byte() {
            self.pos += 1;
            match b {
                b'"' => {
                    // Excel escapes quotes by doubling them.
                    if self.peek_byte() == Some(b'"') {
                        self.pos += 1;
                        out.push(b'"');
                    } else {
                        let text = String::from_utf8(out)
                            .expect("string literal bytes come from valid UTF-8 input");
                        return Ok(Expr::Literal(Value::Text(Arc::from(text))));
                    }
                }
                _ => out.push(b),
            }
        }
        Err(ParseError::UnterminatedString)
    }

    /// Parse an argument list after consuming the opening `(`.
    ///
    /// This is shared between function calls (`SUM(...)`) and call expressions (`expr(...)`).
    fn parse_parenthesized_args(&mut self) -> Result<Vec<Expr>, ParseError> {
        let mut args = Vec::new();
        self.skip_ws();
        if self.peek_byte() != Some(b')') {
            loop {
                if args.len() == crate::EXCEL_MAX_ARGS {
                    return Err(ParseError::TooManyArguments(crate::EXCEL_MAX_ARGS));
                }

                // Excel allows omitted/missing arguments (e.g. `ADDRESS(1,1,,FALSE)` or
                // `IF(FALSE,1,)`). Treat an empty slot between separators as
                // `Expr::Literal(Value::Missing)` so the runtime can distinguish omitted args
                // from blank cell values for functions with optional-argument semantics.
                self.skip_ws();
                match self.peek_byte() {
                    Some(b',') => {
                        args.push(Expr::Literal(Value::Missing));
                        self.pos += 1; // consume comma and continue to the next argument
                        continue;
                    }
                    Some(b')') => {
                        // Trailing comma: implicit missing argument at the end.
                        args.push(Expr::Literal(Value::Missing));
                        break;
                    }
                    Some(_) => args.push(self.parse_bp(0, true)?),
                    None => return Err(ParseError::UnexpectedEof),
                }

                self.skip_ws();
                match self.peek_byte() {
                    Some(b',') => {
                        self.pos += 1;
                        continue;
                    }
                    Some(b')') => break,
                    _ => return Err(ParseError::UnexpectedToken(self.pos)),
                }
            }
        }

        if self.peek_byte() != Some(b')') {
            return Err(ParseError::UnexpectedToken(self.pos));
        }
        self.pos += 1;
        Ok(args)
    }

    fn parse_ident_like(&mut self) -> Result<Expr, ParseError> {
        let start = self.pos;
        while let Some(b) = self.peek_byte() {
            match b {
                b'A'..=b'Z' | b'a'..=b'z' | b'0'..=b'9' | b'_' | b'$' | b'.' => self.pos += 1,
                _ => break,
            }
        }
        if self.pos == start {
            return Err(ParseError::UnexpectedToken(self.pos));
        }
        let ident = std::str::from_utf8(&self.input[start..self.pos])
            .map_err(|_| ParseError::UnexpectedToken(start))?;
        self.skip_ws();

        if self.peek_byte() == Some(b'(') {
            self.pos += 1;
            let mut args = self.parse_parenthesized_args()?;
            let base = if ident.len() >= 6 && ident[..6].eq_ignore_ascii_case("_xlfn.") {
                &ident[6..]
            } else {
                ident
            };
            if base.eq_ignore_ascii_case("LAMBDA") {
                if args.is_empty() {
                    return Err(ParseError::UnexpectedToken(start));
                }
                let body = Box::new(args.pop().expect("checked len"));
                let mut params = Vec::with_capacity(args.len());
                for arg in args {
                    match arg {
                        Expr::NameRef(name) => params.push(name),
                        _ => return Err(ParseError::UnexpectedToken(start)),
                    }
                }
                return Ok(Expr::Lambda {
                    params: Arc::from(params.into_boxed_slice()),
                    body,
                });
            }

            let func = Function::from_name(ident);
            // If the name is not a known built-in function, interpret `name(args...)` as a call
            // expression on a lexical name reference (used for LET/LAMBDA invocation syntax).
            if matches!(func, Function::Unknown(_)) {
                let callee = crate::value::with_ascii_uppercased_key(ident, |upper| Arc::from(upper));
                return Ok(Expr::Call {
                    callee: Box::new(Expr::NameRef(callee)),
                    args,
                });
            }

            return Ok(Expr::FuncCall { func, args });
        }

        if ident.eq_ignore_ascii_case("TRUE") {
            return Ok(Expr::Literal(Value::Bool(true)));
        }
        if ident.eq_ignore_ascii_case("FALSE") {
            return Ok(Expr::Literal(Value::Bool(false)));
        }

        let Some(first) = parse_a1_ref(ident, self.origin) else {
            let name = crate::value::with_ascii_uppercased_key(ident, |upper| Arc::from(upper));
            return Ok(Expr::NameRef(name));
        };
        self.skip_ws();
        if self.peek_byte() == Some(b':') {
            self.pos += 1;
            self.skip_ws();
            let start2 = self.pos;
            while let Some(b) = self.peek_byte() {
                match b {
                    b'A'..=b'Z' | b'a'..=b'z' | b'0'..=b'9' | b'_' | b'$' | b'.' => self.pos += 1,
                    _ => break,
                }
            }
            let ident2 = std::str::from_utf8(&self.input[start2..self.pos])
                .map_err(|_| ParseError::UnexpectedToken(start2))?;
            let second = parse_a1_ref(ident2, self.origin).ok_or(ParseError::InvalidCellRef)?;
            Ok(Expr::RangeRef(RangeRef::new(first, second)))
        } else {
            Ok(Expr::CellRef(first))
        }
    }

    fn peek_infix_op(&self, stop_on_comma: bool) -> Option<(InfixOp, u8, u8)> {
        let b0 = *self.input.get(self.pos)?;
        let b1 = *self.input.get(self.pos + 1).unwrap_or(&0);
        let (op, l_bp, r_bp) = match (b0, b1) {
            (b',', _) if !stop_on_comma => (InfixOp::Binary(BinaryOp::Union), 11, 12),
            (b'+', _) => (InfixOp::Binary(BinaryOp::Add), 5, 6),
            (b'-', _) => (InfixOp::Binary(BinaryOp::Sub), 5, 6),
            (b'*', _) => (InfixOp::Binary(BinaryOp::Mul), 7, 8),
            (b'/', _) => (InfixOp::Binary(BinaryOp::Div), 7, 8),
            (b'^', _) => (InfixOp::Binary(BinaryOp::Pow), 9, 9), // right associative
            (b'&', _) => (InfixOp::Concat, 4, 5),
            (b'=', _) => (InfixOp::Binary(BinaryOp::Eq), 3, 4),
            (b'<', b'>') => (InfixOp::Binary(BinaryOp::Ne), 3, 4),
            (b'<', b'=') => (InfixOp::Binary(BinaryOp::Le), 3, 4),
            (b'>', b'=') => (InfixOp::Binary(BinaryOp::Ge), 3, 4),
            (b'<', _) => (InfixOp::Binary(BinaryOp::Lt), 3, 4),
            (b'>', _) => (InfixOp::Binary(BinaryOp::Gt), 3, 4),
            _ => return None,
        };
        Some((op, l_bp, r_bp))
    }

    fn skip_ws(&mut self) {
        while let Some(b) = self.peek_byte() {
            if b == b' ' || b == b'\t' || b == b'\n' || b == b'\r' {
                self.pos += 1;
            } else {
                break;
            }
        }
    }

    fn skip_ws_report(&mut self) -> bool {
        let start = self.pos;
        self.skip_ws();
        self.pos != start
    }

    fn peek_byte(&self) -> Option<u8> {
        self.input.get(self.pos).copied()
    }
}

impl BinaryOp {
    fn token_len(self) -> usize {
        match self {
            BinaryOp::Ne | BinaryOp::Le | BinaryOp::Ge => 2,
            _ => 1,
        }
    }
}

fn parse_a1_ref(token: &str, origin: CellCoord) -> Option<Ref> {
    // Format: [$]COL[$]ROW (A1, $A$1, A$1, $A1)
    let mut bytes = token.as_bytes();

    let mut col_abs = false;
    if bytes.first() == Some(&b'$') {
        col_abs = true;
        bytes = &bytes[1..];
    }

    let mut col_end = 0;
    while col_end < bytes.len() {
        let b = bytes[col_end];
        if (b'A'..=b'Z').contains(&b) || (b'a'..=b'z').contains(&b) {
            col_end += 1;
        } else {
            break;
        }
    }
    if col_end == 0 {
        return None;
    }
    let col_letters = std::str::from_utf8(&bytes[..col_end]).ok()?;
    let mut row_bytes = &bytes[col_end..];

    let mut row_abs = false;
    if row_bytes.first() == Some(&b'$') {
        row_abs = true;
        row_bytes = &row_bytes[1..];
    }

    if row_bytes.is_empty() {
        return None;
    }

    let row_str = std::str::from_utf8(row_bytes).ok()?;
    let row_1: i32 = row_str.parse().ok()?;
    if row_1 <= 0 {
        return None;
    }
    let row0 = row_1 - 1;

    let col0 = col_letters_to_index(col_letters)? as i32;

    let row = if row_abs { row0 } else { row0 - origin.row };
    let col = if col_abs { col0 } else { col0 - origin.col };

    Some(Ref::new(row, col, row_abs, col_abs))
}

fn col_letters_to_index(col: &str) -> Option<usize> {
    column_label_to_index_lenient(col)
        .ok()
        .and_then(|v| usize::try_from(v).ok())
}

#[cfg(test)]
mod tests {
    use super::*;
    use std::sync::Arc;

    #[test]
    fn parses_error_literals_as_scalar_values() {
        let origin = CellCoord::new(0, 0);

        assert_eq!(
            parse_formula("=#N/A", origin).expect("parse"),
            Expr::Literal(Value::Error(ErrorKind::NA))
        );
        assert_eq!(
            parse_formula("=#DIV/0!", origin).expect("parse"),
            Expr::Literal(Value::Error(ErrorKind::Div0))
        );
        assert_eq!(
            parse_formula("=#GETTING_DATA", origin).expect("parse"),
            Expr::Literal(Value::Error(ErrorKind::GettingData))
        );
    }

    #[test]
    fn parses_getting_data_error_literal_as_scalar_value() {
        let origin = CellCoord::new(0, 0);

        assert_eq!(
            parse_formula("=#GETTING_DATA", origin).expect("parse"),
            Expr::Literal(Value::Error(ErrorKind::GettingData))
        );
    }

    #[test]
    fn rejects_unknown_error_literals() {
        let origin = CellCoord::new(0, 0);

        assert_eq!(
            parse_formula("=#NOT_A_REAL_ERROR!", origin),
            Err(ParseError::UnexpectedToken(1))
        );
    }

    #[test]
    fn parses_concat_operator_as_concat_op_function_call() {
        let origin = CellCoord::new(0, 0);
        let expr = parse_formula("=\"a\"&\"b\"", origin).expect("parse");
        assert_eq!(
            expr,
            Expr::FuncCall {
                func: Function::ConcatOp,
                args: vec![
                    Expr::Literal(Value::Text(Arc::from("a"))),
                    Expr::Literal(Value::Text(Arc::from("b"))),
                ],
            }
        );
    }

    #[test]
    fn concat_binds_looser_than_addition() {
        let origin = CellCoord::new(0, 0);
        let expr = parse_formula("=1+2&3", origin).expect("parse");
        let Expr::FuncCall { func, args } = expr else {
            panic!("expected concat function call");
        };
        assert_eq!(func, Function::ConcatOp);
        assert_eq!(args.len(), 2);
        assert!(matches!(
            args[0],
            Expr::Binary {
                op: BinaryOp::Add,
                ..
            }
        ));
        assert!(matches!(
            args[1],
            Expr::Literal(Value::Number(n)) if n == 3.0
        ));
    }

    #[test]
    fn comparison_binds_looser_than_concat() {
        let origin = CellCoord::new(0, 0);
        let expr = parse_formula("=1&2=12", origin).expect("parse");
        let Expr::Binary { op, left, right } = expr else {
            panic!("expected binary expression");
        };
        assert_eq!(op, BinaryOp::Eq);
        assert!(matches!(
            left.as_ref(),
            Expr::FuncCall {
                func: Function::ConcatOp,
                ..
            }
        ));
        assert!(matches!(
            right.as_ref(),
            Expr::Literal(Value::Number(n)) if *n == 12.0
        ));
    }

    #[test]
    fn percent_binds_tighter_than_exponent() {
        let origin = CellCoord::new(0, 0);
        let expr = parse_formula("=2^3%", origin).expect("parse");
        let Expr::Binary { op, right, .. } = expr else {
            panic!("expected binary expression");
        };
        assert_eq!(op, BinaryOp::Pow);
        assert!(
            matches!(
                right.as_ref(),
                Expr::Binary {
                    op: BinaryOp::Div,
                    ..
                }
            ),
            "expected RHS to be lowered as division by 100"
        );
    }

    #[test]
    fn parses_spill_range_operator_on_cell_ref_as_reference_context_range() {
        let origin = CellCoord::new(0, 0);
        let expr = parse_formula("=A1#", origin).expect("parse");
        assert_eq!(
            expr,
            Expr::SpillRange(Box::new(Expr::RangeRef(RangeRef::new(
                Ref::new(0, 0, false, false),
                Ref::new(0, 0, false, false),
            ))))
        );
    }

    #[test]
    fn unary_minus_binds_looser_than_exponent() {
        let origin = CellCoord::new(0, 0);
        let expr = parse_formula("=-2^2", origin).expect("parse");
        let Expr::Unary { op, expr } = expr else {
            panic!("expected unary expression");
        };
        assert_eq!(op, UnaryOp::Neg);
        assert!(matches!(
            expr.as_ref(),
            Expr::Binary {
                op: BinaryOp::Pow,
                ..
            }
        ));
    }

    #[test]
    fn concat_chains_flatten_into_single_call() {
        let origin = CellCoord::new(0, 0);
        let expr = parse_formula("=\"a\"&\"b\"&\"c\"", origin).expect("parse");
        assert_eq!(
            expr,
            Expr::FuncCall {
                func: Function::ConcatOp,
                args: vec![
                    Expr::Literal(Value::Text(Arc::from("a"))),
                    Expr::Literal(Value::Text(Arc::from("b"))),
                    Expr::Literal(Value::Text(Arc::from("c"))),
                ],
            }
        );

        let expr = parse_formula("=\"a\"&(\"b\"&\"c\")", origin).expect("parse");
        assert_eq!(
            expr,
            Expr::FuncCall {
                func: Function::ConcatOp,
                args: vec![
                    Expr::Literal(Value::Text(Arc::from("a"))),
                    Expr::Literal(Value::Text(Arc::from("b"))),
                    Expr::Literal(Value::Text(Arc::from("c"))),
                ],
            }
        );
    }

    #[test]
    fn parses_string_literals_with_utf8_content() {
        let origin = CellCoord::new(0, 0);

        assert_eq!(
            parse_formula("=\"é\"", origin).expect("parse"),
            Expr::Literal(Value::Text(Arc::from("é")))
        );
        assert_eq!(
            parse_formula("=\"π\"", origin).expect("parse"),
            Expr::Literal(Value::Text(Arc::from("π")))
        );
        assert_eq!(
            parse_formula("=\"💩\"", origin).expect("parse"),
            Expr::Literal(Value::Text(Arc::from("💩")))
        );
        assert_eq!(
            parse_formula("=\"a💩b\"", origin).expect("parse"),
            Expr::Literal(Value::Text(Arc::from("a💩b")))
        );
    }

    #[test]
    fn rejects_function_calls_with_more_than_255_args() {
        let origin = CellCoord::new(0, 0);
        let mut args = String::new();
        for i in 0..(crate::EXCEL_MAX_ARGS + 1) {
            if i > 0 {
                args.push(',');
            }
            args.push('1');
        }
        let formula = format!("=SUM({args})");

        assert!(matches!(
            parse_formula(&formula, origin),
            Err(ParseError::TooManyArguments(_))
        ));
    }

    #[test]
    fn parses_missing_function_arguments_as_missing_literals() {
        let origin = CellCoord::new(0, 0);

        assert_eq!(
            parse_formula("=IF(FALSE,1,)", origin).expect("parse"),
            Expr::FuncCall {
                func: Function::If,
                args: vec![
                    Expr::Literal(Value::Bool(false)),
                    Expr::Literal(Value::Number(1.0)),
                    Expr::Literal(Value::Missing),
                ],
            }
        );

        assert_eq!(
            parse_formula("=ADDRESS(1,1,,FALSE)", origin).expect("parse"),
            Expr::FuncCall {
                func: Function::Address,
                args: vec![
                    Expr::Literal(Value::Number(1.0)),
                    Expr::Literal(Value::Number(1.0)),
                    Expr::Literal(Value::Missing),
                    Expr::Literal(Value::Bool(false)),
                ],
            }
        );

        assert_eq!(
            parse_formula("=IF(,1,2)", origin).expect("parse"),
            Expr::FuncCall {
                func: Function::If,
                args: vec![
                    Expr::Literal(Value::Missing),
                    Expr::Literal(Value::Number(1.0)),
                    Expr::Literal(Value::Number(2.0)),
                ],
            }
        );

        assert_eq!(
            parse_formula("=IF(,,)", origin).expect("parse"),
            Expr::FuncCall {
                func: Function::If,
                args: vec![
                    Expr::Literal(Value::Missing),
                    Expr::Literal(Value::Missing),
                    Expr::Literal(Value::Missing),
                ],
            }
        );
    }

    #[test]
    fn parses_true_false_as_zero_arg_functions() {
        let origin = CellCoord::new(0, 0);

        assert_eq!(
            parse_formula("=TRUE()", origin).expect("parse"),
            Expr::FuncCall {
                func: Function::True,
                args: vec![],
            }
        );
        assert_eq!(
            parse_formula("=FALSE()", origin).expect("parse"),
            Expr::FuncCall {
                func: Function::False,
                args: vec![],
            }
        );
    }

    #[test]
    fn parses_lambda_expressions() {
        let origin = CellCoord::new(0, 0);
        let expr = parse_formula("=LAMBDA(x, x+1)", origin).expect("parse");

        let Expr::Lambda { params, body } = expr else {
            panic!("expected lambda expression");
        };
        assert_eq!(params.len(), 1);
        assert_eq!(params[0].as_ref(), "X");

        let Expr::Binary { op, left, right } = body.as_ref() else {
            panic!("expected lambda body binary expression");
        };
        assert_eq!(*op, BinaryOp::Add);
        assert!(matches!(left.as_ref(), Expr::NameRef(name) if name.as_ref() == "X"));
        assert!(matches!(right.as_ref(), Expr::Literal(Value::Number(n)) if *n == 1.0));
    }

    #[test]
    fn parses_lambda_call_expressions() {
        let origin = CellCoord::new(0, 0);
        let expr = parse_formula("=LAMBDA(x, x+1)(3)", origin).expect("parse");

        let Expr::Call { callee, args } = expr else {
            panic!("expected call expression");
        };
        assert_eq!(args.len(), 1);
        assert!(matches!(args[0], Expr::Literal(Value::Number(n)) if n == 3.0));

        let Expr::Lambda { params, .. } = callee.as_ref() else {
            panic!("expected call callee to be a lambda");
        };
        assert_eq!(params.len(), 1);
        assert_eq!(params[0].as_ref(), "X");
    }
}
