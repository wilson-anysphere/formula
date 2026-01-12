use super::value::{CellCoord, ErrorKind, RangeRef, Ref, Value};
use std::sync::Arc;

#[derive(Clone, Copy, Debug, PartialEq, Eq)]
pub enum UnaryOp {
    Plus,
    Neg,
    ImplicitIntersection,
}

#[derive(Clone, Copy, Debug, PartialEq, Eq)]
pub enum BinaryOp {
    Add,
    Sub,
    Mul,
    Div,
    Pow,
    Eq,
    Ne,
    Lt,
    Le,
    Gt,
    Ge,
}

#[derive(Clone, Debug, PartialEq, Eq)]
pub enum Function {
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
    Concat,
    Not,
    Unknown(Arc<str>),
}

impl Function {
    pub fn from_name(name: &str) -> Self {
        match name.to_ascii_uppercase().as_str() {
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
            "COUNTIF" => Function::CountIf,
            "COUNTIFS" => Function::CountIfs,
            "SUMPRODUCT" => Function::SumProduct,
            "VLOOKUP" => Function::VLookup,
            "HLOOKUP" => Function::HLookup,
            "MATCH" => Function::Match,
            "ABS" => Function::Abs,
            "INT" => Function::Int,
            "ROUND" => Function::Round,
            "ROUNDUP" => Function::RoundUp,
            "ROUNDDOWN" => Function::RoundDown,
            "MOD" => Function::Mod,
            "SIGN" => Function::Sign,
            "CONCAT" => Function::Concat,
            "NOT" => Function::Not,
            other => Function::Unknown(Arc::from(other)),
        }
    }

    pub fn name(&self) -> &str {
        match self {
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
            Function::Concat => "CONCAT",
            Function::Not => "NOT",
            Function::Unknown(s) => s,
        }
    }
}

#[derive(Clone, Debug, PartialEq, Eq)]
pub enum Expr {
    Literal(Value),
    CellRef(Ref),
    RangeRef(RangeRef),
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
}

#[derive(thiserror::Error, Debug, Clone, PartialEq, Eq)]
pub enum ParseError {
    #[error("unexpected end of input")]
    UnexpectedEof,
    #[error("unexpected token at byte {0}")]
    UnexpectedToken(usize),
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
        let expr = self.parse_bp(0)?;
        self.skip_ws();
        if self.pos != self.input.len() {
            return Err(ParseError::UnexpectedToken(self.pos));
        }
        Ok(expr)
    }

    fn parse_bp(&mut self, min_bp: u8) -> Result<Expr, ParseError> {
        self.skip_ws();
        let mut lhs = self.parse_prefix()?;
        loop {
            self.skip_ws();
            let op_pos = self.pos;
            let (op, l_bp, r_bp) = match self.peek_infix_op() {
                Some(v) => v,
                None => break,
            };
            if l_bp < min_bp {
                break;
            }
            self.pos = op_pos + op.token_len();
            let rhs = self.parse_bp(r_bp)?;
            lhs = Expr::Binary {
                op,
                left: Box::new(lhs),
                right: Box::new(rhs),
            };
        }
        Ok(lhs)
    }

    fn parse_prefix(&mut self) -> Result<Expr, ParseError> {
        self.skip_ws();
        match self.peek_byte() {
            Some(b'+') => {
                self.pos += 1;
                Ok(Expr::Unary {
                    op: UnaryOp::Plus,
                    expr: Box::new(self.parse_bp(10)?),
                })
            }
            Some(b'-') => {
                self.pos += 1;
                Ok(Expr::Unary {
                    op: UnaryOp::Neg,
                    expr: Box::new(self.parse_bp(10)?),
                })
            }
            Some(b'@') => {
                self.pos += 1;
                Ok(Expr::Unary {
                    op: UnaryOp::ImplicitIntersection,
                    expr: Box::new(self.parse_bp(10)?),
                })
            }
            Some(b'(') => {
                self.pos += 1;
                let expr = self.parse_bp(0)?;
                self.skip_ws();
                if self.peek_byte() != Some(b')') {
                    return Err(ParseError::UnexpectedToken(self.pos));
                }
                self.pos += 1;
                Ok(expr)
            }
            Some(b'"') => self.parse_string(),
            Some(b'#') => self.parse_error_literal(),
            Some(b'0'..=b'9') | Some(b'.') => self.parse_number(),
            Some(_) => self.parse_ident_like(),
            None => Err(ParseError::UnexpectedEof),
        }
    }

    fn parse_error_literal(&mut self) -> Result<Expr, ParseError> {
        debug_assert_eq!(self.peek_byte(), Some(b'#'));
        let start = self.pos;

        const ERROR_LITERALS: &[(&str, ErrorKind)] = &[
            ("#NULL!", ErrorKind::Null),
            ("#DIV/0!", ErrorKind::Div0),
            ("#VALUE!", ErrorKind::Value),
            ("#REF!", ErrorKind::Ref),
            ("#NAME?", ErrorKind::Name),
            ("#NUM!", ErrorKind::Num),
            ("#N/A", ErrorKind::NA),
            ("#SPILL!", ErrorKind::Spill),
            ("#CALC!", ErrorKind::Calc),
        ];

        for &(lit, kind) in ERROR_LITERALS {
            let end = start.saturating_add(lit.len());
            if self
                .input
                .get(start..end)
                .is_some_and(|slice| slice.eq_ignore_ascii_case(lit.as_bytes()))
            {
                self.pos = end;
                return Ok(Expr::Literal(Value::Error(kind)));
            }
        }

        // Fallback: accept unknown `#...` sequences as error tokens and coerce them to `#VALUE!`,
        // mirroring the canonical parser + compiler behavior.
        self.pos += 1; // '#'
        while let Some(b) = self.peek_byte() {
            if matches!(b, b'_' | b'/' | b'.' | b'0'..=b'9' | b'A'..=b'Z' | b'a'..=b'z') {
                self.pos += 1;
            } else {
                break;
            }
        }

        if self.pos == start + 1 {
            return Err(ParseError::UnexpectedToken(start));
        }

        if matches!(self.peek_byte(), Some(b'!' | b'?')) {
            self.pos += 1;
        }

        Ok(Expr::Literal(Value::Error(ErrorKind::Value)))
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
        let mut out = String::new();
        while let Some(b) = self.peek_byte() {
            self.pos += 1;
            match b {
                b'"' => {
                    // Excel escapes quotes by doubling them.
                    if self.peek_byte() == Some(b'"') {
                        self.pos += 1;
                        out.push('"');
                    } else {
                        return Ok(Expr::Literal(Value::Text(Arc::from(out))));
                    }
                }
                _ => out.push(b as char),
            }
        }
        Err(ParseError::UnterminatedString)
    }

    fn parse_ident_like(&mut self) -> Result<Expr, ParseError> {
        let start = self.pos;
        while let Some(b) = self.peek_byte() {
            match b {
                b'A'..=b'Z' | b'a'..=b'z' | b'0'..=b'9' | b'_' | b'$' => self.pos += 1,
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
            let mut args = Vec::new();
            self.skip_ws();
            if self.peek_byte() != Some(b')') {
                loop {
                    args.push(self.parse_bp(0)?);
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
            return Ok(Expr::FuncCall {
                func: Function::from_name(ident),
                args,
            });
        }

        let upper = ident.to_ascii_uppercase();
        match upper.as_str() {
            "TRUE" => return Ok(Expr::Literal(Value::Bool(true))),
            "FALSE" => return Ok(Expr::Literal(Value::Bool(false))),
            _ => {}
        }

        let first = parse_a1_ref(ident, self.origin).ok_or(ParseError::InvalidCellRef)?;
        self.skip_ws();
        if self.peek_byte() == Some(b':') {
            self.pos += 1;
            self.skip_ws();
            let start2 = self.pos;
            while let Some(b) = self.peek_byte() {
                match b {
                    b'A'..=b'Z' | b'a'..=b'z' | b'0'..=b'9' | b'_' | b'$' => self.pos += 1,
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

    fn peek_infix_op(&self) -> Option<(BinaryOp, u8, u8)> {
        let b0 = *self.input.get(self.pos)?;
        let b1 = *self.input.get(self.pos + 1).unwrap_or(&0);
        let (op, l_bp, r_bp) = match (b0, b1) {
            (b'+', _) => (BinaryOp::Add, 5, 6),
            (b'-', _) => (BinaryOp::Sub, 5, 6),
            (b'*', _) => (BinaryOp::Mul, 7, 8),
            (b'/', _) => (BinaryOp::Div, 7, 8),
            (b'^', _) => (BinaryOp::Pow, 9, 9), // right associative
            (b'=', _) => (BinaryOp::Eq, 3, 4),
            (b'<', b'>') => (BinaryOp::Ne, 3, 4),
            (b'<', b'=') => (BinaryOp::Le, 3, 4),
            (b'>', b'=') => (BinaryOp::Ge, 3, 4),
            (b'<', _) => (BinaryOp::Lt, 3, 4),
            (b'>', _) => (BinaryOp::Gt, 3, 4),
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

    fn peek_byte(&self) -> Option<u8> {
        self.input.get(self.pos).copied()
    }
}

trait BinaryOpExt {
    fn token_len(&self) -> usize;
}

impl BinaryOpExt for BinaryOp {
    fn token_len(&self) -> usize {
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
    let mut acc: usize = 0;
    for b in col.bytes() {
        let u = match b {
            b'A'..=b'Z' => (b - b'A') as usize + 1,
            b'a'..=b'z' => (b - b'a') as usize + 1,
            _ => return None,
        };
        acc = acc.checked_mul(26)?.checked_add(u)?;
    }
    Some(acc - 1)
}

#[cfg(test)]
mod tests {
    use super::*;

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
    }

    #[test]
    fn parses_unknown_error_literals_as_value_error() {
        let origin = CellCoord::new(0, 0);

        assert_eq!(
            parse_formula("=#GETTING_DATA", origin).expect("parse"),
            Expr::Literal(Value::Error(ErrorKind::Value))
        );
    }
}
