use crate::function_ids::{function_id_to_name, function_spec_from_id};

#[derive(Debug, thiserror::Error)]
pub enum DecodeRgceError {
    #[error("unexpected end of rgce stream")]
    UnexpectedEof,
    #[error("unsupported ptg token: 0x{ptg:02x}")]
    UnsupportedToken { ptg: u8 },
    #[error("unknown function id: 0x{func_id:04x}")]
    UnknownFunctionId { func_id: u16 },
    #[error("invalid utf16 string payload")]
    InvalidUtf16,
}

#[derive(Clone, Debug)]
struct ExprFragment {
    text: String,
    precedence: u8,
    /// `true` if this fragment contains the union operator (`,`).
    contains_union: bool,
    is_missing: bool,
}

impl ExprFragment {
    fn new(text: String) -> Self {
        Self {
            text,
            precedence: 100,
            contains_union: false,
            is_missing: false,
        }
    }

    fn missing() -> Self {
        Self {
            text: String::new(),
            precedence: 100,
            contains_union: false,
            is_missing: true,
        }
    }
}

// Precedence values match `formula-engine` (`Expr::precedence`).
fn binary_precedence(ptg: u8) -> Option<u8> {
    match ptg {
        0x11 => Some(82),                                    // range (:)
        0x0F => Some(81),                                    // intersect ( )
        0x10 => Some(80),                                    // union (,)
        0x07 => Some(50),                                    // power (^)
        0x05 | 0x06 => Some(40),                             // mul/div
        0x03 | 0x04 => Some(30),                             // add/sub
        0x08 => Some(20),                                    // concat (&)
        0x09 | 0x0A | 0x0B | 0x0C | 0x0D | 0x0E => Some(10), // comparisons
        _ => None,
    }
}

fn op_str(ptg: u8) -> Option<&'static str> {
    match ptg {
        0x03 => Some("+"),
        0x04 => Some("-"),
        0x05 => Some("*"),
        0x06 => Some("/"),
        0x07 => Some("^"),
        0x08 => Some("&"),
        0x09 => Some("<"),
        0x0A => Some("<="),
        0x0B => Some("="),
        0x0C => Some(">"),
        0x0D => Some(">="),
        0x0E => Some("<>"),
        0x0F => Some(" "),
        0x10 => Some(","),
        0x11 => Some(":"),
        _ => None,
    }
}

/// Best-effort decode of a BIFF12 `rgce` token stream into formula text.
///
/// The returned string does **not** include a leading `=`.
pub fn decode_rgce(rgce: &[u8]) -> Result<String, DecodeRgceError> {
    if rgce.is_empty() {
        return Ok(String::new());
    }

    let mut input = rgce;
    let mut stack: Vec<ExprFragment> = Vec::new();

    while !input.is_empty() {
        let ptg = input[0];
        input = &input[1..];

        match ptg {
            // Binary operators.
            0x03..=0x11 => {
                let Some(op) = op_str(ptg) else {
                    return Err(DecodeRgceError::UnsupportedToken { ptg });
                };
                let prec = binary_precedence(ptg).expect("precedence for binary ops");

                let right = stack.pop().ok_or(DecodeRgceError::UnexpectedEof)?;
                let left = stack.pop().ok_or(DecodeRgceError::UnexpectedEof)?;

                let left_s = if left.precedence < prec && !left.is_missing {
                    format!("({})", left.text)
                } else {
                    left.text
                };
                let right_s = if right.precedence < prec && !right.is_missing {
                    format!("({})", right.text)
                } else {
                    right.text
                };

                let mut text = String::with_capacity(left_s.len() + op.len() + right_s.len());
                text.push_str(&left_s);
                text.push_str(op);
                text.push_str(&right_s);

                stack.push(ExprFragment {
                    text,
                    precedence: prec,
                    contains_union: left.contains_union || right.contains_union || ptg == 0x10,
                    is_missing: false,
                });
            }
            // Unary plus/minus.
            0x12 | 0x13 => {
                let prec = 70;
                let op = if ptg == 0x12 { "+" } else { "-" };
                let expr = stack.pop().ok_or(DecodeRgceError::UnexpectedEof)?;
                let inner = if expr.precedence < prec && !expr.is_missing {
                    format!("({})", expr.text)
                } else {
                    expr.text
                };
                stack.push(ExprFragment {
                    text: format!("{op}{inner}"),
                    precedence: prec,
                    contains_union: expr.contains_union,
                    is_missing: false,
                });
            }
            // Percent postfix.
            0x14 => {
                let prec = 60;
                let expr = stack.pop().ok_or(DecodeRgceError::UnexpectedEof)?;
                let inner = if expr.precedence < prec && !expr.is_missing {
                    format!("({})", expr.text)
                } else {
                    expr.text
                };
                stack.push(ExprFragment {
                    text: format!("{inner}%"),
                    precedence: prec,
                    contains_union: expr.contains_union,
                    is_missing: false,
                });
            }
            // Explicit parentheses.
            0x15 => {
                let expr = stack.pop().ok_or(DecodeRgceError::UnexpectedEof)?;
                stack.push(ExprFragment {
                    text: format!("({})", expr.text),
                    precedence: 100,
                    contains_union: expr.contains_union,
                    is_missing: false,
                });
            }
            // Missing argument.
            0x16 => {
                stack.push(ExprFragment::missing());
            }
            // String literal.
            0x17 => {
                if input.len() < 2 {
                    return Err(DecodeRgceError::UnexpectedEof);
                }
                let cch = u16::from_le_bytes([input[0], input[1]]) as usize;
                input = &input[2..];
                let byte_len = cch.checked_mul(2).ok_or(DecodeRgceError::UnexpectedEof)?;
                if input.len() < byte_len {
                    return Err(DecodeRgceError::UnexpectedEof);
                }
                let raw = &input[..byte_len];
                input = &input[byte_len..];

                let mut units = Vec::with_capacity(cch);
                for chunk in raw.chunks_exact(2) {
                    units.push(u16::from_le_bytes([chunk[0], chunk[1]]));
                }
                let s = String::from_utf16(&units).map_err(|_| DecodeRgceError::InvalidUtf16)?;
                let escaped = s.replace('"', "\"\"");
                stack.push(ExprFragment::new(format!("\"{escaped}\"")));
            }
            // Error literal.
            0x1C => {
                if input.is_empty() {
                    return Err(DecodeRgceError::UnexpectedEof);
                }
                let err = input[0];
                input = &input[1..];
                let text = match err {
                    0x00 => "#NULL!",
                    0x07 => "#DIV/0!",
                    0x0F => "#VALUE!",
                    0x17 => "#REF!",
                    0x1D => "#NAME?",
                    0x24 => "#NUM!",
                    0x2A => "#N/A",
                    0x2B => "#GETTING_DATA",
                    _ => return Err(DecodeRgceError::UnsupportedToken { ptg }),
                };
                stack.push(ExprFragment::new(text.to_string()));
            }
            // Bool literal.
            0x1D => {
                if input.is_empty() {
                    return Err(DecodeRgceError::UnexpectedEof);
                }
                let b = input[0];
                input = &input[1..];
                stack.push(ExprFragment::new(
                    if b == 0 { "FALSE" } else { "TRUE" }.to_string(),
                ));
            }
            // Int literal.
            0x1E => {
                if input.len() < 2 {
                    return Err(DecodeRgceError::UnexpectedEof);
                }
                let n = u16::from_le_bytes([input[0], input[1]]);
                input = &input[2..];
                stack.push(ExprFragment::new(n.to_string()));
            }
            // Num literal.
            0x1F => {
                if input.len() < 8 {
                    return Err(DecodeRgceError::UnexpectedEof);
                }
                let mut bytes = [0u8; 8];
                bytes.copy_from_slice(&input[..8]);
                input = &input[8..];
                stack.push(ExprFragment::new(f64::from_le_bytes(bytes).to_string()));
            }
            // PtgFunc
            0x21 | 0x41 | 0x61 => {
                if input.len() < 2 {
                    return Err(DecodeRgceError::UnexpectedEof);
                }
                let func_id = u16::from_le_bytes([input[0], input[1]]);
                input = &input[2..];

                let Some(spec) = function_spec_from_id(func_id) else {
                    return Err(DecodeRgceError::UnknownFunctionId { func_id });
                };
                if spec.min_args != spec.max_args {
                    return Err(DecodeRgceError::UnknownFunctionId { func_id });
                }

                let argc = spec.min_args as usize;
                let mut args = Vec::with_capacity(argc);
                for _ in 0..argc {
                    args.push(stack.pop().ok_or(DecodeRgceError::UnexpectedEof)?);
                }
                args.reverse();

                let mut text = String::new();
                text.push_str(spec.name);
                text.push('(');
                for (i, arg) in args.into_iter().enumerate() {
                    if i > 0 {
                        text.push(',');
                    }
                    if arg.is_missing {
                        continue;
                    }
                    if arg.contains_union {
                        text.push('(');
                        text.push_str(&arg.text);
                        text.push(')');
                    } else {
                        text.push_str(&arg.text);
                    }
                }
                text.push(')');

                stack.push(ExprFragment::new(text));
            }
            // PtgFuncVar
            0x22 | 0x42 | 0x62 => {
                if input.len() < 3 {
                    return Err(DecodeRgceError::UnexpectedEof);
                }
                let argc = input[0] as usize;
                let func_id = u16::from_le_bytes([input[1], input[2]]);
                input = &input[3..];

                let name = function_id_to_name(func_id)
                    .ok_or(DecodeRgceError::UnknownFunctionId { func_id })?;

                let mut args = Vec::with_capacity(argc);
                for _ in 0..argc {
                    args.push(stack.pop().ok_or(DecodeRgceError::UnexpectedEof)?);
                }
                args.reverse();

                let mut text = String::new();
                text.push_str(name);
                text.push('(');
                for (i, arg) in args.into_iter().enumerate() {
                    if i > 0 {
                        text.push(',');
                    }
                    if arg.is_missing {
                        continue;
                    }
                    if arg.contains_union {
                        text.push('(');
                        text.push_str(&arg.text);
                        text.push(')');
                    } else {
                        text.push_str(&arg.text);
                    }
                }
                text.push(')');

                stack.push(ExprFragment::new(text));
            }
            // PtgRef
            0x24 | 0x44 | 0x64 => {
                if input.len() < 6 {
                    return Err(DecodeRgceError::UnexpectedEof);
                }
                let row = u32::from_le_bytes([input[0], input[1], input[2], input[3]]) + 1;
                let col = u16::from_le_bytes([input[4], input[5] & 0x3F]) as u32;
                let flags = input[5];
                input = &input[6..];

                let mut text = String::new();
                if flags & 0x80 == 0 {
                    text.push('$');
                }
                push_column(col, &mut text);
                if flags & 0x40 == 0 {
                    text.push('$');
                }
                text.push_str(&row.to_string());
                stack.push(ExprFragment::new(text));
            }
            // PtgArea
            0x25 | 0x45 | 0x65 => {
                if input.len() < 12 {
                    return Err(DecodeRgceError::UnexpectedEof);
                }
                let row1 = u32::from_le_bytes([input[0], input[1], input[2], input[3]]) + 1;
                let row2 = u32::from_le_bytes([input[4], input[5], input[6], input[7]]) + 1;
                let col1 = u16::from_le_bytes([input[8], input[9] & 0x3F]) as u32;
                let col2 = u16::from_le_bytes([input[10], input[11] & 0x3F]) as u32;
                let flags1 = input[9];
                let flags2 = input[11];
                input = &input[12..];

                let mut start = String::new();
                if flags1 & 0x80 == 0 {
                    start.push('$');
                }
                push_column(col1, &mut start);
                if flags1 & 0x40 == 0 {
                    start.push('$');
                }
                start.push_str(&row1.to_string());

                let mut end = String::new();
                if flags2 & 0x80 == 0 {
                    end.push('$');
                }
                push_column(col2, &mut end);
                if flags2 & 0x40 == 0 {
                    end.push('$');
                }
                end.push_str(&row2.to_string());

                stack.push(ExprFragment::new(format!("{start}:{end}")));
            }
            _ => return Err(DecodeRgceError::UnsupportedToken { ptg }),
        }
    }

    if stack.len() == 1 {
        Ok(stack.pop().expect("len checked").text)
    } else {
        Err(DecodeRgceError::UnsupportedToken { ptg: 0x00 })
    }
}

fn push_column(mut col: u32, out: &mut String) {
    // Excel column labels are 1-based.
    col += 1;
    let mut buf = [0u8; 10];
    let mut i = 0usize;
    while col > 0 {
        let rem = ((col - 1) % 26) as u8;
        buf[i] = b'A' + rem;
        i += 1;
        col = (col - 1) / 26;
    }
    for ch in buf[..i].iter().rev() {
        out.push(*ch as char);
    }
}

#[cfg(feature = "encode")]
#[derive(Debug, thiserror::Error)]
pub enum EncodeRgceError {
    #[error("formula parse error: {message} (span {start}..{end})")]
    Parse {
        message: String,
        start: usize,
        end: usize,
    },
    #[error("unsupported expression in BIFF12 encoder: {0}")]
    Unsupported(&'static str),
    #[error("unsupported function name: {0}")]
    UnknownFunction(String),
    #[error("invalid argument count for {name}: got {got}, expected {min}..={max}")]
    InvalidArgCount {
        name: String,
        got: usize,
        min: u8,
        max: u8,
    },
    #[error("invalid numeric literal: {0}")]
    InvalidNumber(String),
    #[error("unsupported error literal: {0}")]
    InvalidErrorLiteral(String),
}

#[cfg(feature = "encode")]
pub fn encode_rgce(formula: &str) -> Result<Vec<u8>, EncodeRgceError> {
    use formula_engine::{parse_formula, ParseOptions};

    let ast =
        parse_formula(formula, ParseOptions::default()).map_err(|e| EncodeRgceError::Parse {
            message: e.message,
            start: e.span.start,
            end: e.span.end,
        })?;
    let mut out = Vec::new();
    encode_expr(&ast.expr, &mut out)?;
    Ok(out)
}

#[cfg(feature = "encode")]
fn encode_expr(expr: &formula_engine::Expr, out: &mut Vec<u8>) -> Result<(), EncodeRgceError> {
    use formula_engine::{BinaryOp, Coord, Expr, PostfixOp, UnaryOp};

    match expr {
        Expr::Number(raw) => {
            let n: f64 = raw
                .parse()
                .map_err(|_| EncodeRgceError::InvalidNumber(raw.clone()))?;
            if n.fract() == 0.0 && n >= 0.0 && n <= u16::MAX as f64 {
                out.push(0x1E); // PtgInt
                out.extend_from_slice(&(n as u16).to_le_bytes());
            } else {
                out.push(0x1F); // PtgNum
                out.extend_from_slice(&n.to_le_bytes());
            }
        }
        Expr::String(s) => {
            out.push(0x17); // PtgStr
            let units: Vec<u16> = s.encode_utf16().collect();
            let cch: u16 = units
                .len()
                .try_into()
                .map_err(|_| EncodeRgceError::Unsupported("string literal too long"))?;
            out.extend_from_slice(&cch.to_le_bytes());
            for u in units {
                out.extend_from_slice(&u.to_le_bytes());
            }
        }
        Expr::Boolean(b) => {
            out.push(0x1D); // PtgBool
            out.push(if *b { 1 } else { 0 });
        }
        Expr::Error(raw) => {
            let code = match raw.to_ascii_uppercase().as_str() {
                "#NULL!" => 0x00,
                "#DIV/0!" => 0x07,
                "#VALUE!" => 0x0F,
                "#REF!" => 0x17,
                "#NAME?" => 0x1D,
                "#NUM!" => 0x24,
                "#N/A" => 0x2A,
                "#GETTING_DATA" => 0x2B,
                _ => return Err(EncodeRgceError::InvalidErrorLiteral(raw.clone())),
            };
            out.push(0x1C); // PtgErr
            out.push(code);
        }
        Expr::CellRef(r) => {
            if r.workbook.is_some() || r.sheet.is_some() {
                return Err(EncodeRgceError::Unsupported(
                    "3D/sheet-qualified references",
                ));
            }
            let (col, col_abs) = match &r.col {
                Coord::A1 { index, abs } => (*index, *abs),
                Coord::Offset(_) => return Err(EncodeRgceError::Unsupported("relative offsets")),
            };
            let (row, row_abs) = match &r.row {
                Coord::A1 { index, abs } => (*index, *abs),
                Coord::Offset(_) => return Err(EncodeRgceError::Unsupported("relative offsets")),
            };
            out.push(0x24); // PtgRef
            out.extend_from_slice(&row.to_le_bytes());
            out.extend_from_slice(&encode_col_with_flags(col, col_abs, row_abs));
        }
        Expr::Binary(b) if b.op == BinaryOp::Range => {
            // Prefer encoding simple A1:A2 areas as PtgArea for Excel-compatible rgce.
            if let (Expr::CellRef(a), Expr::CellRef(bref)) = (&*b.left, &*b.right) {
                if a.workbook.is_none()
                    && a.sheet.is_none()
                    && bref.workbook.is_none()
                    && bref.sheet.is_none()
                {
                    if let (Some((c1, c1_abs)), Some((r1, r1_abs))) =
                        (coord_to_a1(&a.col), coord_to_a1(&a.row))
                    {
                        if let (Some((c2, c2_abs)), Some((r2, r2_abs))) =
                            (coord_to_a1(&bref.col), coord_to_a1(&bref.row))
                        {
                            out.push(0x25); // PtgArea
                            out.extend_from_slice(&r1.to_le_bytes());
                            out.extend_from_slice(&r2.to_le_bytes());
                            out.extend_from_slice(&encode_col_with_flags(c1, c1_abs, r1_abs));
                            out.extend_from_slice(&encode_col_with_flags(c2, c2_abs, r2_abs));
                            return Ok(());
                        }
                    }
                }
            }

            // Fallback: encode as operator.
            encode_expr(&b.left, out)?;
            encode_expr(&b.right, out)?;
            out.push(0x11); // PtgRange
        }
        Expr::Binary(b) => {
            encode_expr(&b.left, out)?;
            encode_expr(&b.right, out)?;
            let ptg = match b.op {
                BinaryOp::Add => 0x03,
                BinaryOp::Sub => 0x04,
                BinaryOp::Mul => 0x05,
                BinaryOp::Div => 0x06,
                BinaryOp::Pow => 0x07,
                BinaryOp::Concat => 0x08,
                BinaryOp::Lt => 0x09,
                BinaryOp::Le => 0x0A,
                BinaryOp::Eq => 0x0B,
                BinaryOp::Gt => 0x0C,
                BinaryOp::Ge => 0x0D,
                BinaryOp::Ne => 0x0E,
                BinaryOp::Intersect => 0x0F,
                BinaryOp::Union => 0x10,
                BinaryOp::Range => 0x11,
            };
            out.push(ptg);
        }
        Expr::Unary(u) => {
            encode_expr(&u.expr, out)?;
            match u.op {
                UnaryOp::Plus => out.push(0x12),
                UnaryOp::Minus => out.push(0x13),
                UnaryOp::ImplicitIntersection => {
                    return Err(EncodeRgceError::Unsupported("implicit intersection (@)"));
                }
            }
        }
        Expr::Postfix(p) => {
            encode_expr(&p.expr, out)?;
            match p.op {
                PostfixOp::Percent => out.push(0x14),
                PostfixOp::SpillRange => {
                    return Err(EncodeRgceError::Unsupported("spill range (#)"));
                }
            }
        }
        Expr::FunctionCall(call) => {
            let name = call.name.name_upper.as_str();
            let Some(func) = crate::function_ids::function_spec_from_name(name) else {
                return Err(EncodeRgceError::UnknownFunction(name.to_string()));
            };

            let argc_usize = call.args.len();
            if argc_usize < func.min_args as usize || argc_usize > func.max_args as usize {
                return Err(EncodeRgceError::InvalidArgCount {
                    name: name.to_string(),
                    got: argc_usize,
                    min: func.min_args,
                    max: func.max_args,
                });
            }

            // Encode args.
            for arg in &call.args {
                if matches!(arg, Expr::Missing) {
                    out.push(0x16); // PtgMissArg
                } else {
                    encode_expr(arg, out)?;
                }
            }

            // Choose token form.
            if func.min_args == func.max_args {
                if argc_usize != func.min_args as usize {
                    return Err(EncodeRgceError::InvalidArgCount {
                        name: name.to_string(),
                        got: argc_usize,
                        min: func.min_args,
                        max: func.max_args,
                    });
                }
                // Fixed arity -> PtgFunc
                out.push(0x21);
                out.extend_from_slice(&func.id.to_le_bytes());
            } else {
                // Variable arity -> PtgFuncVar
                out.push(0x22);
                let argc: u8 = call
                    .args
                    .len()
                    .try_into()
                    .map_err(|_| EncodeRgceError::Unsupported("too many function args"))?;
                out.push(argc);
                out.extend_from_slice(&func.id.to_le_bytes());
            }
        }
        Expr::Missing => {
            out.push(0x16); // PtgMissArg
        }
        Expr::NameRef(_) => return Err(EncodeRgceError::Unsupported("named references")),
        Expr::ColRef(_) => return Err(EncodeRgceError::Unsupported("column references")),
        Expr::RowRef(_) => return Err(EncodeRgceError::Unsupported("row references")),
        Expr::StructuredRef(_) => {
            return Err(EncodeRgceError::Unsupported("structured references"))
        }
        Expr::Array(_) => return Err(EncodeRgceError::Unsupported("array literals")),
    }

    Ok(())
}

#[cfg(feature = "encode")]
fn coord_to_a1(coord: &formula_engine::Coord) -> Option<(u32, bool)> {
    match coord {
        formula_engine::Coord::A1 { index, abs } => Some((*index, *abs)),
        formula_engine::Coord::Offset(_) => None,
    }
}

#[cfg(feature = "encode")]
fn encode_col_with_flags(col: u32, col_abs: bool, row_abs: bool) -> [u8; 2] {
    let col: u16 = col as u16;
    let [lo, hi] = col.to_le_bytes();
    let mut hi = hi & 0x3F;
    // In BIFF, these flags indicate *relative* (absence of '$').
    if !col_abs {
        hi |= 0x80;
    }
    if !row_abs {
        hi |= 0x40;
    }
    [lo, hi]
}
