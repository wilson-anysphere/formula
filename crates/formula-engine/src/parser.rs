use crate::locale::FormulaLocale;

#[derive(Debug, Clone, PartialEq)]
pub enum Expr {
    Number(f64),
    Identifier(String),
    FunctionCall { name: String, args: Vec<Expr> },
}

#[derive(Debug, Clone, PartialEq)]
pub struct Formula {
    pub root: Expr,
}

impl Formula {
    /// Serialize the formula in a canonical form suitable for persistence.
    ///
    /// Canonical form uses:
    /// - English function names (e.g. `SUM`)
    /// - `,` as argument separator
    /// - `.` as decimal separator
    pub fn to_canonical_string(&self) -> String {
        format!("={}", self.root.to_canonical_string())
    }

    /// Serialize the formula for display in a specific locale.
    pub fn to_localized_string(&self, locale: &FormulaLocale) -> String {
        format!("={}", self.root.to_localized_string(locale))
    }
}

impl Expr {
    fn to_canonical_string(&self) -> String {
        match self {
            Expr::Number(n) => canonical_number_string(*n),
            Expr::Identifier(s) => s.clone(),
            Expr::FunctionCall { name, args } => {
                let rendered_args = args
                    .iter()
                    .map(|a| a.to_canonical_string())
                    .collect::<Vec<_>>()
                    .join(",");
                format!("{name}({rendered_args})")
            }
        }
    }

    fn to_localized_string(&self, locale: &FormulaLocale) -> String {
        match self {
            Expr::Number(n) => localized_number_string(*n, locale.decimal_separator),
            Expr::Identifier(s) => s.clone(),
            Expr::FunctionCall { name, args } => {
                let localized_name = locale.localized_function_name(name);
                let rendered_args = args
                    .iter()
                    .map(|a| a.to_localized_string(locale))
                    .collect::<Vec<_>>()
                    .join(&locale.argument_separator.to_string());
                format!("{localized_name}({rendered_args})")
            }
        }
    }
}

fn canonical_number_string(value: f64) -> String {
    // `to_string` uses a locale-invariant representation with `.` decimals.
    value.to_string()
}

fn localized_number_string(value: f64, decimal_separator: char) -> String {
    let s = canonical_number_string(value);
    if decimal_separator == '.' {
        return s;
    }
    s.replace('.', &decimal_separator.to_string())
}

#[derive(Debug, Clone, PartialEq)]
pub enum ParseError {
    UnexpectedEof,
    UnexpectedChar { found: char, at: usize },
    ExpectedChar { expected: char, at: usize },
    InvalidNumber { at: usize },
    TrailingInput { at: usize },
}

pub fn parse_formula(input: &str, locale: &FormulaLocale) -> Result<Formula, ParseError> {
    let mut p = Parser::new(input, locale);
    p.skip_ws();
    if p.peek_char() == Some('=') {
        p.bump_char();
    }
    let root = p.parse_expr()?;
    p.skip_ws();
    if !p.is_eof() {
        return Err(ParseError::TrailingInput { at: p.pos });
    }
    Ok(Formula { root })
}

struct Parser<'a> {
    input: &'a str,
    locale: &'a FormulaLocale,
    pos: usize,
}

impl<'a> Parser<'a> {
    fn new(input: &'a str, locale: &'a FormulaLocale) -> Self {
        Self {
            input,
            locale,
            pos: 0,
        }
    }

    fn is_eof(&self) -> bool {
        self.pos >= self.input.len()
    }

    fn peek_char(&self) -> Option<char> {
        self.input[self.pos..].chars().next()
    }

    fn bump_char(&mut self) -> Option<char> {
        let c = self.peek_char()?;
        self.pos += c.len_utf8();
        Some(c)
    }

    fn skip_ws(&mut self) {
        while let Some(c) = self.peek_char() {
            if c.is_whitespace() {
                self.bump_char();
            } else {
                break;
            }
        }
    }

    fn consume_char(&mut self, expected: char) -> Result<(), ParseError> {
        self.skip_ws();
        match self.peek_char() {
            Some(c) if c == expected => {
                self.bump_char();
                Ok(())
            }
            Some(_) => Err(ParseError::ExpectedChar {
                expected,
                at: self.pos,
            }),
            None => Err(ParseError::UnexpectedEof),
        }
    }

    fn parse_expr(&mut self) -> Result<Expr, ParseError> {
        self.skip_ws();
        match self.peek_char() {
            Some(c) if c.is_ascii_digit() => self.parse_number(),
            Some(c) if is_identifier_start(c) => self.parse_identifier_or_call(),
            Some(c) => Err(ParseError::UnexpectedChar { found: c, at: self.pos }),
            None => Err(ParseError::UnexpectedEof),
        }
    }

    fn parse_number(&mut self) -> Result<Expr, ParseError> {
        self.skip_ws();
        let start = self.pos;

        let decimal = self.locale.decimal_separator;
        let thousands = self.locale.numeric_thousands_separator();

        let mut seen_decimal = false;
        let mut raw = String::new();

        while let Some(c) = self.peek_char() {
            if c.is_ascii_digit() {
                raw.push(c);
                self.bump_char();
                continue;
            }

            if Some(c) == thousands {
                // Skip grouping separators in numeric literals.
                self.bump_char();
                continue;
            }

            if c == decimal && !seen_decimal {
                raw.push('.');
                seen_decimal = true;
                self.bump_char();
                continue;
            }

            break;
        }

        if raw.is_empty() {
            return Err(ParseError::InvalidNumber { at: start });
        }

        match raw.parse::<f64>() {
            Ok(n) => Ok(Expr::Number(n)),
            Err(_) => Err(ParseError::InvalidNumber { at: start }),
        }
    }

    fn parse_identifier_or_call(&mut self) -> Result<Expr, ParseError> {
        self.skip_ws();
        let ident = self.parse_identifier()?;

        let save_pos = self.pos;
        self.skip_ws();
        if self.peek_char() == Some('(') {
            self.bump_char();
            let name = self.locale.canonical_function_name(&ident);
            let args = self.parse_arg_list()?;
            self.consume_char(')')?;
            Ok(Expr::FunctionCall { name, args })
        } else {
            // Not a function call; restore whitespace skipping side effects.
            self.pos = save_pos;
            Ok(Expr::Identifier(ident))
        }
    }

    fn parse_identifier(&mut self) -> Result<String, ParseError> {
        self.skip_ws();
        let mut out = String::new();
        match self.peek_char() {
            Some(c) if is_identifier_start(c) => {
                out.push(c);
                self.bump_char();
            }
            Some(c) => return Err(ParseError::UnexpectedChar { found: c, at: self.pos }),
            None => return Err(ParseError::UnexpectedEof),
        }

        while let Some(c) = self.peek_char() {
            if is_identifier_continue(c) {
                out.push(c);
                self.bump_char();
            } else {
                break;
            }
        }

        Ok(out)
    }

    fn parse_arg_list(&mut self) -> Result<Vec<Expr>, ParseError> {
        let mut args = Vec::new();
        self.skip_ws();
        if self.peek_char() == Some(')') {
            return Ok(args);
        }

        loop {
            let expr = self.parse_expr()?;
            args.push(expr);

            self.skip_ws();
            match self.peek_char() {
                Some(c) if c == self.locale.argument_separator => {
                    self.bump_char();
                    continue;
                }
                _ => break,
            }
        }

        Ok(args)
    }
}

fn is_identifier_start(c: char) -> bool {
    c.is_alphabetic() || c == '_' || c == '$'
}

fn is_identifier_continue(c: char) -> bool {
    c.is_alphanumeric() || c == '_' || c == '.' || c == '$'
}

