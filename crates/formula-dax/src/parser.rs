use crate::engine::{DaxError, DaxResult};

#[derive(Clone, Debug, PartialEq)]
pub enum Expr {
    Number(f64),
    Text(String),
    Boolean(bool),
    TableLiteral {
        rows: Vec<Vec<Expr>>,
    },
    TableName(String),
    Measure(String),
    Let {
        bindings: Vec<(String, Expr)>,
        body: Box<Expr>,
    },
    ColumnRef {
        table: String,
        column: String,
    },
    Call {
        name: String,
        args: Vec<Expr>,
    },
    UnaryOp {
        op: UnaryOp,
        expr: Box<Expr>,
    },
    BinaryOp {
        op: BinaryOp,
        left: Box<Expr>,
        right: Box<Expr>,
    },
}

#[derive(Clone, Copy, Debug, PartialEq, Eq)]
pub enum UnaryOp {
    Negate,
}

#[derive(Clone, Copy, Debug, PartialEq, Eq)]
pub enum BinaryOp {
    Add,
    Subtract,
    Multiply,
    Divide,
    Concat,
    Equals,
    NotEquals,
    Less,
    LessEquals,
    Greater,
    GreaterEquals,
    In,
    And,
    Or,
}

#[derive(Clone, Debug, PartialEq)]
enum Token {
    Identifier(String),
    BracketIdentifier(String),
    Number(f64),
    String(String),
    Var,
    Return,
    Comma,
    Semicolon,
    LParen,
    RParen,
    LBrace,
    RBrace,
    Plus,
    Minus,
    Star,
    Slash,
    Ampersand,
    Equals,
    NotEquals,
    Less,
    LessEquals,
    Greater,
    GreaterEquals,
    In,
    AndAnd,
    OrOr,
    Eof,
}

struct Lexer<'a> {
    input: &'a str,
    chars: std::str::Chars<'a>,
    peeked: Option<char>,
}

impl<'a> Lexer<'a> {
    fn new(input: &'a str) -> Self {
        let mut chars = input.chars();
        let peeked = chars.next();
        Self {
            input,
            chars,
            peeked,
        }
    }

    fn bump(&mut self) -> Option<char> {
        let current = self.peeked.take();
        self.peeked = self.chars.next();
        current
    }

    fn peek(&self) -> Option<char> {
        self.peeked
    }

    fn consume_while<F>(&mut self, mut predicate: F) -> String
    where
        F: FnMut(char) -> bool,
    {
        let mut buf = String::new();
        while let Some(ch) = self.peek() {
            if !predicate(ch) {
                break;
            }
            buf.push(ch);
            self.bump();
        }
        buf
    }

    fn skip_whitespace(&mut self) {
        self.consume_while(|c| c.is_whitespace());
    }

    fn next_token(&mut self) -> DaxResult<Token> {
        self.skip_whitespace();
        let Some(ch) = self.peek() else {
            return Ok(Token::Eof);
        };

        match ch {
            '(' => {
                self.bump();
                Ok(Token::LParen)
            }
            ')' => {
                self.bump();
                Ok(Token::RParen)
            }
            '{' => {
                self.bump();
                Ok(Token::LBrace)
            }
            '}' => {
                self.bump();
                Ok(Token::RBrace)
            }
            ',' => {
                self.bump();
                Ok(Token::Comma)
            }
            ';' => {
                self.bump();
                Ok(Token::Semicolon)
            }
            '+' => {
                self.bump();
                Ok(Token::Plus)
            }
            '-' => {
                self.bump();
                Ok(Token::Minus)
            }
            '*' => {
                self.bump();
                Ok(Token::Star)
            }
            '/' => {
                self.bump();
                Ok(Token::Slash)
            }
            '=' => {
                self.bump();
                Ok(Token::Equals)
            }
            '<' => {
                self.bump();
                match self.peek() {
                    Some('=') => {
                        self.bump();
                        Ok(Token::LessEquals)
                    }
                    Some('>') => {
                        self.bump();
                        Ok(Token::NotEquals)
                    }
                    _ => Ok(Token::Less),
                }
            }
            '>' => {
                self.bump();
                if self.peek() == Some('=') {
                    self.bump();
                    Ok(Token::GreaterEquals)
                } else {
                    Ok(Token::Greater)
                }
            }
            '&' => {
                self.bump();
                if self.peek() == Some('&') {
                    self.bump();
                    Ok(Token::AndAnd)
                } else {
                    Ok(Token::Ampersand)
                }
            }
            '|' => {
                self.bump();
                if self.peek() == Some('|') {
                    self.bump();
                    Ok(Token::OrOr)
                } else {
                    Err(DaxError::Parse(format!(
                        "unexpected character '|' in {:?}",
                        self.input
                    )))
                }
            }
            '"' => {
                self.bump();
                let mut out = String::new();
                loop {
                    match self.peek() {
                        None => return Err(DaxError::Parse("unterminated string".into())),
                        Some('"') => {
                            self.bump();
                            if self.peek() == Some('"') {
                                self.bump();
                                out.push('"');
                                continue;
                            }
                            break;
                        }
                        Some(c) => {
                            out.push(c);
                            self.bump();
                        }
                    }
                }
                Ok(Token::String(out))
            }
            '\'' => {
                self.bump();
                let mut out = String::new();
                loop {
                    match self.peek() {
                        None => return Err(DaxError::Parse("unterminated identifier".into())),
                        Some('\'') => {
                            self.bump();
                            if self.peek() == Some('\'') {
                                self.bump();
                                out.push('\'');
                                continue;
                            }
                            break;
                        }
                        Some(c) => {
                            out.push(c);
                            self.bump();
                        }
                    }
                }
                Ok(Token::Identifier(out))
            }
            '[' => {
                self.bump();
                let mut out = String::new();
                while let Some(c) = self.peek() {
                    if c == ']' {
                        break;
                    }
                    out.push(c);
                    self.bump();
                }
                if self.peek() != Some(']') {
                    return Err(DaxError::Parse("unterminated bracket identifier".into()));
                }
                self.bump();
                Ok(Token::BracketIdentifier(out.trim().to_string()))
            }
            c if c.is_ascii_digit() || c == '.' => {
                let num_str = self.consume_while(|c| c.is_ascii_digit() || c == '.');
                let num: f64 = num_str
                    .parse()
                    .map_err(|_| DaxError::Parse(format!("invalid number {num_str:?}")))?;
                Ok(Token::Number(num))
            }
            c if is_ident_start(c) => {
                let ident = self.consume_while(is_ident_part);
                if ident.eq_ignore_ascii_case("VAR") {
                    Ok(Token::Var)
                } else if ident.eq_ignore_ascii_case("RETURN") {
                    Ok(Token::Return)
                } else if ident.eq_ignore_ascii_case("IN") {
                    Ok(Token::In)
                } else {
                    Ok(Token::Identifier(ident))
                }
            }
            other => Err(DaxError::Parse(format!(
                "unexpected character {other:?} in {:?}",
                self.input
            ))),
        }
    }
}

fn is_ident_start(c: char) -> bool {
    c.is_ascii_alphabetic() || c == '_' || c == '.'
}

fn is_ident_part(c: char) -> bool {
    is_ident_start(c) || c.is_ascii_digit()
}

struct Parser<'a> {
    lexer: Lexer<'a>,
    lookahead: Token,
}

impl<'a> Parser<'a> {
    fn new(input: &'a str) -> DaxResult<Self> {
        let mut lexer = Lexer::new(input);
        let lookahead = lexer.next_token()?;
        Ok(Self { lexer, lookahead })
    }

    fn bump(&mut self) -> DaxResult<Token> {
        let current = std::mem::replace(&mut self.lookahead, Token::Eof);
        self.lookahead = self.lexer.next_token()?;
        Ok(current)
    }

    fn expect(&mut self, token: Token) -> DaxResult<()> {
        if self.lookahead == token {
            self.bump()?;
            Ok(())
        } else {
            Err(DaxError::Parse(format!(
                "expected {token:?}, found {:?}",
                self.lookahead
            )))
        }
    }

    fn parse(&mut self) -> DaxResult<Expr> {
        let expr = self.parse_expr(0)?;
        if self.lookahead != Token::Eof {
            return Err(DaxError::Parse(format!(
                "unexpected token {:?}",
                self.lookahead
            )));
        }
        Ok(expr)
    }

    fn parse_expr(&mut self, min_prec: u8) -> DaxResult<Expr> {
        let mut left = self.parse_prefix()?;
        loop {
            let (op, prec) = match self.infix_binding_power() {
                Some(v) => v,
                None => break,
            };
            if prec < min_prec {
                break;
            }
            self.bump()?;
            let right = self.parse_expr(prec + 1)?;
            left = Expr::BinaryOp {
                op,
                left: Box::new(left),
                right: Box::new(right),
            };
        }
        Ok(left)
    }

    fn parse_prefix(&mut self) -> DaxResult<Expr> {
        match &self.lookahead {
            Token::Var => self.parse_let_expression(),
            Token::Minus => {
                self.bump()?;
                let expr = self.parse_expr(7)?;
                Ok(Expr::UnaryOp {
                    op: UnaryOp::Negate,
                    expr: Box::new(expr),
                })
            }
            Token::Number(n) => {
                let n = *n;
                self.bump()?;
                Ok(Expr::Number(n))
            }
            Token::String(s) => {
                let s = s.clone();
                self.bump()?;
                Ok(Expr::Text(s))
            }
            Token::Identifier(_) => self.parse_ident_like(),
            Token::BracketIdentifier(name) => {
                let name = name.clone();
                self.bump()?;
                Ok(Expr::Measure(name))
            }
            Token::LParen => {
                self.bump()?;
                let expr = self.parse_expr(0)?;
                self.expect(Token::RParen)?;
                Ok(expr)
            }
            Token::LBrace => self.parse_table_literal(),
            other => Err(DaxError::Parse(format!(
                "unexpected token in expression: {other:?}"
            ))),
        }
    }

    fn parse_let_expression(&mut self) -> DaxResult<Expr> {
        let mut bindings = Vec::new();
        while self.lookahead == Token::Var {
            self.bump()?; // VAR
            let name = match self.bump()? {
                Token::Identifier(name) => name,
                other => {
                    return Err(DaxError::Parse(format!(
                        "expected identifier after VAR, found {other:?}"
                    )))
                }
            };
            self.expect(Token::Equals)?;
            let expr = self.parse_expr(0)?;
            bindings.push((name, expr));
        }
        if bindings.is_empty() {
            return Err(DaxError::Parse("expected at least one VAR binding".into()));
        }
        self.expect(Token::Return)?;
        let body = self.parse_expr(0)?;
        Ok(Expr::Let {
            bindings,
            body: Box::new(body),
        })
    }

    fn parse_table_literal(&mut self) -> DaxResult<Expr> {
        self.expect(Token::LBrace)?;
        let mut rows: Vec<Vec<Expr>> = Vec::new();
        if self.lookahead != Token::RBrace {
            loop {
                let expr = self.parse_expr(0)?;
                if matches!(expr, Expr::TableLiteral { .. }) {
                    return Err(DaxError::Parse(
                        "nested table constructors are not supported".into(),
                    ));
                }
                rows.push(vec![expr]);

                match self.lookahead {
                    Token::Comma | Token::Semicolon => {
                        self.bump()?;
                        continue;
                    }
                    _ => break,
                }
            }
        }
        self.expect(Token::RBrace)?;
        Ok(Expr::TableLiteral { rows })
    }

    fn parse_ident_like(&mut self) -> DaxResult<Expr> {
        let Token::Identifier(ident) = self.bump()? else {
            unreachable!();
        };

        match &self.lookahead {
            Token::LParen => {
                self.bump()?;
                let mut args = Vec::new();
                if self.lookahead != Token::RParen {
                    loop {
                        args.push(self.parse_expr(0)?);
                        if matches!(self.lookahead, Token::Comma | Token::Semicolon) {
                            self.bump()?;
                            continue;
                        }
                        break;
                    }
                }
                self.expect(Token::RParen)?;
                Ok(Expr::Call { name: ident, args })
            }
            Token::BracketIdentifier(col) => {
                let col = col.clone();
                self.bump()?;
                Ok(Expr::ColumnRef {
                    table: ident,
                    column: col,
                })
            }
            _ => Ok(Expr::TableName(ident)),
        }
    }

    fn infix_binding_power(&self) -> Option<(BinaryOp, u8)> {
        match self.lookahead {
            Token::OrOr => Some((BinaryOp::Or, 1)),
            Token::AndAnd => Some((BinaryOp::And, 2)),
            Token::Equals => Some((BinaryOp::Equals, 3)),
            Token::NotEquals => Some((BinaryOp::NotEquals, 3)),
            Token::Less => Some((BinaryOp::Less, 3)),
            Token::LessEquals => Some((BinaryOp::LessEquals, 3)),
            Token::Greater => Some((BinaryOp::Greater, 3)),
            Token::GreaterEquals => Some((BinaryOp::GreaterEquals, 3)),
            Token::In => Some((BinaryOp::In, 3)),
            // DAX operator precedence (higher binds tighter):
            //   * /  >  + -  >  &  >  comparisons  >  &&  >  ||
            Token::Ampersand => Some((BinaryOp::Concat, 4)),
            Token::Plus => Some((BinaryOp::Add, 5)),
            Token::Minus => Some((BinaryOp::Subtract, 5)),
            Token::Star => Some((BinaryOp::Multiply, 6)),
            Token::Slash => Some((BinaryOp::Divide, 6)),
            _ => None,
        }
    }
}

pub fn parse(input: &str) -> DaxResult<Expr> {
    Parser::new(input)?.parse()
}
