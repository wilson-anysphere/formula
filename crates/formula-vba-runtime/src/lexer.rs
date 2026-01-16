use crate::runtime::VbaError;

#[derive(Debug, Clone, PartialEq)]
pub enum TokenKind {
    Identifier(String),
    Number(f64),
    String(String),

    // Operators / punctuation
    LParen,
    RParen,
    Comma,
    Dot,
    Colon,
    ColonEq,
    Eq,
    Plus,
    Minus,
    Star,
    Slash,
    Backslash,
    Amp,
    Caret,
    Lt,
    Gt,
    Le,
    Ge,
    Ne,

    Newline,
    Eof,

    // Keywords (case-insensitive)
    Keyword(String),
}

#[derive(Debug, Clone)]
pub struct Token {
    pub kind: TokenKind,
    pub line: usize,
    pub col: usize,
}

#[derive(Clone)]
pub struct Lexer<'a> {
    chars: std::str::Chars<'a>,
    peeked: Option<char>,
    line: usize,
    col: usize,
}

impl<'a> Lexer<'a> {
    pub fn new(src: &'a str) -> Self {
        Self {
            chars: src.chars(),
            peeked: None,
            line: 1,
            col: 0,
        }
    }

    fn bump(&mut self) -> Option<char> {
        let ch = if let Some(ch) = self.peeked.take() {
            Some(ch)
        } else {
            self.chars.next()
        };
        if let Some(ch) = ch {
            if ch == '\n' {
                self.line += 1;
                self.col = 0;
            } else {
                self.col += 1;
            }
        }
        ch
    }

    fn peek(&mut self) -> Option<char> {
        if self.peeked.is_none() {
            self.peeked = self.chars.next();
        }
        self.peeked
    }

    fn is_ident_start(ch: char) -> bool {
        ch.is_ascii_alphabetic() || ch == '_'
    }

    fn is_ident_continue(ch: char) -> bool {
        ch.is_ascii_alphanumeric() || ch == '_'
    }

    fn lex_number(&mut self, first: char, line: usize, col: usize) -> Result<Token, VbaError> {
        let mut buf = String::new();
        buf.push(first);
        while let Some(ch) = self.peek() {
            if ch.is_ascii_digit() || ch == '.' {
                buf.push(self.bump().unwrap());
            } else {
                break;
            }
        }
        let value = buf.parse::<f64>().map_err(|_| {
            VbaError::Parse(format!("Invalid number literal `{buf}` at {line}:{col}"))
        })?;
        Ok(Token {
            kind: TokenKind::Number(value),
            line,
            col,
        })
    }

    fn lex_identifier(&mut self, first: char, line: usize, col: usize) -> Result<Token, VbaError> {
        let mut buf = String::new();
        buf.push(first);
        while let Some(ch) = self.peek() {
            if Self::is_ident_continue(ch) {
                buf.push(self.bump().unwrap());
            } else {
                break;
            }
        }

        // VBA keywords we care about. Use keyword tokens to simplify parsing.
        //
        // Avoid allocating a lowercased copy for all identifiers: only lower-case in place when we
        // actually emit a keyword token.
        fn is_keyword(s: &str) -> bool {
            s.eq_ignore_ascii_case("sub")
                || s.eq_ignore_ascii_case("function")
                || s.eq_ignore_ascii_case("end")
                || s.eq_ignore_ascii_case("if")
                || s.eq_ignore_ascii_case("then")
                || s.eq_ignore_ascii_case("else")
                || s.eq_ignore_ascii_case("elseif")
                || s.eq_ignore_ascii_case("for")
                || s.eq_ignore_ascii_case("each")
                || s.eq_ignore_ascii_case("in")
                || s.eq_ignore_ascii_case("to")
                || s.eq_ignore_ascii_case("step")
                || s.eq_ignore_ascii_case("next")
                || s.eq_ignore_ascii_case("dim")
                || s.eq_ignore_ascii_case("const")
                || s.eq_ignore_ascii_case("as")
                || s.eq_ignore_ascii_case("integer")
                || s.eq_ignore_ascii_case("long")
                || s.eq_ignore_ascii_case("string")
                || s.eq_ignore_ascii_case("date")
                || s.eq_ignore_ascii_case("boolean")
                || s.eq_ignore_ascii_case("is")
                || s.eq_ignore_ascii_case("byval")
                || s.eq_ignore_ascii_case("byref")
                || s.eq_ignore_ascii_case("set")
                || s.eq_ignore_ascii_case("on")
                || s.eq_ignore_ascii_case("error")
                || s.eq_ignore_ascii_case("resume")
                || s.eq_ignore_ascii_case("goto")
                || s.eq_ignore_ascii_case("exit")
                || s.eq_ignore_ascii_case("do")
                || s.eq_ignore_ascii_case("while")
                || s.eq_ignore_ascii_case("loop")
                || s.eq_ignore_ascii_case("until")
                || s.eq_ignore_ascii_case("wend")
                || s.eq_ignore_ascii_case("select")
                || s.eq_ignore_ascii_case("case")
                || s.eq_ignore_ascii_case("with")
                || s.eq_ignore_ascii_case("call")
                || s.eq_ignore_ascii_case("true")
                || s.eq_ignore_ascii_case("false")
                || s.eq_ignore_ascii_case("nothing")
                || s.eq_ignore_ascii_case("and")
                || s.eq_ignore_ascii_case("or")
                || s.eq_ignore_ascii_case("mod")
                || s.eq_ignore_ascii_case("not")
                || s.eq_ignore_ascii_case("new")
                || s.eq_ignore_ascii_case("rem")
                || s.eq_ignore_ascii_case("private")
                || s.eq_ignore_ascii_case("public")
                || s.eq_ignore_ascii_case("option")
                || s.eq_ignore_ascii_case("attribute")
                || s.eq_ignore_ascii_case("explicit")
                || s.eq_ignore_ascii_case("debug")
                || s.eq_ignore_ascii_case("print")
        }

        if is_keyword(&buf) {
            buf.make_ascii_lowercase();
            return Ok(Token {
                kind: TokenKind::Keyword(buf),
                line,
                col,
            });
        }

        Ok(Token {
            kind: TokenKind::Identifier(buf),
            line,
            col,
        })
    }

    fn lex_string(&mut self, line: usize, col: usize) -> Result<Token, VbaError> {
        let mut buf = String::new();
        loop {
            match self.bump() {
                Some('"') => {
                    // doubled quote is an escape
                    if self.peek() == Some('"') {
                        self.bump();
                        buf.push('"');
                        continue;
                    }
                    break;
                }
                Some(ch) => buf.push(ch),
                None => {
                    return Err(VbaError::Parse(format!(
                        "Unterminated string literal at {line}:{col}"
                    )))
                }
            }
        }
        Ok(Token {
            kind: TokenKind::String(buf),
            line,
            col,
        })
    }

    fn skip_whitespace(&mut self) {
        while let Some(ch) = self.peek() {
            if ch == ' ' || ch == '\t' || ch == '\r' {
                self.bump();
            } else {
                break;
            }
        }
    }

    fn skip_comment(&mut self) {
        while let Some(ch) = self.peek() {
            self.bump();
            if ch == '\n' {
                break;
            }
        }
    }

    pub fn next_token(&mut self) -> Result<Token, VbaError> {
        self.skip_whitespace();
        let line = self.line;
        let col = self.col + 1;
        match self.bump() {
            Some('\n') => Ok(Token {
                kind: TokenKind::Newline,
                line,
                col,
            }),
            Some('\'') => {
                self.skip_comment();
                Ok(Token {
                    kind: TokenKind::Newline,
                    line,
                    col,
                })
            }
            Some('"') => self.lex_string(line, col),
            Some(ch) if ch.is_ascii_digit() => self.lex_number(ch, line, col),
            Some(ch) if Self::is_ident_start(ch) => {
                // Handle `Rem` comments by turning the whole rest of the line into a newline.
                let tok = self.lex_identifier(ch, line, col)?;
                if matches!(tok.kind, TokenKind::Keyword(ref k) if k == "rem") {
                    self.skip_comment();
                    Ok(Token {
                        kind: TokenKind::Newline,
                        line,
                        col,
                    })
                } else {
                    Ok(tok)
                }
            }
            Some('(') => Ok(Token {
                kind: TokenKind::LParen,
                line,
                col,
            }),
            Some(')') => Ok(Token {
                kind: TokenKind::RParen,
                line,
                col,
            }),
            Some(',') => Ok(Token {
                kind: TokenKind::Comma,
                line,
                col,
            }),
            Some('.') => Ok(Token {
                kind: TokenKind::Dot,
                line,
                col,
            }),
            Some(':') => Ok(Token {
                kind: if self.peek() == Some('=') {
                    self.bump();
                    TokenKind::ColonEq
                } else {
                    TokenKind::Colon
                },
                line,
                col,
            }),
            Some('=') => Ok(Token {
                kind: TokenKind::Eq,
                line,
                col,
            }),
            Some('+') => Ok(Token {
                kind: TokenKind::Plus,
                line,
                col,
            }),
            Some('-') => Ok(Token {
                kind: TokenKind::Minus,
                line,
                col,
            }),
            Some('*') => Ok(Token {
                kind: TokenKind::Star,
                line,
                col,
            }),
            Some('/') => Ok(Token {
                kind: TokenKind::Slash,
                line,
                col,
            }),
            Some('\\') => Ok(Token {
                kind: TokenKind::Backslash,
                line,
                col,
            }),
            Some('&') => Ok(Token {
                kind: TokenKind::Amp,
                line,
                col,
            }),
            Some('^') => Ok(Token {
                kind: TokenKind::Caret,
                line,
                col,
            }),
            Some('<') => {
                if self.peek() == Some('=') {
                    self.bump();
                    Ok(Token {
                        kind: TokenKind::Le,
                        line,
                        col,
                    })
                } else if self.peek() == Some('>') {
                    self.bump();
                    Ok(Token {
                        kind: TokenKind::Ne,
                        line,
                        col,
                    })
                } else {
                    Ok(Token {
                        kind: TokenKind::Lt,
                        line,
                        col,
                    })
                }
            }
            Some('>') => {
                if self.peek() == Some('=') {
                    self.bump();
                    Ok(Token {
                        kind: TokenKind::Ge,
                        line,
                        col,
                    })
                } else {
                    Ok(Token {
                        kind: TokenKind::Gt,
                        line,
                        col,
                    })
                }
            }
            None => Ok(Token {
                kind: TokenKind::Eof,
                line,
                col,
            }),
            Some(other) => Err(VbaError::Parse(format!(
                "Unexpected character `{other}` at {line}:{col}"
            ))),
        }
    }
}
