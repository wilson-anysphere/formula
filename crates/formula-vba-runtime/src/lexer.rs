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
    Eq,
    Plus,
    Minus,
    Star,
    Slash,
    Amp,
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

        let lower = buf.to_ascii_lowercase();
        // VBA keywords we care about. Use keyword token to simplify parsing.
        let is_keyword = matches!(
            lower.as_str(),
            "sub"
                | "function"
                | "end"
                | "if"
                | "then"
                | "else"
                | "elseif"
                | "for"
                | "to"
                | "step"
                | "next"
                | "dim"
                | "as"
                | "byval"
                | "byref"
                | "set"
                | "on"
                | "error"
                | "resume"
                | "goto"
                | "exit"
                | "do"
                | "while"
                | "loop"
                | "until"
                | "wend"
                | "call"
                | "true"
                | "false"
                | "nothing"
                | "and"
                | "or"
                | "not"
                | "new"
                | "rem"
                | "private"
                | "public"
                | "option"
                | "explicit"
                | "debug"
                | "print"
        );

        if is_keyword {
            Ok(Token {
                kind: TokenKind::Keyword(lower),
                line,
                col,
            })
        } else {
            Ok(Token {
                kind: TokenKind::Identifier(buf),
                line,
                col,
            })
        }
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
                kind: TokenKind::Colon,
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
            Some('&') => Ok(Token {
                kind: TokenKind::Amp,
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
