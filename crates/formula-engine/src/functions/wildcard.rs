#[derive(Debug, Clone, Copy, PartialEq, Eq)]
enum Token {
    Star,
    QMark,
    Literal(char),
}

#[derive(Debug, Clone)]
pub(crate) struct WildcardPattern {
    tokens: Vec<Token>,
}

impl WildcardPattern {
    pub(crate) fn new(pattern: &str) -> Self {
        Self {
            tokens: tokenize_pattern(pattern),
        }
    }

    pub(crate) fn matches(&self, text: &str) -> bool {
        let text: Vec<char> = text.chars().map(|c| c.to_ascii_uppercase()).collect();
        wildcard_match_tokens(&self.tokens, &text)
    }
}

pub(crate) fn wildcard_match(pattern: &str, text: &str) -> bool {
    WildcardPattern::new(pattern).matches(text)
}

fn tokenize_pattern(pattern: &str) -> Vec<Token> {
    let mut tokens = Vec::new();
    let mut chars = pattern.chars().peekable();
    while let Some(c) = chars.next() {
        match c {
            '~' => {
                if let Some(next) = chars.next() {
                    tokens.push(Token::Literal(next.to_ascii_uppercase()));
                } else {
                    tokens.push(Token::Literal('~'));
                }
            }
            '*' => tokens.push(Token::Star),
            '?' => tokens.push(Token::QMark),
            other => tokens.push(Token::Literal(other.to_ascii_uppercase())),
        }
    }
    tokens
}

fn wildcard_match_tokens(pattern: &[Token], text: &[char]) -> bool {
    let mut pi = 0usize;
    let mut ti = 0usize;
    let mut star: Option<usize> = None;
    let mut star_text = 0usize;

    while ti < text.len() {
        if pi < pattern.len() {
            match pattern[pi] {
                Token::Literal(c) if c == text[ti] => {
                    pi += 1;
                    ti += 1;
                    continue;
                }
                Token::QMark => {
                    pi += 1;
                    ti += 1;
                    continue;
                }
                Token::Star => {
                    star = Some(pi);
                    pi += 1;
                    star_text = ti;
                    continue;
                }
                _ => {}
            }
        }

        if let Some(star_pos) = star {
            pi = star_pos + 1;
            star_text += 1;
            ti = star_text;
        } else {
            return false;
        }
    }

    while pi < pattern.len() && pattern[pi] == Token::Star {
        pi += 1;
    }

    pi == pattern.len()
}
