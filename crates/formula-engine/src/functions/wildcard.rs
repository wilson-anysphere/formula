#[derive(Debug, Clone, Copy, PartialEq, Eq)]
enum Token {
    Star,
    QMark,
    Literal(char),
}

#[derive(Debug, Clone, PartialEq, Eq)]
pub(crate) struct WildcardPattern {
    tokens: Vec<Token>,
    has_wildcards: bool,
}

impl WildcardPattern {
    pub(crate) fn new(pattern: &str) -> Self {
        let tokens = tokenize_pattern(pattern);
        let has_wildcards = tokens
            .iter()
            .any(|t| matches!(t, Token::Star | Token::QMark));
        Self {
            tokens,
            has_wildcards,
        }
    }

    pub(crate) fn matches(&self, text: &str) -> bool {
        if !self.has_wildcards {
            return literal_tokens_match_text_unicode_case_insensitive(&self.tokens, text);
        }

        // Excel wildcard matching is case-insensitive. Use Unicode uppercasing so patterns like
        // "straße" match "STRASSE" (ß uppercases to SS), but keep an ASCII fast-path to avoid
        // the overhead for the common case.
        let text: Vec<char> = if text.is_ascii() {
            let mut out = Vec::with_capacity(text.len());
            if !text.as_bytes().iter().any(|b| b.is_ascii_lowercase()) {
                out.extend(text.chars());
            } else {
                out.extend(text.chars().map(|c| c.to_ascii_uppercase()));
            }
            out
        } else {
            text.chars().flat_map(|c| c.to_uppercase()).collect()
        };
        wildcard_match_tokens(&self.tokens, &text)
    }

    /// Like [`WildcardPattern::matches`], but assumes `text` is already case-folded to the same
    /// representation used when tokenizing the pattern.
    pub(crate) fn matches_folded(&self, text: &str) -> bool {
        let text: Vec<char> = text.chars().collect();
        wildcard_match_tokens(&self.tokens, &text)
    }

    pub(crate) fn has_wildcards(&self) -> bool {
        self.has_wildcards
    }

    /// Returns the literal representation of this pattern with wildcard operators expressed as
    /// `*` / `?` and escape sequences resolved.
    pub(crate) fn literal_pattern(&self) -> String {
        self.tokens
            .iter()
            .map(|t| match t {
                Token::Star => '*',
                Token::QMark => '?',
                Token::Literal(c) => *c,
            })
            .collect()
    }
}

fn tokenize_pattern(pattern: &str) -> Vec<Token> {
    let mut tokens = Vec::new();
    let mut chars = pattern.chars().peekable();
    while let Some(c) = chars.next() {
        match c {
            '~' => {
                let Some(&next) = chars.peek() else {
                    tokens.push(Token::Literal('~'));
                    continue;
                };

                if matches!(next, '*' | '?' | '~') {
                    let _ = chars.next();
                    tokens.push(Token::Literal(next));
                } else {
                    tokens.push(Token::Literal('~'));
                }
            }
            '*' => tokens.push(Token::Star),
            '?' => tokens.push(Token::QMark),
            other if other.is_ascii() => tokens.push(Token::Literal(other.to_ascii_uppercase())),
            other => tokens.extend(other.to_uppercase().map(Token::Literal)),
        }
    }
    tokens
}

struct FoldedUppercaseChars<'a> {
    chars: std::str::Chars<'a>,
    pending: Option<std::char::ToUppercase>,
    ascii_fast_path: bool,
    ascii_needs_uppercasing: bool,
}

impl<'a> FoldedUppercaseChars<'a> {
    fn new(s: &'a str) -> Self {
        let ascii_fast_path = s.is_ascii();
        let ascii_needs_uppercasing =
            ascii_fast_path && s.as_bytes().iter().any(|b| b.is_ascii_lowercase());
        Self {
            chars: s.chars(),
            pending: None,
            ascii_fast_path,
            ascii_needs_uppercasing,
        }
    }
}

impl Iterator for FoldedUppercaseChars<'_> {
    type Item = char;

    fn next(&mut self) -> Option<char> {
        loop {
            if let Some(pending) = &mut self.pending {
                if let Some(ch) = pending.next() {
                    return Some(ch);
                }
                self.pending = None;
            }

            let ch = self.chars.next()?;
            if self.ascii_fast_path {
                if self.ascii_needs_uppercasing {
                    return Some(ch.to_ascii_uppercase());
                }
                return Some(ch);
            }
            self.pending = Some(ch.to_uppercase());
        }
    }
}

fn literal_tokens_match_text_unicode_case_insensitive(tokens: &[Token], text: &str) -> bool {
    let mut ti = 0usize;
    for folded in FoldedUppercaseChars::new(text) {
        match tokens.get(ti) {
            Some(Token::Literal(c)) if *c == folded => {
                ti += 1;
            }
            _ => return false,
        }
    }
    ti == tokens.len()
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

#[cfg(test)]
mod tests {
    use super::*;

    #[test]
    fn matches_literal_without_wildcards_case_insensitive() {
        let pat = WildcardPattern::new("AbC");
        assert!(!pat.has_wildcards());
        assert!(pat.matches("abc"));
        assert!(pat.matches("ABC"));
        assert!(!pat.matches("abcd"));
    }

    #[test]
    fn matches_literal_handles_unicode_uppercase_expansion() {
        let pat = WildcardPattern::new("straße");
        assert!(!pat.has_wildcards());
        assert!(pat.matches("STRASSE"));
        assert!(pat.matches("straße"));
        assert!(!pat.matches("S"));
    }

    #[test]
    fn matches_literal_respects_tilde_escapes() {
        let pat = WildcardPattern::new("a~*b");
        assert!(!pat.has_wildcards());
        assert!(pat.matches("a*b"));
        assert!(!pat.matches("ab"));
    }
}
