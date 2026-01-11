use crate::{LiteralLayoutHint, LiteralLayoutOp};

#[derive(Debug, Clone, PartialEq, Eq)]
pub(crate) struct RenderedText {
    pub text: String,
    layout: LiteralLayoutHint,
}

impl RenderedText {
    pub(crate) fn new(text: String) -> Self {
        Self {
            text,
            layout: LiteralLayoutHint::default(),
        }
    }

    pub(crate) fn layout_hint(&self) -> Option<LiteralLayoutHint> {
        if self.layout.ops.is_empty() {
            None
        } else {
            Some(self.layout.clone())
        }
    }

    pub(crate) fn push_str(&mut self, s: &str) {
        self.text.push_str(s);
    }

    pub(crate) fn push(&mut self, ch: char) {
        self.text.push(ch);
    }

    pub(crate) fn push_layout_op(&mut self, op: LiteralLayoutOp) {
        self.layout.ops.push(op);
    }

    pub(crate) fn extend(&mut self, other: RenderedText) {
        let offset = self.text.len();
        self.text.push_str(&other.text);
        if !other.layout.ops.is_empty() {
            self.layout.ops.extend(other.layout.ops.into_iter().map(|op| match op {
                LiteralLayoutOp::Underscore { byte_index, width_of } => LiteralLayoutOp::Underscore {
                    byte_index: offset + byte_index,
                    width_of,
                },
                LiteralLayoutOp::Fill { byte_index, fill_with } => LiteralLayoutOp::Fill {
                    byte_index: offset + byte_index,
                    fill_with,
                },
            }));
        }
    }

    pub(crate) fn prepend_char(&mut self, ch: char) {
        let mut buf = String::new();
        buf.push(ch);
        buf.push_str(&self.text);
        self.text = buf;
        if !self.layout.ops.is_empty() {
            let shift = ch.len_utf8();
            for op in &mut self.layout.ops {
                match op {
                    LiteralLayoutOp::Underscore { byte_index, .. } | LiteralLayoutOp::Fill { byte_index, .. } => {
                        *byte_index += shift;
                    }
                }
            }
        }
    }
}

pub(crate) fn render_text_section(section: &str, text: &str) -> RenderedText {
    let mut out = RenderedText::new(String::new());
    let mut in_quotes = false;
    let mut chars = section.chars().peekable();

    while let Some(ch) = chars.next() {
        if in_quotes {
            if ch == '"' {
                in_quotes = false;
            } else {
                out.push(ch);
            }
            continue;
        }

        match ch {
            '"' => in_quotes = true,
            '\\' => {
                if let Some(next) = chars.next() {
                    out.push(next);
                }
            }
            '[' => {
                // Bracket tokens (colors, locales, currencies, conditions).
                // For display text we ignore most bracket tokens, except currency
                // markers of the form `[$$-409]` which render a currency symbol.
                let mut content = String::new();
                let mut closed = false;
                while let Some(c) = chars.next() {
                    if c == ']' {
                        closed = true;
                        break;
                    }
                    content.push(c);
                }

                if closed {
                    if let Some(symbol) = currency_symbol_from_bracket(&content) {
                        out.push_str(&symbol);
                    }
                } else {
                    // No closing `]`: treat as literal.
                    out.push('[');
                    out.push_str(&content);
                }
            }
            '_' => {
                // underscore: skip next character, output a space (approximation)
                if let Some(width_of) = chars.next() {
                    let idx = out.text.len();
                    out.push_layout_op(LiteralLayoutOp::Underscore {
                        byte_index: idx,
                        width_of,
                    });
                } else {
                    // No following char; Excel treats `_` as literal.
                    out.push('_');
                    continue;
                }
                out.push(' ');
            }
            '*' => {
                if let Some(fill_with) = chars.next() {
                    out.push_layout_op(LiteralLayoutOp::Fill {
                        byte_index: out.text.len(),
                        fill_with,
                    });
                } else {
                    out.push('*');
                }
            }
            '@' => out.push_str(text),
            _ => out.push(ch),
        }
    }

    out
}

pub(crate) fn render_literal_segment(segment: &str) -> RenderedText {
    let mut out = RenderedText::new(String::new());
    let mut in_quotes = false;
    let mut chars = segment.chars().peekable();

    while let Some(ch) = chars.next() {
        if in_quotes {
            if ch == '"' {
                in_quotes = false;
            } else {
                out.push(ch);
            }
            continue;
        }

        match ch {
            '"' => in_quotes = true,
            '\\' => {
                if let Some(next) = chars.next() {
                    out.push(next);
                }
            }
            '[' => {
                let mut content = String::new();
                let mut closed = false;
                while let Some(c) = chars.next() {
                    if c == ']' {
                        closed = true;
                        break;
                    }
                    content.push(c);
                }

                if closed {
                    if let Some(symbol) = currency_symbol_from_bracket(&content) {
                        out.push_str(&symbol);
                    }
                    // All other bracket tokens are ignored.
                } else {
                    out.push('[');
                    out.push_str(&content);
                }
            }
            '_' => {
                if let Some(width_of) = chars.next() {
                    let idx = out.text.len();
                    out.push_layout_op(LiteralLayoutOp::Underscore {
                        byte_index: idx,
                        width_of,
                    });
                    out.push(' ');
                } else {
                    out.push('_');
                }
            }
            '*' => {
                if let Some(fill_with) = chars.next() {
                    out.push_layout_op(LiteralLayoutOp::Fill {
                        byte_index: out.text.len(),
                        fill_with,
                    });
                } else {
                    out.push('*');
                }
            }
            _ => out.push(ch),
        }
    }

    out
}

fn currency_symbol_from_bracket(content: &str) -> Option<String> {
    let after = content.strip_prefix('$')?;
    let (symbol, locale) = after.split_once('-').unwrap_or((after, ""));
    if !symbol.is_empty() {
        return Some(symbol.to_string());
    }
    let locale = locale.trim();
    if locale.is_empty() {
        return None;
    }
    let lcid = u32::from_str_radix(locale, 16).ok()?;
    default_currency_symbol_for_lcid(lcid).map(|s| s.to_string())
}

fn default_currency_symbol_for_lcid(lcid: u32) -> Option<&'static str> {
    // This is intentionally a small, deterministic mapping covering the most common LCIDs seen in
    // OOXML format codes. It can be expanded as we encounter additional locales in the corpus.
    Some(match lcid {
        0x0409 | 0x1009 => "$", // en-US, en-CA
        0x0809 => "£",         // en-GB
        0x0407 | 0x040c | 0x0410 | 0x040a | 0x0413 => "€", // de-DE, fr-FR, it-IT, es-ES, nl-NL
        0x0411 | 0x0804 => "¥", // ja-JP, zh-CN
        0x0412 => "₩",         // ko-KR
        _ => return None,
    })
}
