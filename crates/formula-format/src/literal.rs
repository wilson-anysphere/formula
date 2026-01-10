pub(crate) fn render_text_section(section: &str, text: &str) -> String {
    let mut out = String::new();
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
                let _ = chars.next();
                out.push(' ');
            }
            '*' => {
                // fill: skip the next character entirely
                let _ = chars.next();
            }
            '@' => out.push_str(text),
            _ => out.push(ch),
        }
    }

    out
}

pub(crate) fn render_literal_segment(segment: &str) -> String {
    let mut out = String::new();
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
                let _ = chars.next();
                out.push(' ');
            }
            '*' => {
                let _ = chars.next();
            }
            _ => out.push(ch),
        }
    }

    out
}

fn currency_symbol_from_bracket(content: &str) -> Option<String> {
    let after = content.strip_prefix('$')?;
    let symbol = after.split_once('-').map(|(s, _)| s).unwrap_or(after);
    if symbol.is_empty() {
        None
    } else {
        Some(symbol.to_string())
    }
}
