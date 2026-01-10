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
