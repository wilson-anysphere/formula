use std::fmt;

#[derive(Debug, Clone, PartialEq, Eq)]
pub struct ParseError {
    pub message: String,
}

impl fmt::Display for ParseError {
    fn fmt(&self, f: &mut fmt::Formatter<'_>) -> fmt::Result {
        write!(f, "{}", self.message)
    }
}

impl std::error::Error for ParseError {}

#[derive(Debug, Clone, Copy, PartialEq, Eq)]
enum CmpOp {
    Lt,
    Le,
    Gt,
    Ge,
    Eq,
    Ne,
}

#[derive(Debug, Clone, Copy, PartialEq)]
struct Condition {
    op: CmpOp,
    rhs: f64,
}

impl Condition {
    fn matches(self, v: f64) -> bool {
        match self.op {
            CmpOp::Lt => v < self.rhs,
            CmpOp::Le => v <= self.rhs,
            CmpOp::Gt => v > self.rhs,
            CmpOp::Ge => v >= self.rhs,
            CmpOp::Eq => v == self.rhs,
            CmpOp::Ne => v != self.rhs,
        }
    }
}

#[derive(Debug, Clone, PartialEq)]
struct Section {
    raw: String,
    condition: Option<Condition>,
}

/// Parsed Excel number format code split into `;`-delimited sections.
#[derive(Debug, Clone, PartialEq)]
pub struct FormatCode {
    sections: Vec<Section>,
}

#[derive(Debug, Clone, Copy)]
pub(crate) struct SelectedSection<'a> {
    pub pattern: &'a str,
    pub auto_negative_sign: bool,
}

impl FormatCode {
    pub fn general() -> Self {
        Self {
            sections: vec![Section {
                raw: "General".to_string(),
                condition: None,
            }],
        }
    }

    pub fn parse(code: &str) -> Result<Self, ParseError> {
        let sections = split_sections(code)
            .into_iter()
            .map(|s| parse_section(&s))
            .collect::<Result<Vec<_>, _>>()?;

        Ok(Self { sections })
    }

    pub(crate) fn text_section(&self) -> Option<&str> {
        if self.sections.len() >= 4 {
            Some(self.sections[3].raw.as_str())
        } else {
            None
        }
    }

    pub(crate) fn select_section_for_number(&self, v: f64) -> SelectedSection<'_> {
        // If any section has a condition, Excel evaluates conditions in-order,
        // then uses the first unconditional section as an "else".
        if self.sections.iter().any(|s| s.condition.is_some()) {
            let mut fallback: Option<&Section> = None;
            for section in &self.sections {
                match section.condition {
                    Some(cond) => {
                        if cond.matches(v) {
                            return SelectedSection {
                                pattern: section.raw.as_str(),
                                auto_negative_sign: false,
                            };
                        }
                    }
                    None => {
                        if fallback.is_none() {
                            fallback = Some(section);
                        }
                    }
                }
            }
            let section = fallback.unwrap_or_else(|| &self.sections[0]);
            return SelectedSection {
                pattern: section.raw.as_str(),
                auto_negative_sign: false,
            };
        }

        let section_count = self.sections.len();

        // Excel semantics for 1-4 sections:
        // 1: positive; used for all numbers, negatives get a '-' automatically.
        // 2: positive;negative
        // 3: positive;negative;zero
        // 4: positive;negative;zero;text
        if v < 0.0 {
            if section_count >= 2 {
                SelectedSection {
                    pattern: self.sections[1].raw.as_str(),
                    auto_negative_sign: false,
                }
            } else {
                SelectedSection {
                    pattern: self.sections[0].raw.as_str(),
                    auto_negative_sign: true,
                }
            }
        } else if v == 0.0 {
            if section_count >= 3 {
                SelectedSection {
                    pattern: self.sections[2].raw.as_str(),
                    auto_negative_sign: false,
                }
            } else {
                SelectedSection {
                    pattern: self.sections[0].raw.as_str(),
                    auto_negative_sign: false,
                }
            }
        } else {
            SelectedSection {
                pattern: self.sections[0].raw.as_str(),
                auto_negative_sign: false,
            }
        }
    }
}

fn split_sections(code: &str) -> Vec<String> {
    let mut sections = Vec::new();
    let mut buf = String::new();
    let mut in_quotes = false;
    let mut chars = code.chars().peekable();

    while let Some(ch) = chars.next() {
        match ch {
            '"' => {
                in_quotes = !in_quotes;
                buf.push(ch);
            }
            '\\' => {
                buf.push(ch);
                if let Some(next) = chars.next() {
                    buf.push(next);
                }
            }
            ';' if !in_quotes => {
                sections.push(buf);
                buf = String::new();
            }
            _ => buf.push(ch),
        }
    }

    sections.push(buf);
    sections
}

fn parse_section(input: &str) -> Result<Section, ParseError> {
    let mut rest = input.trim().to_string();
    let mut condition: Option<Condition> = None;
    let mut leading_literal = String::new();

    // Strip leading bracketed components like colors, locale tags, currencies,
    // and conditions. Conditions are of the form `[>=100]`.
    loop {
        let Some(stripped) = rest.strip_prefix('[') else {
            break;
        };
        let Some(end) = stripped.find(']') else {
            break;
        };
        let content = &stripped[..end];

        // Do not strip elapsed time tokens `[h]`, `[m]`, `[s]` which are part of
        // date/time formatting.
        let lower = content.to_ascii_lowercase();
        if matches!(lower.as_str(), "h" | "hh" | "m" | "mm" | "s" | "ss") {
            break;
        }

        if let Some(symbol) = currency_symbol_from_bracket(content) {
            leading_literal.push_str(&symbol);
        }

        if condition.is_none() {
            if let Some(cond) = parse_condition(content) {
                condition = Some(cond);
            }
        }

        rest = stripped[end + 1..].to_string();
        rest = rest.trim_start().to_string();
    }

    Ok(Section {
        raw: format!("{leading_literal}{rest}"),
        condition,
    })
}

fn parse_condition(content: &str) -> Option<Condition> {
    let (op, rhs_str) = if let Some(rest) = content.strip_prefix(">=") {
        (CmpOp::Ge, rest)
    } else if let Some(rest) = content.strip_prefix("<=") {
        (CmpOp::Le, rest)
    } else if let Some(rest) = content.strip_prefix("<>") {
        (CmpOp::Ne, rest)
    } else if let Some(rest) = content.strip_prefix('>') {
        (CmpOp::Gt, rest)
    } else if let Some(rest) = content.strip_prefix('<') {
        (CmpOp::Lt, rest)
    } else if let Some(rest) = content.strip_prefix('=') {
        (CmpOp::Eq, rest)
    } else {
        return None;
    };

    let rhs = rhs_str.trim().parse::<f64>().ok()?;
    Some(Condition { op, rhs })
}

fn currency_symbol_from_bracket(content: &str) -> Option<String> {
    // Currency/locale tags are encoded as `[$$-409]` where the currency symbol
    // is between the first `$` and the optional `-` locale suffix.
    let after = content.strip_prefix('$')?;
    let symbol = after.split_once('-').map(|(s, _)| s).unwrap_or(after);
    if symbol.is_empty() {
        None
    } else {
        Some(symbol.to_string())
    }
}
