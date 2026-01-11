use std::fmt;

use crate::{ColorOverride, Locale};

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
    color: Option<ColorOverride>,
    locale_override: Option<Locale>,
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
    pub color: Option<ColorOverride>,
    pub locale_override: Option<Locale>,
}

impl FormatCode {
    pub fn general() -> Self {
        Self {
            sections: vec![Section {
                raw: "General".to_string(),
                condition: None,
                color: None,
                locale_override: None,
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

    pub(crate) fn select_section_for_text(&self) -> (Option<&str>, Option<ColorOverride>) {
        if self.sections.len() >= 4 {
            let section = &self.sections[3];
            return (Some(section.raw.as_str()), section.color);
        }

        // Excel's built-in Text format is `@` and has only one section.
        // When a format code has no explicit 4th (text) section, Excel still
        // treats sections containing `@` as eligible for text rendering.
        self.sections
            .iter()
            .find(|s| contains_at_placeholder(&s.raw))
            .map(|s| (Some(s.raw.as_str()), s.color))
            .unwrap_or((None, None))
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
                                color: section.color,
                                locale_override: section.locale_override,
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
                color: section.color,
                locale_override: section.locale_override,
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
                    color: self.sections[1].color,
                    locale_override: self.sections[1].locale_override,
                }
            } else {
                SelectedSection {
                    pattern: self.sections[0].raw.as_str(),
                    auto_negative_sign: true,
                    color: self.sections[0].color,
                    locale_override: self.sections[0].locale_override,
                }
            }
        } else if v == 0.0 {
            if section_count >= 3 {
                SelectedSection {
                    pattern: self.sections[2].raw.as_str(),
                    auto_negative_sign: false,
                    color: self.sections[2].color,
                    locale_override: self.sections[2].locale_override,
                }
            } else {
                SelectedSection {
                    pattern: self.sections[0].raw.as_str(),
                    auto_negative_sign: false,
                    color: self.sections[0].color,
                    locale_override: self.sections[0].locale_override,
                }
            }
        } else {
            SelectedSection {
                pattern: self.sections[0].raw.as_str(),
                auto_negative_sign: false,
                color: self.sections[0].color,
                locale_override: self.sections[0].locale_override,
            }
        }
    }
}

fn contains_at_placeholder(pattern: &str) -> bool {
    let mut in_quotes = false;
    let mut escape = false;
    let mut in_brackets = false;

    for ch in pattern.chars() {
        if escape {
            escape = false;
            continue;
        }

        if in_quotes {
            if ch == '"' {
                in_quotes = false;
            }
            continue;
        }

        if in_brackets {
            if ch == ']' {
                in_brackets = false;
            }
            continue;
        }

        match ch {
            '"' => in_quotes = true,
            '\\' => escape = true,
            '[' => in_brackets = true,
            '@' => return true,
            _ => {}
        }
    }

    false
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
    let mut rest = input;
    let mut condition: Option<Condition> = None;
    let mut color: Option<ColorOverride> = None;
    let mut locale_override: Option<Locale> = None;

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
        if !lower.is_empty()
            && (lower.chars().all(|c| c == 'h')
                || lower.chars().all(|c| c == 'm')
                || lower.chars().all(|c| c == 's'))
        {
            break;
        }

        if color.is_none() {
            if let Some(parsed) = parse_color_token(&lower) {
                color = Some(parsed);
            }
        }

        if locale_override.is_none() {
            if let Some(locale) = parse_locale_override(content) {
                locale_override = Some(locale);
            }
        }

        if condition.is_none() {
            if let Some(cond) = parse_condition(content) {
                condition = Some(cond);
            }
        }

        rest = &stripped[end + 1..];
        rest = rest.trim_start();
    }

    Ok(Section {
        raw: input.to_string(),
        condition,
        color,
        locale_override,
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

fn parse_color_token(lower: &str) -> Option<ColorOverride> {
    Some(match lower {
        "black" => ColorOverride::Argb(0xFF000000),
        "white" => ColorOverride::Argb(0xFFFFFFFF),
        "red" => ColorOverride::Argb(0xFFFF0000),
        "green" => ColorOverride::Argb(0xFF00FF00),
        "blue" => ColorOverride::Argb(0xFF0000FF),
        "cyan" => ColorOverride::Argb(0xFF00FFFF),
        "magenta" => ColorOverride::Argb(0xFFFF00FF),
        "yellow" => ColorOverride::Argb(0xFFFFFF00),
        _ => {
            let rest = lower.strip_prefix("color")?;
            let idx: u8 = rest.trim().parse().ok()?;
            ColorOverride::Indexed(idx)
        }
    })
}

fn parse_locale_override(content: &str) -> Option<Locale> {
    // Locale/currency tags are encoded as `[$$-409]` where the locale is the hex LCID suffix.
    let after = content.strip_prefix('$')?;
    let (_, locale) = after.split_once('-')?;
    let lcid = u32::from_str_radix(locale.trim(), 16).ok()?;
    locale_for_lcid(lcid)
}

fn locale_for_lcid(lcid: u32) -> Option<Locale> {
    // Deterministic subset covering the most common locales seen in OOXML format codes.
    Some(match lcid {
        0x0409 | 0x0809 | 0x1009 => Locale::en_us(), // en-US, en-GB, en-CA
        0x0407 => Locale::de_de(),                  // de-DE
        0x040c => Locale::fr_fr(),                  // fr-FR
        0x0410 => Locale::it_it(),                  // it-IT
        0x040a => Locale::es_es(),                  // es-ES
        _ => return None,
    })
}
