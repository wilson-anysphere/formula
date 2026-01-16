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
        if !content.is_empty()
            && (content.chars().all(|c| matches!(c, 'h' | 'H'))
                || content.chars().all(|c| matches!(c, 'm' | 'M'))
                || content.chars().all(|c| matches!(c, 's' | 'S')))
        {
            break;
        }

        if color.is_none() {
            if let Some(parsed) = parse_color_token(content) {
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

fn parse_color_token(content: &str) -> Option<ColorOverride> {
    Some(if content.eq_ignore_ascii_case("black") {
        ColorOverride::Argb(0xFF000000)
    } else if content.eq_ignore_ascii_case("white") {
        ColorOverride::Argb(0xFFFFFFFF)
    } else if content.eq_ignore_ascii_case("red") {
        ColorOverride::Argb(0xFFFF0000)
    } else if content.eq_ignore_ascii_case("green") {
        ColorOverride::Argb(0xFF00FF00)
    } else if content.eq_ignore_ascii_case("blue") {
        ColorOverride::Argb(0xFF0000FF)
    } else if content.eq_ignore_ascii_case("cyan") {
        ColorOverride::Argb(0xFF00FFFF)
    } else if content.eq_ignore_ascii_case("magenta") {
        ColorOverride::Argb(0xFFFF00FF)
    } else if content.eq_ignore_ascii_case("yellow") {
        ColorOverride::Argb(0xFFFFFF00)
    } else {
        let rest = content
            .get(.."color".len())
            .is_some_and(|p| p.eq_ignore_ascii_case("color"))
            .then(|| &content["color".len()..])?;
        let idx: u8 = rest.trim().parse().ok()?;
        ColorOverride::Indexed(idx)
    })
}

fn parse_locale_override(content: &str) -> Option<Locale> {
    // Locale/currency tags are encoded as `[$$-409]` where the locale is the hex LCID suffix.
    let after = content.strip_prefix('$')?;
    let (_, locale) = after.split_once('-')?;
    let lcid = u32::from_str_radix(locale.trim(), 16).ok()?;
    locale_for_lcid(lcid)
}

/// Best-effort mapping from a Windows/Excel LCID (locale identifier) to a
/// [`Locale`] definition.
///
/// Excel format codes sometimes embed locale tags in bracket tokens such as
/// `[$€-407]` (where `0x0407` is `de-DE`). This helper lets importers map those
/// LCIDs to a `Locale` so that thousands/decimal/date separators render
/// correctly.
///
/// Note: [`Locale`] only models separators. Many LCIDs differ in other ways
/// (currency, calendar, month names, date ordering, etc).
pub fn locale_for_lcid(lcid: u32) -> Option<Locale> {
    // Best-effort LCID -> Locale mapping for the most common locales seen in
    // Excel format codes (e.g. currency tokens like `[$€-407]`).
    //
    // Note: `Locale` only models separators. Many LCIDs differ in other ways
    // (currency, calendar, month names, etc). Those aspects are intentionally
    // out of scope for this mapping.
    Some(match lcid {
        // --- English (decimal '.', thousands ',', date '/', time ':') ---
        0x0409 | // en-US
        0x0809 | // en-GB
        0x0C09 | // en-AU
        0x1009 | // en-CA
        0x1409 | // en-NZ
        0x1809 | // en-IE
        0x1C09   // en-ZA
            => Locale::en_us(),

        // --- German (decimal ',', thousands '.', date '.', time ':') ---
        0x0407 | // de-DE
        0x0C07 | // de-AT
        0x1007 | // de-LU
        0x1407   // de-LI
            => Locale::de_de(),

        // --- Swiss German (decimal '.', thousands ''', date '.', time ':') ---
        0x0807 => Locale {
            decimal_sep: '.',
            thousands_sep: '\'',
            date_sep: '.',
            time_sep: ':',
        }, // de-CH

        // --- French (decimal ',', thousands NBSP, date '/', time ':') ---
        0x040C | // fr-FR
        0x080C | // fr-BE
        0x0C0C | // fr-CA
        0x140C | // fr-LU
        0x180C   // fr-MC
            => Locale::fr_fr(),

        // --- Swiss French (decimal '.', thousands ''', date '.', time ':') ---
        0x100C => Locale {
            decimal_sep: '.',
            thousands_sep: '\'',
            date_sep: '.',
            time_sep: ':',
        }, // fr-CH

        // --- Italian (decimal ',', thousands '.', date '/', time ':') ---
        0x0410 => Locale::it_it(), // it-IT
        // Swiss Italian (decimal '.', thousands ''', date '.', time ':')
        0x0810 => Locale {
            decimal_sep: '.',
            thousands_sep: '\'',
            date_sep: '.',
            time_sep: ':',
        }, // it-CH

        // --- Spanish (decimal ',', thousands '.', date '/', time ':') ---
        0x040A | // es-ES (traditional sort)
        0x0C0A | // es-ES (modern sort)
        0x2C0A | // es-AR
        0x240A   // es-CO
            => Locale::es_es(),

        // --- Spanish (Mexico): decimal '.', thousands ',', date '/', time ':' ---
        0x080A => Locale::en_us(), // es-MX

        // --- Portuguese (approximate with Spanish separators) ---
        0x0416 | // pt-BR
        0x0816   // pt-PT
            => Locale::es_es(),

        // --- Dutch (decimal ',', thousands '.', date '-', time ':') ---
        0x0413 | // nl-NL
        0x0813   // nl-BE
            => Locale {
                decimal_sep: ',',
                thousands_sep: '.',
                date_sep: '-',
                time_sep: ':',
            },

        // --- Swedish (decimal ',', thousands NBSP, date '-', time ':') ---
        0x041D | // sv-SE
        0x081D   // sv-FI
            => Locale {
                decimal_sep: ',',
                thousands_sep: '\u{00A0}',
                date_sep: '-',
                time_sep: ':',
            },

        // --- Danish (decimal ',', thousands '.', date '-', time ':') ---
        0x0406 => Locale {
            decimal_sep: ',',
            thousands_sep: '.',
            date_sep: '-',
            time_sep: ':',
        }, // da-DK

        // --- Norwegian (decimal ',', thousands NBSP, date '.', time ':') ---
        0x0414 | // nb-NO
        0x0814   // nn-NO
            => Locale {
                decimal_sep: ',',
                thousands_sep: '\u{00A0}',
                date_sep: '.',
                time_sep: ':',
            },

        // --- Finnish (decimal ',', thousands NBSP, date '.', time ':') ---
        0x040B => Locale {
            decimal_sep: ',',
            thousands_sep: '\u{00A0}',
            date_sep: '.',
            time_sep: ':',
        }, // fi-FI

        // --- Polish (decimal ',', thousands NBSP, date '.', time ':') ---
        0x0415 => Locale {
            decimal_sep: ',',
            thousands_sep: '\u{00A0}',
            date_sep: '.',
            time_sep: ':',
        }, // pl-PL

        // --- Czech / Slovak (decimal ',', thousands NBSP, date '.', time ':') ---
        0x0405 | // cs-CZ
        0x041B   // sk-SK
            => Locale {
                decimal_sep: ',',
                thousands_sep: '\u{00A0}',
                date_sep: '.',
                time_sep: ':',
            },

        // --- Hungarian (decimal ',', thousands NBSP, date '.', time ':') ---
        0x040E => Locale {
            decimal_sep: ',',
            thousands_sep: '\u{00A0}',
            date_sep: '.',
            time_sep: ':',
        }, // hu-HU

        // --- Romanian (decimal ',', thousands '.', date '.', time ':') ---
        0x0418 => Locale {
            decimal_sep: ',',
            thousands_sep: '.',
            date_sep: '.',
            time_sep: ':',
        }, // ro-RO

        // --- Russian / Ukrainian (decimal ',', thousands NBSP, date '.', time ':') ---
        0x0419 | // ru-RU
        0x0422   // uk-UA
            => Locale {
                decimal_sep: ',',
                thousands_sep: '\u{00A0}',
                date_sep: '.',
                time_sep: ':',
            },

        // --- Turkish (decimal ',', thousands '.', date '.', time ':') ---
        0x041F => Locale {
            decimal_sep: ',',
            thousands_sep: '.',
            date_sep: '.',
            time_sep: ':',
        }, // tr-TR

        // --- Greek (decimal ',', thousands '.', date '/', time ':') ---
        0x0408 => Locale {
            decimal_sep: ',',
            thousands_sep: '.',
            date_sep: '/',
            time_sep: ':',
        }, // el-GR

        // --- Hebrew (decimal '.', thousands ',', date '/', time ':') ---
        0x040D => Locale {
            decimal_sep: '.',
            thousands_sep: ',',
            date_sep: '/',
            time_sep: ':',
        }, // he-IL

        // --- Arabic / Hindi (approximate with en-US separators) ---
        0x0401 | // ar-SA
        0x0439   // hi-IN
            => Locale::en_us(),

        // --- Korean (decimal '.', thousands ',', date '-', time ':') ---
        0x0412 => Locale {
            decimal_sep: '.',
            thousands_sep: ',',
            date_sep: '-',
            time_sep: ':',
        }, // ko-KR

        // --- Thai (decimal '.', thousands ',', date '/', time ':') ---
        0x041E => Locale::en_us(), // th-TH (approx)

        // --- Indonesian / Vietnamese (decimal ',', thousands '.', date '/', time ':') ---
        0x0421 | // id-ID
        0x042A   // vi-VN
            => Locale::es_es(),

        // --- Bulgarian / Croatian / Slovenian (decimal ',', thousands NBSP/dot, date '.', time ':') ---
        0x0402 => Locale {
            decimal_sep: ',',
            thousands_sep: '\u{00A0}',
            date_sep: '.',
            time_sep: ':',
        }, // bg-BG
        0x041A | // hr-HR
        0x0424   // sl-SI
            => Locale {
                decimal_sep: ',',
                thousands_sep: '.',
                date_sep: '.',
                time_sep: ':',
            },

        // --- Baltic (decimal ',', thousands NBSP, date '.', time ':') ---
        0x0425 | // et-EE
        0x0426 | // lv-LV
        0x0427   // lt-LT
            => Locale {
                decimal_sep: ',',
                thousands_sep: '\u{00A0}',
                date_sep: '.',
                time_sep: ':',
            },

        // --- Japanese (map to en-US separators) ---
        0x0411 => Locale::en_us(), // ja-JP

        // --- Chinese (map to en-US separators) ---
        0x0404 | // zh-TW
        0x0804 | // zh-CN
        0x0C04 | // zh-HK
        0x1004 | // zh-SG
        0x1404   // zh-MO
            => Locale::en_us(),

        _ => return None,
    })
}

#[cfg(test)]
mod tests {
    use super::locale_for_lcid;
    use crate::Locale;

    #[test]
    fn lcid_maps_to_locale_separators() {
        // English family.
        assert_eq!(locale_for_lcid(0x0409), Some(Locale::en_us())); // en-US
        assert_eq!(locale_for_lcid(0x0809), Some(Locale::en_us())); // en-GB

        // German family.
        assert_eq!(locale_for_lcid(0x0407), Some(Locale::de_de())); // de-DE
        assert_eq!(locale_for_lcid(0x0C07), Some(Locale::de_de())); // de-AT

        // French family.
        assert_eq!(locale_for_lcid(0x040C), Some(Locale::fr_fr())); // fr-FR
        assert_eq!(locale_for_lcid(0x0C0C), Some(Locale::fr_fr())); // fr-CA

        // Romance languages.
        assert_eq!(locale_for_lcid(0x0410), Some(Locale::it_it())); // it-IT
        assert_eq!(locale_for_lcid(0x0C0A), Some(Locale::es_es())); // es-ES modern
        assert_eq!(locale_for_lcid(0x0416), Some(Locale::es_es())); // pt-BR (approx)

        // Dutch / Swedish (custom date separators).
        assert_eq!(
            locale_for_lcid(0x0413),
            Some(Locale {
                decimal_sep: ',',
                thousands_sep: '.',
                date_sep: '-',
                time_sep: ':',
            })
        );
        assert_eq!(
            locale_for_lcid(0x041D),
            Some(Locale {
                decimal_sep: ',',
                thousands_sep: '\u{00A0}',
                date_sep: '-',
                time_sep: ':',
            })
        );

        // Asian locales mapped to en-US separators.
        assert_eq!(locale_for_lcid(0x0411), Some(Locale::en_us())); // ja-JP
        assert_eq!(locale_for_lcid(0x0804), Some(Locale::en_us())); // zh-CN

        // Swiss variants.
        assert_eq!(
            locale_for_lcid(0x0807),
            Some(Locale {
                decimal_sep: '.',
                thousands_sep: '\'',
                date_sep: '.',
                time_sep: ':',
            })
        ); // de-CH

        // Additional common European locales.
        assert_eq!(
            locale_for_lcid(0x0419),
            Some(Locale {
                decimal_sep: ',',
                thousands_sep: '\u{00A0}',
                date_sep: '.',
                time_sep: ':',
            })
        ); // ru-RU
        assert_eq!(
            locale_for_lcid(0x0415),
            Some(Locale {
                decimal_sep: ',',
                thousands_sep: '\u{00A0}',
                date_sep: '.',
                time_sep: ':',
            })
        ); // pl-PL
        assert_eq!(
            locale_for_lcid(0x0406),
            Some(Locale {
                decimal_sep: ',',
                thousands_sep: '.',
                date_sep: '-',
                time_sep: ':',
            })
        ); // da-DK
        assert_eq!(
            locale_for_lcid(0x041F),
            Some(Locale {
                decimal_sep: ',',
                thousands_sep: '.',
                date_sep: '.',
                time_sep: ':',
            })
        ); // tr-TR

        // es-MX uses en-US separators (best-effort).
        assert_eq!(locale_for_lcid(0x080A), Some(Locale::en_us())); // es-MX

        // Korean uses '-' date separator.
        assert_eq!(
            locale_for_lcid(0x0412),
            Some(Locale {
                decimal_sep: '.',
                thousands_sep: ',',
                date_sep: '-',
                time_sep: ':',
            })
        ); // ko-KR

        assert_eq!(locale_for_lcid(0x9999), None);
    }
}
