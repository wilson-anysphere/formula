use crate::{builtin_format_code, FormatCode, BUILTIN_NUM_FMT_ID_PLACEHOLDER_PREFIX};

/// Classification result for Excel's `CELL("format")`, `CELL("color")`, and
/// `CELL("parentheses")` info types.
#[derive(Debug, Clone, PartialEq, Eq)]
pub struct CellFormatClassification {
    /// Return value for `CELL("format")` (e.g. `"G"`, `"F2"`, `"C0"`, `"D1"`).
    pub cell_format_code: String,
    /// `CELL("color")`: whether negative numbers are formatted with a color token.
    pub negative_in_color: bool,
    /// `CELL("parentheses")`: whether negative numbers are formatted with parentheses.
    pub negative_in_parentheses: bool,
}

/// Classify an Excel/OOXML number format code into the semantics needed by
/// `CELL("format")`, `CELL("color")`, and `CELL("parentheses")`.
///
/// The input format code is expected to match the representation stored in
/// `formula-model::Style.number_format`, including built-in placeholders like
/// `__builtin_numFmtId:14`.
///
/// - `None` or empty/whitespace-only strings are treated as `"General"`.
/// - Built-in placeholders (`__builtin_numFmtId:<id>`) are resolved against
///   [`builtin_format_code`] when `id` is within the standard OOXML built-in
///   range 0–49.
/// - For reserved built-in ids outside that range (notably 50–58), classification
///   is best-effort and defaults to date/time.
pub fn classify_cell_format(format_code: Option<&str>) -> CellFormatClassification {
    let mut code = format_code.unwrap_or("General").trim();
    if code.is_empty() {
        code = "General";
    }

    // Built-in placeholder handling.
    if let Some(rest) = code.strip_prefix(BUILTIN_NUM_FMT_ID_PLACEHOLDER_PREFIX) {
        match rest.trim().parse::<u16>() {
            Ok(id) => {
                if let Some(resolved) = builtin_format_code(id) {
                    code = resolved;
                } else if matches!(id, 50..=58) {
                    // Excel reserves many built-in ids beyond 0–49 for locale-specific
                    // date/time formats. We don't have the concrete format code, but
                    // we can at least classify it as date/time and default flags false.
                    return CellFormatClassification {
                        cell_format_code: classify_reserved_datetime_format_id(id).to_string(),
                        negative_in_color: false,
                        negative_in_parentheses: false,
                    };
                } else {
                    // Unknown placeholder id; treat as unrecognized.
                    return CellFormatClassification {
                        cell_format_code: "N".to_string(),
                        negative_in_color: false,
                        negative_in_parentheses: false,
                    };
                }
            }
            Err(_) => {
                // Malformed placeholder; treat as General.
                code = "General";
            }
        }
    }

    let parsed = FormatCode::parse(code).unwrap_or_else(|_| FormatCode::general());

    let positive = parsed.select_section_for_number(1.0);
    let negative = parsed.select_section_for_number(-1.0);

    let cell_format_code = classify_cell_format_section(positive.pattern);

    // Excel reports `CELL("color")=0` / `CELL("parentheses")=0` for one-section formats where
    // the negative sign is applied automatically (i.e. there is no explicit negative section).
    if negative.auto_negative_sign {
        return CellFormatClassification {
            cell_format_code,
            negative_in_color: false,
            negative_in_parentheses: false,
        };
    }

    CellFormatClassification {
        cell_format_code,
        negative_in_color: negative.color.is_some(),
        negative_in_parentheses: section_has_parentheses(negative.pattern),
    }
}

fn classify_cell_format_section(pattern: &str) -> String {
    let pattern = pattern.trim();
    if pattern.is_empty() || pattern.eq_ignore_ascii_case("general") {
        return "G".to_string();
    }

    if crate::number::pattern_is_text(pattern) {
        return "@".to_string();
    }

    if crate::datetime::looks_like_datetime(pattern) {
        return classify_datetime_section(pattern);
    }

    classify_numeric_section(pattern).unwrap_or_else(|| "N".to_string())
}

fn classify_numeric_section(pattern: &str) -> Option<String> {
    let analysis = analyze_numeric_pattern(pattern);
    if !analysis.has_placeholders || analysis.is_fraction {
        return None;
    }

    let family = if analysis.has_percent {
        'P'
    } else if analysis.is_scientific {
        'S'
    } else if analysis.has_currency {
        'C'
    } else {
        'F'
    };

    Some(format!("{family}{}", analysis.decimal_count.min(9)))
}

#[derive(Debug, Clone, Copy, Default)]
struct NumericPatternAnalysis {
    has_placeholders: bool,
    decimal_count: usize,
    has_percent: bool,
    is_scientific: bool,
    has_currency: bool,
    is_fraction: bool,
}

fn analyze_numeric_pattern(pattern: &str) -> NumericPatternAnalysis {
    let mut out = NumericPatternAnalysis::default();

    let mut in_quotes = false;
    let mut escape = false;
    let mut in_brackets = false;
    let mut after_decimal = false;
    let mut in_exponent = false;
    let mut saw_exponent_digits = false;

    for (idx, ch) in pattern.char_indices() {
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
            '[' => {
                in_brackets = true;
                // Currency/locale token is of the form `[$€-407]`.
                if bracket_token_is_currency(pattern, idx) {
                    out.has_currency = true;
                }
            }
            '%' => out.has_percent = true,
            // Heuristic: treat a slash in a numeric pattern as a fraction.
            '/' if out.has_placeholders => out.is_fraction = true,
            '$' | '€' | '£' | '¥' => out.has_currency = true,
            'E' | 'e' if out.has_placeholders => {
                in_exponent = true;
            }
            '.' if out.has_placeholders && !in_exponent => after_decimal = true,
            '0' | '#' | '?' => {
                out.has_placeholders = true;
                if in_exponent {
                    saw_exponent_digits = true;
                } else if after_decimal {
                    out.decimal_count += 1;
                }
            }
            _ => {}
        }
    }

    out.is_scientific = in_exponent && saw_exponent_digits;
    out
}

fn bracket_token_is_currency(pattern: &str, start_idx: usize) -> bool {
    let rest = &pattern[start_idx..];
    let Some(end) = rest.find(']') else {
        return false;
    };
    if end <= 1 {
        return false;
    }

    let content = rest[1..end].trim();
    let Some(after) = content.strip_prefix('$') else {
        return false;
    };

    // Bracket currency/locale tokens are encoded as `[$<currency>-<lcid>]`.
    // Locale-only overrides like `[$-409]` have an empty currency portion.
    let Some((currency, _lcid)) = after.rsplit_once('-') else {
        return false;
    };
    !currency.is_empty()
}

fn classify_datetime_section(pattern: &str) -> String {
    let pattern = strip_leading_non_elapsed_bracket_tokens(pattern);

    // --- Exact matches for Excel built-ins (case-insensitive) ---
    // Date formats.
    if pattern.eq_ignore_ascii_case("m/d/yyyy") || pattern.eq_ignore_ascii_case("m/d/yy") {
        return "D1".to_string();
    }
    if pattern.eq_ignore_ascii_case("d-mmm-yy") {
        return "D2".to_string();
    }
    if pattern.eq_ignore_ascii_case("d-mmm") {
        return "D3".to_string();
    }
    if pattern.eq_ignore_ascii_case("mmm-yy") {
        return "D4".to_string();
    }
    // Datetime built-in.
    if pattern.eq_ignore_ascii_case("m/d/yyyy h:mm") || pattern.eq_ignore_ascii_case("m/d/yy h:mm") {
        return "D5".to_string();
    }

    // Time formats.
    if pattern.eq_ignore_ascii_case("h:mm am/pm") || pattern.eq_ignore_ascii_case("h:mm a/p") {
        return "T1".to_string();
    }
    if pattern.eq_ignore_ascii_case("h:mm:ss am/pm") || pattern.eq_ignore_ascii_case("h:mm:ss a/p") {
        return "T2".to_string();
    }
    if pattern.eq_ignore_ascii_case("h:mm") {
        return "T3".to_string();
    }
    if pattern.eq_ignore_ascii_case("h:mm:ss") {
        return "T4".to_string();
    }
    if pattern.eq_ignore_ascii_case("mm:ss") {
        return "T5".to_string();
    }
    if pattern.eq_ignore_ascii_case("[h]:mm:ss") {
        return "T6".to_string();
    }
    if pattern.eq_ignore_ascii_case("mm:ss.0") {
        return "T7".to_string();
    }

    // --- Heuristic fallback for custom patterns ---
    let analysis = analyze_datetime_pattern(pattern);
    if analysis.has_date() && analysis.has_time() {
        return "D5".to_string();
    }
    if analysis.has_date() {
        return classify_custom_date_pattern(&analysis);
    }
    if analysis.has_time() {
        return classify_custom_time_pattern(&analysis);
    }

    // As a last resort, default to a date-like code.
    "D1".to_string()
}

#[derive(Debug, Clone, Copy, Default)]
struct DateTimePatternAnalysis {
    // Date tokens.
    has_year: bool,
    has_day: bool,
    // `m` tokens can be month or minute depending on context.
    has_m: bool,
    // Month-name tokens (`mmm`, `mmmm`).
    has_month_name: bool,

    // Time tokens.
    has_hour: bool,
    has_second: bool,
    has_ampm: bool,
    has_elapsed: bool,

    // Separators / modifiers.
    has_colon: bool,
    has_fractional_seconds: bool,
}

impl DateTimePatternAnalysis {
    fn has_time_context(self) -> bool {
        self.has_hour || self.has_second || self.has_ampm || self.has_elapsed || self.has_colon
    }

    fn has_date(self) -> bool {
        self.has_year
            || self.has_day
            || self.has_month_name
            || (self.has_m && !self.has_time_context())
    }

    fn has_time(self) -> bool {
        self.has_time_context()
    }
}

fn analyze_datetime_pattern(pattern: &str) -> DateTimePatternAnalysis {
    let mut out = DateTimePatternAnalysis::default();

    let mut in_quotes = false;
    let mut escape = false;

    // Track whether we've seen an `s` token; used for fractional seconds detection.
    let mut saw_second_token = false;

    // Iterate with `Peekable` so we can detect token runs like `mmm`.
    let mut chars = pattern.chars().peekable();

    while let Some(ch) = chars.next() {
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

        match ch {
            '"' => in_quotes = true,
            '\\' => escape = true,
            '[' => {
                // Elapsed time: [h], [m], [s]
                let mut first: Option<char> = None;
                let mut all_same = true;
                let mut saw_any = false;
                while let Some(c) = chars.next() {
                    if c == ']' {
                        break;
                    }
                    saw_any = true;
                    let lower = c.to_ascii_lowercase();
                    match first {
                        None => first = Some(lower),
                        Some(f) if f != lower => all_same = false,
                        _ => {}
                    }
                }

                if saw_any {
                    if let Some(f) = first {
                        if all_same && matches!(f, 'h' | 'm' | 's') {
                            out.has_elapsed = true;
                        }
                    }
                }
            }
            ':' => out.has_colon = true,
            // Years
            'y' | 'Y' => out.has_year = true,
            // Days
            'd' | 'D' => out.has_day = true,
            // Hours
            'h' | 'H' => out.has_hour = true,
            // Seconds
            's' | 'S' => {
                out.has_second = true;
                saw_second_token = true;
            }
            // Month / minute (m)
            'm' | 'M' => {
                out.has_m = true;

                // Count the run length (`m`, `mm`, `mmm`, …).
                let mut run_len = 1usize;
                while matches!(chars.peek(), Some('m' | 'M')) {
                    chars.next();
                    run_len += 1;
                }
                if run_len >= 3 {
                    out.has_month_name = true;
                }
            }
            // AM/PM marker.
            'a' | 'A' => {
                // Check for `AM/PM` or `A/P` markers (case-insensitive) without
                // consuming from the main iterator.
                let mut clone = chars.clone();
                let c1 = clone.next().map(|c| c.to_ascii_lowercase());
                let c2 = clone.next().map(|c| c.to_ascii_lowercase());
                let c3 = clone.next().map(|c| c.to_ascii_lowercase());
                let c4 = clone.next().map(|c| c.to_ascii_lowercase());

                if matches!((c1, c2, c3, c4), (Some('m'), Some('/'), Some('p'), Some('m')))
                    || matches!((c1, c2), (Some('/'), Some('p')))
                {
                    out.has_ampm = true;
                }
            }
            '.' if saw_second_token => {
                // Fractional seconds are encoded as `.0`, `.00`, ... after seconds.
                if matches!(chars.peek(), Some('0' | '#' | '?')) {
                    out.has_fractional_seconds = true;
                }
            }
            _ => {}
        }
    }

    out
}

fn classify_custom_date_pattern(analysis: &DateTimePatternAnalysis) -> String {
    // Best-effort mapping to Excel's D* codes.
    if analysis.has_month_name {
        if analysis.has_day && analysis.has_year {
            return "D2".to_string();
        }
        if analysis.has_day {
            return "D3".to_string();
        }
        if analysis.has_year {
            return "D4".to_string();
        }
    }

    if analysis.has_year {
        return "D1".to_string();
    }
    // Date without an explicit year: fall back to a short date variant.
    "D5".to_string()
}

fn classify_custom_time_pattern(analysis: &DateTimePatternAnalysis) -> String {
    // Best-effort mapping to Excel's T* codes.
    if analysis.has_ampm {
        return if analysis.has_second {
            "T2".to_string()
        } else {
            "T1".to_string()
        };
    }

    if analysis.has_elapsed {
        // Excel has a dedicated code for elapsed `[h]:mm:ss` formats.
        return "T6".to_string();
    }

    // Time without hour tokens is usually `mm:ss` or `mm:ss.0`.
    if !analysis.has_hour && analysis.has_second && analysis.has_colon {
        return if analysis.has_fractional_seconds {
            "T7".to_string()
        } else {
            "T5".to_string()
        };
    }

    if analysis.has_second {
        return "T4".to_string();
    }

    "T3".to_string()
}

fn strip_leading_non_elapsed_bracket_tokens(pattern: &str) -> &str {
    let mut rest = pattern.trim_start();

    loop {
        let Some(stripped) = rest.strip_prefix('[') else {
            return rest.trim();
        };
        let Some(end) = stripped.find(']') else {
            return rest.trim();
        };
        let content = &stripped[..end];

        if is_elapsed_time_token(content) {
            return rest.trim();
        }

        rest = &stripped[end + 1..];
        rest = rest.trim_start();
    }
}

fn is_elapsed_time_token(content: &str) -> bool {
    let mut chars = content.chars();
    let Some(first) = chars.next() else {
        return false;
    };
    let first = first.to_ascii_lowercase();
    if !matches!(first, 'h' | 'm' | 's') {
        return false;
    }
    chars.all(|c| c.to_ascii_lowercase() == first)
}

fn classify_reserved_datetime_format_id(id: u16) -> &'static str {
    // Best-effort mapping for the most common reserved format ids used by Excel.
    // Most callers encounter these via `__builtin_numFmtId:<id>` placeholders.
    match id {
        // Commonly-observed reserved ids are date/time variants. Without the concrete
        // pattern we default to a short date.
        _ => "D1",
    }
}

fn section_has_parentheses(pattern: &str) -> bool {
    let mut in_quotes = false;
    let mut escape = false;
    let mut in_brackets = false;
    let mut skip_next = false;
    let mut saw_open = false;
    let mut saw_close = false;

    for ch in pattern.chars() {
        if skip_next {
            skip_next = false;
            continue;
        }

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
            '\\' => {
                escape = true;
            }
            '_' => {
                // `_X` reserves the width of `X` but does not display it. Ignore the
                // following character for parentheses detection.
                skip_next = true;
            }
            '*' => {
                // `*X` repeats `X` to fill the cell width, but `X` is a layout operand
                // rather than a literal. Ignore the following character for
                // parentheses detection.
                skip_next = true;
            }
            '[' => in_brackets = true,
            '(' => saw_open = true,
            ')' => saw_close = true,
            _ => {}
        }
    }

    saw_open && saw_close
}
