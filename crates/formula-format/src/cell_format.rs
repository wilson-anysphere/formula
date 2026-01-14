use crate::{builtin_format_code, FormatCode, BUILTIN_NUM_FMT_ID_PLACEHOLDER_PREFIX};

/// Return Excel-compatible `CELL("format")` classification code for an Excel number format string.
///
/// This is a **classification** helper: it does not attempt to fully parse/render the format code.
///
/// The return value follows Excel's `CELL("format")` conventions for common numeric formats:
/// - `"G"` for General
/// - `"F<n>"` for fixed formats (`n` = decimal places)
/// - `"N<n>"` for number formats that use the thousands separator (`n` = decimal places)
/// - `"C<n>"` for currency formats (`n` = decimal places)
/// - `"P<n>"` for percent formats (`n` = decimal places)
/// - `"S<n>"` for scientific formats (`n` = decimal places)
/// - `"D<n>"` for date/time formats (Excel uses `D1`..`D9`)
/// - `"@"` for text formats
/// - `"N"` when the format does not match any of the recognized families (e.g. fractions)
///
/// Currency detection accounts for:
/// - common currency symbols (`$`, `€`, `£`, `¥`) outside quotes/escapes
/// - OOXML bracket currency tokens like `[$€-407]` (but *not* locale-only tokens like `[$-409]`).
pub fn cell_format_code(format_code: Option<&str>) -> String {
    let code = format_code.unwrap_or("General");
    let code = if code.trim().is_empty() { "General" } else { code };

    // Excel's `CELL("format")` has specific `D*` classification codes for the built-in
    // date/time number formats. Some importers preserve these built-ins as placeholder strings
    // like `__builtin_numFmtId:14` instead of embedding the concrete format code.
    //
    // We map the common built-in ids directly to Excel's `D*` codes (per the Microsoft Support
    // `CELL` docs) so callers get Excel-compatible results even when the format code isn't known.
    //
    // Reference: Microsoft Support → "CELL format codes"
    // https://support.microsoft.com/en-us/office/cell-function-51bd39a5-f338-4dbe-a33f-955d67c2b2cf
    if let Some(id) = parse_builtin_placeholder_id(code) {
        if let Some(mapped) = builtin_datetime_cell_format_code(id) {
            return mapped.to_string();
        }
    }

    let code = resolve_builtin_placeholder(code).unwrap_or(code);

    // Parse into sections so we can correctly choose the "positive" section when conditions are
    // present. When parsing fails, fall back to General classification.
    let parsed = FormatCode::parse(code).unwrap_or_else(|_| FormatCode::general());
    let positive = parsed.select_section_for_number(1.0);
    let pattern = positive.pattern;

    if pattern.trim().eq_ignore_ascii_case("general") {
        return "G".to_string();
    }

    if let Some(code) = classify_datetime_pattern_as_cell_format_code(pattern) {
        return code;
    }

    if crate::number::pattern_is_text(pattern) {
        return "@".to_string();
    }

    // For number formats, Excel returns `N` for patterns that don't fit the standard
    // fixed/currency/percent/scientific families. Examples include:
    // - fractions (`# ?/?`, `# ??/??`)
    // - patterns with no numeric placeholders (literal-only formats like `"foo"`)
    if !pattern_has_number_placeholders(pattern) || is_fraction_format(pattern) {
        return "N".to_string();
    }

    let decimals = count_decimal_places(pattern).min(9);

    let kind = if is_currency_format(pattern) {
        'C'
    } else if is_percent_format(pattern) {
        'P'
    } else if is_scientific_format(pattern) {
        'S'
    } else if pattern_has_thousands_separator(pattern) {
        'N'
    } else {
        'F'
    };

    format!("{kind}{decimals}")
}

/// Return Excel-compatible `CELL("parentheses")` flag for an Excel number format string.
///
/// Excel returns `1` when negative numbers are displayed using parentheses, and `0` otherwise.
/// With a single-section format (where Excel auto-prefixes `-` for negatives), this reports `0`
/// even if the pattern contains parentheses literals.
///
/// This helper selects the section that Excel would use for a negative numeric value and then
/// scans that section for parenthesis characters, ignoring:
/// - quoted literals (`"..."`)
/// - escaped characters (`\X`)
/// - bracket tokens (`[...]`)
/// - underscore (`_X`) and fill (`*X`) layout tokens whose operands are not rendered literally.
pub fn cell_parentheses_flag(format_code: Option<&str>) -> u8 {
    let code = format_code.unwrap_or("General");
    let code = if code.trim().is_empty() { "General" } else { code };
    let code = resolve_builtin_placeholder(code).unwrap_or(code);

    let parsed = FormatCode::parse(code).unwrap_or_else(|_| FormatCode::general());
    let negative = parsed.select_section_for_number(-1.0);

    // Excel reports 0 for one-section formats (where negatives reuse the first section and Excel
    // automatically prefixes a '-' sign).
    if negative.auto_negative_sign {
        return 0;
    }

    u8::from(pattern_contains_balanced_parentheses(negative.pattern))
}

fn classify_datetime_pattern_as_cell_format_code(pattern: &str) -> Option<String> {
    let tokens = tokenize_datetime_pattern(pattern);
    if tokens.is_empty() {
        return None;
    }

    let tokens = disambiguate_month_minute(tokens);

    let mut has_year = false;
    let mut has_month = false;
    let mut has_day_of_month = false;
    let mut has_weekday = false;

    let mut has_hour = false;
    let mut has_minute = false;
    let mut has_second = false;
    let mut has_fractional_seconds = false;
    let mut has_elapsed_hours = false;
    let mut has_ampm = false;

    let mut has_month_name = false;

    for token in &tokens {
        match token {
            DateTimeToken::Year(_) => has_year = true,
            DateTimeToken::Day(count) => {
                if *count >= 3 {
                    has_weekday = true;
                } else {
                    has_day_of_month = true;
                }
            }
            DateTimeToken::Month(count) => {
                has_month = true;
                if *count >= 3 {
                    has_month_name = true;
                }
            }
            DateTimeToken::Hour(_) => has_hour = true,
            DateTimeToken::Minute(_) => has_minute = true,
            DateTimeToken::Second(_) => has_second = true,
            DateTimeToken::FractionalSeconds(_) => has_fractional_seconds = true,
            DateTimeToken::ElapsedHours(_) => {
                has_elapsed_hours = true;
                has_hour = true;
            }
            DateTimeToken::ElapsedMinutes(_) => has_minute = true,
            DateTimeToken::ElapsedSeconds(_) => has_second = true,
            DateTimeToken::AmPmLong | DateTimeToken::AmPmShort => has_ampm = true,
            DateTimeToken::DateSep | DateTimeToken::TimeSep | DateTimeToken::Literal(_) => {}
            DateTimeToken::MonthOrMinute(_) => unreachable!("month/minute disambiguation should run first"),
        }
    }

    let has_date = has_year || has_month || has_day_of_month || has_weekday || has_month_name;
    let has_time = has_hour || has_minute || has_second || has_ampm || has_elapsed_hours || has_fractional_seconds;

    // If the format includes any date component, Excel reports a D* code even when time components
    // are present (e.g. `m/d/yyyy h:mm`).
    if has_date {
        let d_code = classify_date_tokens_to_cell_code(
            has_year,
            has_month,
            has_day_of_month,
            has_month_name,
            has_weekday,
        );
        return Some(d_code);
    }

    if has_time {
        let t_code = classify_time_tokens_to_cell_code(has_hour, has_second, has_fractional_seconds, has_ampm, has_elapsed_hours);
        return Some(t_code);
    }

    None
}

fn classify_date_tokens_to_cell_code(
    has_year: bool,
    has_month: bool,
    has_day_of_month: bool,
    has_month_name: bool,
    has_weekday: bool,
) -> String {
    // Best-effort mapping to Excel's `CELL("format")` `D1..D5` *date* codes, based on the
    // Microsoft Support table.
    //
    // - D1: `d-mmm-yy` / `dd-mmm-yy`
    // - D2: `d-mmm` / `dd-mmm`
    // - D3: `mmm-yy`
    // - D4: `m/d/yy` (and also `m/d/yy h:mm`)
    // - D5: `mm/dd` (month/day, no year)
    //
    // Excel also uses these codes for locale variants of the above; we classify based on the
    // *shape* of the pattern rather than exact separators.
    if has_month_name {
        if has_day_of_month && has_year {
            return "D1".to_string();
        }
        if has_day_of_month {
            return "D2".to_string();
        }
        if has_year {
            return "D3".to_string();
        }
        // Month/weekday name without day/year is uncommon; treat it as the closest day/month-name
        // family.
        return "D2".to_string();
    }

    // Weekday-only formats like `ddd` / `dddd` don't match the Microsoft table exactly; treat them
    // as a day/month-name category.
    if has_weekday && !has_month && !has_day_of_month && !has_year {
        return "D2".to_string();
    }

    // Numeric month/day (no year).
    if has_month && has_day_of_month && !has_year {
        return "D5".to_string();
    }

    // Default for full numeric dates (including year-first ISO-like dates).
    "D4".to_string()
}

fn classify_time_tokens_to_cell_code(
    has_hour: bool,
    has_second: bool,
    has_fractional_seconds: bool,
    has_ampm: bool,
    has_elapsed_hours: bool,
) -> String {
    // Excel uses `D6..D9` for time-of-day formats.
    //
    // Microsoft Support table:
    // - D6: `h:mm:ss AM/PM`
    // - D7: `h:mm AM/PM`
    // - D8: `h:mm:ss`
    // - D9: `h:mm`
    //
    // For elapsed/duration formats like `[h]:mm:ss` and `mm:ss.0`, Excel does not document
    // dedicated codes; we map them to the closest match (`D8`, time with seconds).
    let has_seconds = has_second || has_fractional_seconds;

    if has_ampm {
        return if has_seconds { "D6".to_string() } else { "D7".to_string() };
    }

    if has_elapsed_hours {
        return "D8".to_string();
    }

    if has_seconds {
        "D8".to_string()
    } else {
        // Treat `mm`-only time formats as `h:mm` for `CELL("format")` purposes.
        // (Excel's table only distinguishes the presence of seconds and AM/PM.)
        let _ = has_hour;
        "D9".to_string()
    }
}

fn parse_builtin_placeholder_id(code: &str) -> Option<u16> {
    code.strip_prefix(BUILTIN_NUM_FMT_ID_PLACEHOLDER_PREFIX)?
        .trim()
        .parse::<u16>()
        .ok()
}

/// Mapping of built-in number format ids to the `CELL("format")` `D*` codes that Excel reports.
///
/// Derived by combining:
/// - Microsoft Support `CELL` documentation ("CELL format codes")
/// - ECMA-376 Part 1 §18.8.30 `numFmt` (built-in `numFmtId` assignments)
fn builtin_datetime_cell_format_code(id: u16) -> Option<&'static str> {
    match id {
        // Standard date/time built-ins (OOXML 0–49).
        14 => Some("D4"), // m/d/yy
        15 => Some("D1"), // d-mmm-yy
        16 => Some("D2"), // d-mmm
        17 => Some("D3"), // mmm-yy
        18 => Some("D7"), // h:mm AM/PM
        19 => Some("D6"), // h:mm:ss AM/PM
        20 => Some("D9"), // h:mm
        21 => Some("D8"), // h:mm:ss
        22 => Some("D4"), // m/d/yy h:mm

        // Locale-reserved ids in the OOXML built-in table.
        27..=31 => Some("D4"),
        32..=36 => Some("D8"),

        // Duration/elapsed-time built-ins.
        45..=47 => Some("D8"),

        // Excel-reserved locale date/time ids (not part of the OOXML 0–49 table).
        50..=58 => Some("D4"),

        _ => None,
    }
}

fn resolve_builtin_placeholder(code: &str) -> Option<&'static str> {
    let rest = code.strip_prefix(BUILTIN_NUM_FMT_ID_PLACEHOLDER_PREFIX)?;
    let id = match rest.trim().parse::<u16>() {
        Ok(id) => id,
        // A malformed placeholder should not be interpreted as a literal format code because
        // the placeholder text contains letters like `m`/`d` which would look like a date/time
        // pattern. Treat it as General instead.
        Err(_) => return Some("General"),
    };

    // Standard OOXML/BIFF built-ins: 0-49.
    if let Some(resolved) = builtin_format_code(id) {
        return Some(resolved);
    }

    // Excel reserves additional (non-OOXML) built-in ids for locale-specific date/time formats.
    // When importers preserve them as placeholders, avoid passing the placeholder string into the
    // parser (which can misclassify it as date/time based on the placeholder text itself).
    //
    // Fall back to a representative date format so callers still classify it as date/time.
    if builtin_id_is_common_datetime(id) {
        return Some(builtin_format_code(14).unwrap_or("General"));
    }

    // Unknown placeholder: behave like General.
    Some("General")
}

fn builtin_id_is_common_datetime(id: u16) -> bool {
    matches!(id, 14..=22 | 27..=36 | 45..=47 | 50..=58)
}
fn pattern_contains_balanced_parentheses(pattern: &str) -> bool {
    let mut in_quotes = false;
    let mut escape = false;
    let mut in_brackets = false;
    let mut chars = pattern.chars();
    let mut has_open = false;
    let mut has_close = false;

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
            // `_X` / `*X` layout tokens consume the following character. The operand is not a
            // literal character for CELL("parentheses") purposes (even if it is '(' or ')').
            '_' | '*' => {
                let _ = chars.next();
            }
            '(' => has_open = true,
            ')' => has_close = true,
            _ => {}
        }
    }

    has_open && has_close
}

fn count_decimal_places(pattern: &str) -> usize {
    let mut in_quotes = false;
    let mut escape = false;
    let mut in_brackets = false;
    let mut after_decimal = false;
    let mut count = 0usize;

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
            '.' => {
                after_decimal = true;
                count = 0;
            }
            '0' | '#' | '?' if after_decimal => count += 1,
            _ if after_decimal => break,
            _ => {}
        }
    }

    if after_decimal { count } else { 0 }
}

fn pattern_has_thousands_separator(pattern: &str) -> bool {
    // Excel uses `,` to indicate grouping in number formats. This helper is intentionally simple:
    // treat any comma outside quoted literals, escapes, and bracket tokens as a thousands
    // separator. (Scaling commas are also counted for now.)
    //
    // Note: `_X` and `*X` are layout tokens whose operands are not rendered literally. A comma used
    // only as a layout operand (e.g. `0_,0` or `0*,0`) should not be treated as grouping.
    let mut in_quotes = false;
    let mut escape = false;
    let mut in_brackets = false;
    let mut chars = pattern.chars();

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
            // Layout tokens consume their operand as a literal character. Skip it so commas used as
            // layout operands do not trigger thousands-separator classification.
            '_' | '*' => {
                let _ = chars.next();
            }
            ',' => return true,
            _ => {}
        }
    }

    false
}

fn is_percent_format(pattern: &str) -> bool {
    scan_outside_quotes(pattern, |ch| ch == '%')
}

fn is_scientific_format(pattern: &str) -> bool {
    scan_outside_quotes(pattern, |ch| ch == 'E' || ch == 'e')
}

fn is_currency_format(pattern: &str) -> bool {
    // Detect explicit currency symbols outside quotes/escapes, OR bracket currency tokens like
    // `[$€-407]`. Locale-only tokens like `[$-409]` should *not* be treated as currency.
    scan_outside_quotes(pattern, |ch| matches!(ch, '$' | '€' | '£' | '¥'))
        || contains_bracket_currency_token(pattern)
}

fn contains_bracket_currency_token(pattern: &str) -> bool {
    let mut in_quotes = false;
    let mut escape = false;
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
                let mut content = String::new();
                let mut closed = false;
                while let Some(c) = chars.next() {
                    if c == ']' {
                        closed = true;
                        break;
                    }
                    content.push(c);
                }
                if !closed {
                    // No closing bracket: treat as literal and stop probing this token.
                    continue;
                }
                if bracket_is_currency(&content) {
                    return true;
                }
            }
            _ => {}
        }
    }

    false
}

fn bracket_is_currency(content: &str) -> bool {
    let content = content.trim();
    let Some(after) = content.strip_prefix('$') else {
        return false;
    };
    // Bracket currency/locale tokens are encoded as `[$<currency>-<lcid>]`.
    //
    // Real-world OOXML often embeds 3-letter currency codes (e.g. `USD`) or multi-character
    // symbols (e.g. `R$`, `kr`), so `<currency>` may contain multiple characters. We also want to
    // avoid treating locale-only overrides like `[$-409]` as currency.
    //
    // Parse the LCID suffix from the *last* `-` so we don't assume the currency portion is a
    // single character.
    let Some((currency, _lcid)) = after.rsplit_once('-') else {
        return false;
    };
    !currency.is_empty() && currency.chars().any(|c| c != '-')
}

fn scan_outside_quotes(pattern: &str, pred: impl Fn(char) -> bool) -> bool {
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
            _ if pred(ch) => return true,
            _ => {}
        }
    }

    false
}

fn pattern_has_number_placeholders(pattern: &str) -> bool {
    scan_outside_quotes(pattern, |ch| matches!(ch, '0' | '#' | '?'))
}

fn is_fraction_format(pattern: &str) -> bool {
    let mut in_quotes = false;
    let mut escape = false;
    let mut in_brackets = false;
    let mut slash_idx: Option<usize> = None;

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
            '[' => in_brackets = true,
            '/' => {
                slash_idx = Some(idx);
                break;
            }
            _ => {}
        }
    }

    let Some(slash_idx) = slash_idx else {
        return false;
    };

    let left = &pattern[..slash_idx];
    let right = &pattern[slash_idx + 1..];

    // Fractions must have at least one numeric placeholder before the '/'.
    if !scan_outside_quotes(left, |ch| matches!(ch, '0' | '#' | '?')) {
        return false;
    }

    // Denominators can either be placeholder-driven (`??`) or fixed (`16`).
    if scan_outside_quotes(right, |ch| matches!(ch, '0' | '#' | '?')) {
        return true;
    }

    // Fixed denominator: parse a leading integer after trimming whitespace.
    let trimmed = right.trim_start();
    let mut saw_digit = false;
    for ch in trimmed.chars() {
        if ch.is_ascii_digit() {
            saw_digit = true;
            continue;
        }
        break;
    }

    saw_digit
}

#[derive(Debug, Clone, PartialEq, Eq)]
enum DateTimeToken {
    Literal(String),
    Year(usize),
    Day(usize),
    MonthOrMinute(usize),
    Month(usize),
    Hour(usize),
    Minute(usize),
    Second(usize),
    FractionalSeconds(usize),
    DateSep,
    TimeSep,
    AmPmLong,
    AmPmShort,
    ElapsedHours(usize),
    ElapsedMinutes(usize),
    ElapsedSeconds(usize),
}

fn tokenize_datetime_pattern(pattern: &str) -> Vec<DateTimeToken> {
    // This tokenization is intentionally best-effort and mirrors the rules used by
    // `crate::datetime` for rendering:
    // - ignore content inside quotes
    // - respect `\\` escapes
    // - ignore bracket tokens except elapsed time `[h]`, `[m]`, `[s]`
    // - detect `AM/PM` and `A/P` markers
    let mut tokens = Vec::new();
    let mut in_quotes = false;
    let mut literal_buf = String::new();
    let mut chars = pattern.chars().peekable();

    while let Some(ch) = chars.next() {
        if in_quotes {
            if ch == '"' {
                in_quotes = false;
            } else {
                literal_buf.push(ch);
            }
            continue;
        }

        match ch {
            '"' => in_quotes = true,
            '\\' => {
                if let Some(next) = chars.next() {
                    literal_buf.push(next);
                }
            }
            '/' => {
                flush_literal(&mut literal_buf, &mut tokens);
                tokens.push(DateTimeToken::DateSep);
            }
            ':' => {
                flush_literal(&mut literal_buf, &mut tokens);
                tokens.push(DateTimeToken::TimeSep);
            }
            '[' => {
                let mut content = String::new();
                while let Some(c) = chars.next() {
                    if c == ']' {
                        break;
                    }
                    content.push(c);
                }
                let lower = content.to_ascii_lowercase();
                if let Some(tok) = parse_elapsed_time_token(&lower) {
                    flush_literal(&mut literal_buf, &mut tokens);
                    tokens.push(tok);
                }
                // Ignore all other bracket tokens.
            }
            'a' | 'A' => {
                // AM/PM or A/P markers (case-insensitive).
                let mut probe = String::new();
                probe.push(ch);
                let mut clone = chars.clone();
                for _ in 0..4 {
                    if let Some(c) = clone.next() {
                        probe.push(c);
                    } else {
                        break;
                    }
                }

                let lower = probe.to_ascii_lowercase();
                if lower.starts_with("am/pm") {
                    // Consume `M/PM` (4 chars).
                    for _ in 0..4 {
                        chars.next();
                    }
                    flush_literal(&mut literal_buf, &mut tokens);
                    tokens.push(DateTimeToken::AmPmLong);
                } else if lower.starts_with("a/p") {
                    // Consume `/P` (2 chars).
                    for _ in 0..2 {
                        chars.next();
                    }
                    flush_literal(&mut literal_buf, &mut tokens);
                    tokens.push(DateTimeToken::AmPmShort);
                } else {
                    literal_buf.push(ch);
                }
            }
            'y' | 'Y' => {
                let count = consume_run(ch, &mut chars);
                flush_literal(&mut literal_buf, &mut tokens);
                tokens.push(DateTimeToken::Year(count));
            }
            'd' | 'D' => {
                let count = consume_run(ch, &mut chars);
                flush_literal(&mut literal_buf, &mut tokens);
                tokens.push(DateTimeToken::Day(count));
            }
            'h' | 'H' => {
                let count = consume_run(ch, &mut chars);
                flush_literal(&mut literal_buf, &mut tokens);
                tokens.push(DateTimeToken::Hour(count));
            }
            's' | 'S' => {
                let count = consume_run(ch, &mut chars);
                flush_literal(&mut literal_buf, &mut tokens);
                tokens.push(DateTimeToken::Second(count));

                // Fractional seconds: `ss.0`, `ss.00`, `ss.000`, etc.
                if chars.peek().copied() == Some('.') {
                    let mut clone = chars.clone();
                    let _ = clone.next(); // '.'
                    let mut zeros = 0usize;
                    while let Some('0') = clone.next() {
                        zeros += 1;
                    }
                    if zeros > 0 {
                        // Consume '.' + zeros from the original iterator.
                        let _ = chars.next();
                        for _ in 0..zeros {
                            let _ = chars.next();
                        }
                        tokens.push(DateTimeToken::FractionalSeconds(zeros));
                    }
                }
            }
            'm' | 'M' => {
                let count = consume_run(ch, &mut chars);
                flush_literal(&mut literal_buf, &mut tokens);
                tokens.push(DateTimeToken::MonthOrMinute(count));
            }
            _ => literal_buf.push(ch),
        }
    }

    flush_literal(&mut literal_buf, &mut tokens);
    tokens
}

fn parse_elapsed_time_token(lower: &str) -> Option<DateTimeToken> {
    if lower.is_empty() {
        return None;
    }
    let count = lower.chars().count();
    if lower.chars().all(|c| c == 'h') {
        return Some(DateTimeToken::ElapsedHours(count));
    }
    if lower.chars().all(|c| c == 'm') {
        return Some(DateTimeToken::ElapsedMinutes(count));
    }
    if lower.chars().all(|c| c == 's') {
        return Some(DateTimeToken::ElapsedSeconds(count));
    }
    None
}

fn consume_run(first: char, chars: &mut std::iter::Peekable<std::str::Chars<'_>>) -> usize {
    let mut count = 1;
    while let Some(next) = chars.peek().copied() {
        if next.eq_ignore_ascii_case(&first) {
            chars.next();
            count += 1;
        } else {
            break;
        }
    }
    count
}

fn flush_literal(buf: &mut String, tokens: &mut Vec<DateTimeToken>) {
    if buf.is_empty() {
        return;
    }
    tokens.push(DateTimeToken::Literal(std::mem::take(buf)));
}

fn disambiguate_month_minute(mut tokens: Vec<DateTimeToken>) -> Vec<DateTimeToken> {
    // Replace MonthOrMinute based on adjacent time tokens.
    for idx in 0..tokens.len() {
        let DateTimeToken::MonthOrMinute(count) = tokens[idx] else {
            continue;
        };

        // In Excel formats, `mmm`/`mmmm`/`mmmmm` are always month name variants. Minutes only use
        // `m` or `mm` (disambiguated by neighboring hour/second tokens).
        if count >= 3 {
            tokens[idx] = DateTimeToken::Month(count);
            continue;
        }

        let prev = prev_non_literal(&tokens, idx);
        let next = next_non_literal(&tokens, idx);

        let is_minute = matches!(prev, Some(DateTimeToken::Hour(_)) | Some(DateTimeToken::ElapsedHours(_)))
            || matches!(next, Some(DateTimeToken::Second(_)) | Some(DateTimeToken::ElapsedSeconds(_)));

        tokens[idx] = if is_minute {
            DateTimeToken::Minute(count)
        } else {
            DateTimeToken::Month(count)
        };
    }

    tokens
}

fn prev_non_literal(tokens: &[DateTimeToken], idx: usize) -> Option<&DateTimeToken> {
    for j in (0..idx).rev() {
        match &tokens[j] {
            DateTimeToken::Literal(_) | DateTimeToken::DateSep | DateTimeToken::TimeSep => continue,
            t => return Some(t),
        }
    }
    None
}

fn next_non_literal(tokens: &[DateTimeToken], idx: usize) -> Option<&DateTimeToken> {
    for j in idx + 1..tokens.len() {
        match &tokens[j] {
            DateTimeToken::Literal(_) | DateTimeToken::DateSep | DateTimeToken::TimeSep => continue,
            t => return Some(t),
        }
    }
    None
}
