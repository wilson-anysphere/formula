use crate::{builtin_format_code, FormatCode, BUILTIN_NUM_FMT_ID_PLACEHOLDER_PREFIX};

/// Return Excel-compatible `CELL("format")` classification code for an Excel number format string.
///
/// This is a **classification** helper: it does not attempt to fully parse/render the format code.
///
/// The return value follows Excel's `CELL("format")` conventions for common numeric formats:
/// - `"G"` for General
/// - `"F<n>"` for fixed/number formats (`n` = decimal places)
/// - `"C<n>"` for currency formats (`n` = decimal places)
/// - `"P<n>"` for percent formats (`n` = decimal places)
/// - `"S<n>"` for scientific formats (`n` = decimal places)
/// - `"D<n>"` for date formats (best-effort; Excel uses `D1`..`D9`)
/// - `"T<n>"` for time formats (best-effort; Excel uses `T1`..`T9`)
/// - `"@"` for text formats
///
/// Currency detection accounts for:
/// - common currency symbols (`$`, `€`, `£`, `¥`) outside quotes/escapes
/// - OOXML bracket currency tokens like `[$€-407]` (but *not* locale-only tokens like `[$-409]`).
pub fn cell_format_code(format_code: Option<&str>) -> String {
    let code = format_code.unwrap_or("General");
    let code = if code.trim().is_empty() { "General" } else { code };
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

    let decimals = count_decimal_places(pattern).min(9);

    let kind = if is_currency_format(pattern) {
        'C'
    } else if is_percent_format(pattern) {
        'P'
    } else if is_scientific_format(pattern) {
        'S'
    } else {
        'F'
    };

    format!("{kind}{decimals}")
}

/// Return Excel-compatible `CELL("parentheses")` flag for an Excel number format string.
///
/// Excel returns `1` when negative numbers are displayed using parentheses, and `0` otherwise.
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

    if pattern_contains_parentheses(negative.pattern) {
        1
    } else {
        0
    }
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
    // Best-effort mapping to Excel's `CELL("format")` date codes `D1..D9`.
    //
    // Observed conventions (Excel/OOXML built-ins):
    // - D1: short numeric date (month/day/year ordering is locale-dependent)
    // - D2: long date (month names, weekday names, etc)
    // - D3: day + month name (no year)
    // - D4: month name + year (no day)
    // - D5: month/day (no year)
    //
    // Note: Excel has additional D codes; we map to the closest of the above.
    if has_year && has_month && has_day_of_month {
        if has_month_name || has_weekday {
            return "D2".to_string();
        }
        return "D1".to_string();
    }

    if has_year && has_month && !has_day_of_month {
        if has_month_name {
            return "D4".to_string();
        }
        return "D1".to_string();
    }

    if !has_year && has_month && has_day_of_month {
        if has_month_name {
            return "D3".to_string();
        }
        return "D5".to_string();
    }

    if has_month_name || has_weekday {
        return "D2".to_string();
    }

    // Fallback for partial/ambiguous date formats.
    "D1".to_string()
}

fn classify_time_tokens_to_cell_code(
    has_hour: bool,
    has_second: bool,
    has_fractional_seconds: bool,
    has_ampm: bool,
    has_elapsed_hours: bool,
) -> String {
    // Best-effort mapping to Excel's `CELL("format")` time codes `T1..T9`.
    //
    // Observed conventions (Excel/OOXML built-ins):
    // - T1: h:mm AM/PM
    // - T2: h:mm:ss AM/PM
    // - T3: h:mm (24-hour)
    // - T4: h:mm:ss (24-hour)
    // - T5: mm:ss
    // - T6: [h]:mm:ss (elapsed time)
    // - T7: mm:ss.0 (fractional seconds)
    if has_ampm {
        return if has_second || has_fractional_seconds {
            "T2".to_string()
        } else {
            "T1".to_string()
        };
    }

    if has_elapsed_hours {
        return "T6".to_string();
    }

    if has_fractional_seconds {
        return "T7".to_string();
    }

    if has_second {
        return if has_hour {
            "T4".to_string()
        } else {
            "T5".to_string()
        };
    }

    // No seconds.
    if has_hour {
        "T3".to_string()
    } else {
        // Time formats with no explicit hours (rare) still fall under `T3` in Excel's classification
        // table; treat them like `h:mm`.
        "T3".to_string()
    }
}

fn resolve_builtin_placeholder(code: &str) -> Option<&'static str> {
    let id = code
        .strip_prefix(BUILTIN_NUM_FMT_ID_PLACEHOLDER_PREFIX)?
        .trim()
        .parse::<u16>()
        .ok()?;
    builtin_format_code(id)
}

fn pattern_contains_parentheses(pattern: &str) -> bool {
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
            // `_X` / `*X` layout tokens consume the following character. The operand is not a
            // literal character for CELL("parentheses") purposes (even if it is '(' or ')').
            '_' | '*' => {
                let _ = chars.next();
            }
            '(' | ')' => return true,
            _ => {}
        }
    }

    false
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
