use crate::{literal, FormatOptions};

pub(crate) fn pattern_is_text(pattern: &str) -> bool {
    let (at_count, _) = scan_outside_quotes(pattern, '@');
    if at_count == 0 {
        return false;
    }

    let (first, last) = find_placeholder_span(pattern);
    first.is_none() && last.is_none()
}

pub(crate) fn format_number(
    value: f64,
    pattern: &str,
    auto_negative_sign: bool,
    options: &FormatOptions,
) -> literal::RenderedText {
    if pattern.trim().eq_ignore_ascii_case("general") {
        return literal::RenderedText::new(format_general(value, options));
    }

    if !value.is_finite() {
        return literal::RenderedText::new(value.to_string());
    }

    // Text placeholder format (built-in 49) and related custom patterns:
    // render the number using General, then substitute into the `@` slot(s).
    if pattern_is_text(pattern) {
        let general = format_general(value, options);
        return literal::render_text_section(pattern, &general);
    }

    let (percent_count, _) = scan_outside_quotes(pattern, '%');
    let (first_idx, last_idx) = find_placeholder_span(pattern);

    // If there is no placeholder, treat the pattern as a literal.
    let Some((start, end_placeholder)) = first_idx.zip(last_idx) else {
        return literal::render_literal_segment(pattern);
    };

    // Extend the placeholder span to include trailing scaling commas, e.g. "#,##0,,"
    let mut end = end_placeholder;
    while end < pattern.len() {
        let rest = &pattern[end..];
        if rest.starts_with(',') {
            end += 1;
        } else {
            break;
        }
    }

    let prefix_raw = &pattern[..start];
    let number_raw = &pattern[start..end];
    let suffix_raw = &pattern[end..];

    let prefix = literal::render_literal_segment(prefix_raw);
    let suffix = literal::render_literal_segment(suffix_raw);

    // Apply scaling for percent.
    let mut v = value.abs();
    for _ in 0..percent_count {
        v *= 100.0;
    }

    // Scientific
    if let Some(spec) = parse_scientific(number_raw) {
        let out = format_scientific(v, &spec, options);
        let is_zero = v == 0.0;
        let mut rendered = prefix;
        rendered.push_str(&out);
        rendered.extend(suffix);
        if value < 0.0 && auto_negative_sign && !is_zero {
            rendered.prepend_char('-');
        }
        return rendered;
    }

    if let Some(spec) = parse_fraction(number_raw) {
        let out = format_fraction(v, &spec, options);
        let is_zero = out.trim() == "0" || out.trim().is_empty();
        let mut rendered = prefix;
        rendered.push_str(&out);
        rendered.extend(suffix);
        if value < 0.0 && auto_negative_sign && !is_zero {
            rendered.prepend_char('-');
        }
        return rendered;
    }

    let spec = parse_fixed(number_raw);
    let out = format_fixed(v, &spec, options);
    let is_zero = is_effective_zero_fixed(v, &spec);
    let mut rendered = prefix;
    rendered.push_str(&out);
    rendered.extend(suffix);
    if value < 0.0 && auto_negative_sign && !is_zero {
        rendered.prepend_char('-');
    }
    rendered
}

fn is_effective_zero_fixed(mut value: f64, spec: &FixedSpec) -> bool {
    // Match `format_fixed` scaling + rounding behavior.
    for _ in 0..spec.scale_commas {
        value /= 1000.0;
    }
    let max_frac = spec.frac_placeholders.len();
    round_to(value, max_frac) == 0.0
}

fn format_general(value: f64, options: &FormatOptions) -> String {
    // Excel stores numbers with 15 significant digits. The rendering rules for General are:
    // - Round to 15 significant digits
    // - Use fixed-point for "reasonable" magnitudes, otherwise scientific
    // - Trim insignificant trailing zeros
    // - Avoid displaying negative zero
    if !value.is_finite() {
        return value.to_string();
    }
    let rounded = round_to_significant_digits(value, 15);
    if rounded == 0.0 {
        return "0".to_string();
    }

    let abs = rounded.abs();
    let exponent = decimal_exponent(abs);
    // Empirically Excel switches to scientific notation outside this exponent window.
    let use_scientific = exponent >= 11 || exponent <= -10;
    let mut s = if use_scientific {
        format_general_scientific(rounded, exponent)
    } else {
        format_general_decimal(rounded, exponent)
    };

    if options.locale.decimal_sep != '.' {
        s = s.replace('.', &options.locale.decimal_sep.to_string());
    }
    s
}

fn round_to_significant_digits(value: f64, digits: i32) -> f64 {
    if value == 0.0 {
        return 0.0;
    }
    let abs = value.abs();
    let exp = decimal_exponent(abs);
    let scale = digits - 1 - exp;
    // For extreme exponents this scaling factor can overflow/underflow. In those cases return the
    // value unchanged; further rounding isn't meaningful at f64 precision anyway.
    if scale > 308 || scale < -308 {
        return value;
    }
    let factor = 10_f64.powi(scale);
    (value * factor).round() / factor
}

fn decimal_exponent(abs: f64) -> i32 {
    if abs == 0.0 || !abs.is_finite() {
        return 0;
    }
    // Using `log10` directly is prone to boundary errors near powers of 10 because of floating
    // rounding. Adjust the initial estimate until 10^exp <= abs < 10^(exp+1).
    let mut exp = abs.log10().floor() as i32;
    let mut pow = 10_f64.powi(exp);

    // Handle overflow/underflow from powi.
    if pow.is_infinite() {
        while pow.is_infinite() {
            exp -= 1;
            pow = 10_f64.powi(exp);
        }
    } else if pow == 0.0 {
        while pow == 0.0 {
            exp += 1;
            pow = 10_f64.powi(exp);
        }
    }

    while abs < pow {
        exp -= 1;
        pow /= 10.0;
    }
    while abs >= pow * 10.0 {
        exp += 1;
        pow *= 10.0;
    }
    exp
}

fn format_general_decimal(value: f64, exponent: i32) -> String {
    // `exponent` is floor(log10(abs(value))).
    // Render with enough fractional digits to preserve 15 significant digits, then trim.
    let decimals = (15 - 1 - exponent).max(0) as usize;
    let mut s = format!("{:.*}", decimals, value);
    trim_trailing_zeros(&mut s);
    s
}

fn format_general_scientific(value: f64, exponent: i32) -> String {
    // Render mantissa with up to 15 significant digits and an exponent with at least 2 digits.
    // Use `E` (upper) like Excel.
    let mut exp = exponent;
    let mut mantissa = value / 10_f64.powi(exp);
    mantissa = round_to_significant_digits(mantissa, 15);

    if mantissa.abs() >= 10.0 {
        mantissa /= 10.0;
        exp += 1;
    }

    let sign = if exp < 0 { '-' } else { '+' };
    let exp_abs = exp.abs();
    let exp_width = std::cmp::max(2, exp_abs.to_string().len());

    let mut mantissa_str = format!("{:.14}", mantissa);
    trim_trailing_zeros(&mut mantissa_str);

    format!(
        "{}E{}{:0width$}",
        mantissa_str,
        sign,
        exp_abs,
        width = exp_width
    )
}

fn trim_trailing_zeros(s: &mut String) {
    if let Some(dot) = s.find('.') {
        while s.ends_with('0') {
            s.pop();
        }
        if s.ends_with('.') && s.len() == dot + 1 {
            s.pop();
        }
    }
}

fn scan_outside_quotes(s: &str, needle: char) -> (usize, Vec<usize>) {
    let mut in_quotes = false;
    let mut escape = false;
    let mut in_brackets = false;
    let mut count = 0;
    let mut positions = Vec::new();
    for (idx, ch) in s.char_indices() {
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
            _ if ch == needle => {
                count += 1;
                positions.push(idx);
            }
            _ => {}
        }
    }
    (count, positions)
}

fn find_placeholder_span(s: &str) -> (Option<usize>, Option<usize>) {
    let mut in_quotes = false;
    let mut escape = false;
    let mut in_brackets = false;
    let mut first: Option<usize> = None;
    let mut last: Option<usize> = None;

    for (idx, ch) in s.char_indices() {
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
            '0' | '#' | '?' => {
                if first.is_none() {
                    first = Some(idx);
                }
                // `idx` is at the start of the placeholder; update last to the end of it.
                last = Some(idx + ch.len_utf8());
            }
            _ => {}
        }
    }

    (first, last)
}

#[derive(Debug, Clone)]
struct FixedSpec {
    int_placeholders: Vec<PlaceholderKind>,
    frac_placeholders: Vec<PlaceholderKind>,
    grouping: Option<GroupingSpec>,
    scale_commas: usize,
    has_decimal_point: bool,
}

#[derive(Debug, Clone, Copy)]
struct GroupingSpec {
    primary: usize,
    secondary: usize,
}

#[derive(Debug, Clone, Copy, PartialEq, Eq)]
enum PlaceholderKind {
    Zero,
    Hash,
    Question,
}

fn parse_fixed(number_raw: &str) -> FixedSpec {
    // Scaling commas can appear:
    // - after all digit placeholders (e.g. `0.0,`)
    // - after the integer portion but before the decimal separator (e.g. `0,.0`)
    //
    // Both divide the value by 1000 per comma.
    let mut raw = number_raw.to_string();
    let mut scale_commas = 0;
    while raw.ends_with(',') {
        raw.pop();
        scale_commas += 1;
    }

    let mut in_quotes = false;
    let mut escape = false;
    let mut in_brackets = false;
    let mut decimal_pos: Option<usize> = None;
    for (idx, ch) in raw.char_indices() {
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
                decimal_pos = Some(idx);
                break;
            }
            _ => {}
        }
    }

    let (int_pat_raw, frac_pat) = match decimal_pos {
        Some(pos) => (&raw[..pos], &raw[pos + 1..]),
        None => (raw.as_str(), ""),
    };

    // Count scaling commas at the end of the integer portion, even if fractional placeholders
    // exist after the decimal point.
    let mut int_cut = int_pat_raw.len();
    while int_cut > 0 && int_pat_raw.as_bytes()[int_cut - 1] == b',' {
        int_cut -= 1;
        scale_commas += 1;
    }
    let int_pat = &int_pat_raw[..int_cut];

    let grouping = parse_grouping(int_pat);
    let int_placeholders = parse_placeholders(int_pat);
    let frac_placeholders = parse_placeholders(frac_pat);

    FixedSpec {
        int_placeholders,
        frac_placeholders,
        grouping,
        scale_commas,
        has_decimal_point: decimal_pos.is_some(),
    }
}

fn parse_grouping(int_pat: &str) -> Option<GroupingSpec> {
    let mut in_quotes = false;
    let mut escape = false;
    let mut in_brackets = false;
    let mut tokens = Vec::new();
    for ch in int_pat.chars() {
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
            '0' | '#' | '?' => tokens.push(1u8),
            ',' => tokens.push(0u8),
            _ => {}
        }
    }

    if !tokens.iter().any(|t| *t == 0) {
        return None;
    }

    let mut group_sizes: Vec<usize> = Vec::new();
    let mut count = 0usize;
    for t in tokens.into_iter().rev() {
        if t == 1 {
            count += 1;
        } else {
            if count > 0 {
                group_sizes.push(count);
            }
            count = 0;
        }
    }

    let primary = group_sizes.first().copied()?;
    let secondary = group_sizes.get(1).copied().unwrap_or(primary);
    Some(GroupingSpec { primary, secondary })
}

fn parse_placeholders(pat: &str) -> Vec<PlaceholderKind> {
    pat.chars()
        .filter_map(|c| match c {
            '0' => Some(PlaceholderKind::Zero),
            '#' => Some(PlaceholderKind::Hash),
            '?' => Some(PlaceholderKind::Question),
            _ => None,
        })
        .collect()
}

fn format_fixed(mut value: f64, spec: &FixedSpec, options: &FormatOptions) -> String {
    // Scaling commas divide by 1000 per comma.
    for _ in 0..spec.scale_commas {
        value /= 1000.0;
    }

    let max_frac = spec.frac_placeholders.len();

    // Round to max fraction digits.
    let rounded = round_to(value, max_frac);

    let (int_digits, frac_digits) = if max_frac == 0 {
        (format!("{:.0}", rounded), String::new())
    } else {
        let s = format!("{:.*}", max_frac, rounded);
        let mut split = s.splitn(2, '.');
        (
            split.next().unwrap_or("").to_string(),
            split.next().unwrap_or("").to_string(),
        )
    };

    let int_part = apply_int_placeholders(&int_digits, &spec.int_placeholders, value.abs() < 1.0);
    let int_part = if let Some(grouping) = spec.grouping {
        if !int_part.trim().is_empty() {
            // Grouping should ignore leading spaces produced by `?`.
            group_digits_with_padding(&int_part, options.locale.thousands_sep, grouping)
        } else {
            int_part
        }
    } else {
        // Grouping should ignore leading spaces produced by `?`.
        int_part
    };

    let frac_part = apply_frac_placeholders(&frac_digits, &spec.frac_placeholders);

    let mut out = String::new();
    out.push_str(&int_part);

    if spec.has_decimal_point {
        let show_decimal = spec.frac_placeholders.is_empty()
            || !frac_part.is_empty()
            || spec
                .frac_placeholders
                .iter()
                .any(|k| matches!(k, PlaceholderKind::Zero | PlaceholderKind::Question));
        if show_decimal {
            out.push(options.locale.decimal_sep);
        }
    }
    out.push_str(&frac_part);

    out
}

fn round_to(value: f64, decimals: usize) -> f64 {
    if decimals == 0 {
        return value.round();
    }
    let factor = 10_f64.powi(decimals as i32);
    (value * factor).round() / factor
}

fn group_digits_with_padding(int_part: &str, sep: char, spec: GroupingSpec) -> String {
    let trimmed = int_part.trim_start_matches(' ');
    let pad_len = int_part.len() - trimmed.len();
    let mut grouped = if trimmed.is_empty() {
        String::new()
    } else {
        group_digits(trimmed, sep, spec)
    };
    if pad_len > 0 {
        let mut out = String::new();
        let _ = out.try_reserve(pad_len + grouped.len());
        for _ in 0..pad_len {
            out.push(' ');
        }
        out.push_str(&grouped);
        grouped = out;
    }
    grouped
}

fn group_digits(int_part: &str, sep: char, spec: GroupingSpec) -> String {
    let primary = spec.primary.max(1);
    let secondary = spec.secondary.max(1);
    let len = int_part.len();
    if len <= primary {
        return int_part.to_string();
    }

    let mut groups: Vec<&str> = Vec::new();
    let mut idx = len;
    let mut first = true;
    while idx > 0 {
        let size = if first { primary } else { secondary };
        let start = idx.saturating_sub(size);
        groups.push(&int_part[start..idx]);
        idx = start;
        first = false;
    }
    groups.reverse();

    let mut out = String::new();
    let _ = out.try_reserve(len + groups.len().saturating_sub(1));
    for (i, group) in groups.iter().enumerate() {
        if i > 0 {
            out.push(sep);
        }
        out.push_str(group);
    }
    out
}

fn apply_int_placeholders(digits: &str, placeholders: &[PlaceholderKind], is_less_than_one: bool) -> String {
    if placeholders.is_empty() {
        return String::new();
    }

    // For values < 1, Excel suppresses the leading "0" when the integer placeholders are optional.
    // We treat the integer digits as empty in that case and let placeholders decide what to output.
    let digits = if is_less_than_one && digits == "0" { "" } else { digits };

    let digit_chars: Vec<char> = digits.chars().collect();
    let placeholder_len = placeholders.len();

    // Digits beyond the placeholder width are still displayed.
    let extra = digit_chars.len().saturating_sub(placeholder_len);
    let mut out = String::new();
    for ch in digit_chars.iter().take(extra) {
        out.push(*ch);
    }

    let tail_digits = &digit_chars[extra..];
    let mut tail_out: Vec<char> = Vec::new();
    let _ = tail_out.try_reserve_exact(placeholder_len);
    let mut digit_idx: i32 = tail_digits.len() as i32 - 1;

    for kind in placeholders.iter().rev() {
        let digit = if digit_idx >= 0 {
            let c = tail_digits[digit_idx as usize];
            digit_idx -= 1;
            Some(c)
        } else {
            None
        };

        match kind {
            PlaceholderKind::Zero => tail_out.push(digit.unwrap_or('0')),
            PlaceholderKind::Hash => {
                if let Some(c) = digit {
                    tail_out.push(c);
                }
            }
            PlaceholderKind::Question => {
                if let Some(c) = digit {
                    tail_out.push(c);
                } else {
                    tail_out.push(' ');
                }
            }
        }
    }

    tail_out.reverse();
    out.extend(tail_out);
    out
}

fn apply_frac_placeholders(digits: &str, placeholders: &[PlaceholderKind]) -> String {
    if placeholders.is_empty() {
        return String::new();
    }

    let mut digit_chars: Vec<char> = digits.chars().collect();
    while digit_chars.len() < placeholders.len() {
        digit_chars.push('0');
    }

    // Determine how many trailing digits are insignificant zeros for optional placeholders.
    let mut cut = digit_chars.len();
    while cut > 0 {
        let idx = cut - 1;
        let kind = placeholders.get(idx).copied().unwrap_or(PlaceholderKind::Hash);
        let digit = digit_chars[idx];
        match kind {
            PlaceholderKind::Zero => break,
            PlaceholderKind::Hash => {
                if digit == '0' {
                    cut -= 1;
                } else {
                    break;
                }
            }
            PlaceholderKind::Question => {
                if digit == '0' {
                    cut -= 1;
                } else {
                    break;
                }
            }
        }
    }

    let mut out = String::new();
    let _ = out.try_reserve(placeholders.len());
    for (idx, kind) in placeholders.iter().enumerate() {
        let digit = digit_chars.get(idx).copied().unwrap_or('0');
        if idx >= cut {
            match kind {
                PlaceholderKind::Zero => out.push('0'),
                PlaceholderKind::Hash => {}
                PlaceholderKind::Question => out.push(' '),
            }
        } else {
            out.push(digit);
        }
    }
    out
}

#[derive(Debug, Clone)]
struct ScientificSpec {
    mantissa: FixedSpec,
    exp_width: usize,
    exp_sign_always: bool,
    e_char: char,
}

fn parse_scientific(number_raw: &str) -> Option<ScientificSpec> {
    let mut in_quotes = false;
    let mut escape = false;
    let mut in_brackets = false;
    let mut e_pos: Option<(usize, char)> = None;
    for (idx, ch) in number_raw.char_indices() {
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
            'E' | 'e' => {
                e_pos = Some((idx, ch));
                break;
            }
            _ => {}
        }
    }

    let (e_idx, e_char) = e_pos?;
    let mantissa_raw = &number_raw[..e_idx];
    let exponent_raw = &number_raw[e_idx + 1..];

    let mut exp_sign_always = false;
    let mut rest = exponent_raw;
    if let Some(first) = exponent_raw.chars().next() {
        if first == '+' {
            exp_sign_always = true;
            rest = &exponent_raw[1..];
        } else if first == '-' {
            exp_sign_always = false;
            rest = &exponent_raw[1..];
        }
    }

    // Count placeholders in exponent.
    let exp_width = rest.chars().filter(|c| matches!(c, '0' | '#' | '?')).count();
    if exp_width == 0 {
        return None;
    }

    Some(ScientificSpec {
        mantissa: parse_fixed(mantissa_raw),
        exp_width,
        exp_sign_always,
        e_char,
    })
}

fn format_scientific(value: f64, spec: &ScientificSpec, options: &FormatOptions) -> String {
    if value == 0.0 {
        let mantissa = format_fixed(0.0, &spec.mantissa, options);
        let exp = format_exponent(0, spec, options);
        return format!("{mantissa}{}{exp}", spec.e_char);
    }

    let exponent = value.abs().log10().floor() as i32;
    let pow = 10_f64.powi(exponent);
    let mut mantissa = value / pow;

    // Apply rounding to mantissa based on the mantissa spec.
    let mantissa_frac = spec.mantissa.frac_placeholders.len();
    mantissa = round_to(mantissa, mantissa_frac);

    // Rounding can bump mantissa to 10.0; normalize.
    let mut exp = exponent;
    if mantissa >= 10.0 {
        mantissa /= 10.0;
        exp += 1;
    }

    let mantissa_str = format_fixed(mantissa, &spec.mantissa, options);
    let exp_str = format_exponent(exp, spec, options);

    format!("{mantissa_str}{}{exp_str}", spec.e_char)
}

fn format_exponent(exp: i32, spec: &ScientificSpec, _options: &FormatOptions) -> String {
    let sign = if exp < 0 {
        "-"
    } else if spec.exp_sign_always {
        "+"
    } else {
        ""
    };

    let abs = exp.abs();
    let digits = format!("{:0width$}", abs, width = spec.exp_width);
    format!("{sign}{digits}")
}

#[derive(Debug, Clone)]
struct FractionSpec {
    int_spec: FixedSpec,
    int_has_sep: bool,
    num_width: usize,
    den_width: usize,
    den_fixed: Option<i64>,
}

fn parse_fraction(number_raw: &str) -> Option<FractionSpec> {
    // Find a `/` outside quotes and bracket tokens.
    let mut in_quotes = false;
    let mut escape = false;
    let mut in_brackets = false;
    let mut slash_idx: Option<usize> = None;

    for (idx, ch) in number_raw.char_indices() {
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

    let slash_idx = slash_idx?;
    let left = &number_raw[..slash_idx];
    let right = &number_raw[slash_idx + 1..];

    let den_width = right.chars().filter(|c| matches!(c, '0' | '#' | '?')).count();
    let den_fixed = if den_width == 0 {
        right.trim().parse::<i64>().ok()
    } else {
        None
    };
    if den_width == 0 && den_fixed.is_none() {
        return None;
    }

    // Split left into integer + numerator by finding the last space before the slash.
    let (int_part, num_part, int_has_sep) = if let Some(space_idx) = left.rfind(' ') {
        let num_candidate = &left[space_idx + 1..];
        if num_candidate.chars().any(|c| matches!(c, '0' | '#' | '?')) {
            (&left[..space_idx], num_candidate, true)
        } else {
            ("", left, false)
        }
    } else {
        ("", left, false)
    };

    let num_width = num_part.chars().filter(|c| matches!(c, '0' | '#' | '?')).count();
    if num_width == 0 {
        return None;
    }

    Some(FractionSpec {
        int_spec: parse_fixed(int_part.trim()),
        int_has_sep,
        num_width,
        den_width,
        den_fixed,
    })
}

fn format_fraction(value: f64, spec: &FractionSpec, options: &FormatOptions) -> String {
    if value == 0.0 {
        return "0".to_string();
    }

    let int_value = value.floor();
    let frac = value - int_value;
    let mut int_i64 = int_value as i64;

    let (mut num, mut den) = if let Some(fixed) = spec.den_fixed {
        let d = fixed.max(1);
        let n = (frac * d as f64).round() as i64;
        (n, d)
    } else {
        let max_den = 10_i64.pow(spec.den_width as u32) - 1;
        let max_den = max_den.max(1).min(10_000); // avoid pathological scans

        let mut best_n = 0_i64;
        let mut best_d = 1_i64;
        let mut best_err = f64::INFINITY;
        for d in 1..=max_den {
            let n = (frac * d as f64).round() as i64;
            let err = (frac - (n as f64) / (d as f64)).abs();
            if err < best_err {
                best_err = err;
                best_n = n;
                best_d = d;
                if err == 0.0 {
                    break;
                }
            }
        }
        (best_n, best_d)
    };

    if num == 0 {
        let int_str = format_fixed(int_value, &spec.int_spec, options);
        return if int_str.is_empty() {
            "0".to_string()
        } else {
            int_str
        };
    }

    if num == den {
        int_i64 += 1;
        num = 0;
    }

    if num == 0 {
        let int_str = format_fixed(int_i64 as f64, &spec.int_spec, options);
        return if int_str.is_empty() {
            "0".to_string()
        } else {
            int_str
        };
    }

    let g = gcd(num.abs(), den);
    num /= g;
    den /= g;

    let int_str = format_fixed(int_i64 as f64, &spec.int_spec, options);
    let num_str = pad_left(num.to_string(), spec.num_width, ' ');
    let den_str = pad_left(den.to_string(), spec.den_width.max(den.to_string().len()), ' ');

    if int_str.is_empty() {
        format!("{num_str}/{den_str}")
    } else if spec.int_has_sep {
        format!("{int_str} {num_str}/{den_str}")
    } else {
        format!("{int_str}{num_str}/{den_str}")
    }
}

fn pad_left(mut s: String, width: usize, pad: char) -> String {
    while s.len() < width {
        s.insert(0, pad);
    }
    s
}

fn gcd(mut a: i64, mut b: i64) -> i64 {
    while b != 0 {
        let r = a % b;
        a = b;
        b = r;
    }
    a.abs().max(1)
}
