use crate::{literal, FormatOptions};

pub(crate) fn pattern_is_text(pattern: &str) -> bool {
    let (at_count, _) = scan_outside_quotes(pattern, '@');
    if at_count == 0 {
        return false;
    }

    let (first, last) = find_placeholder_span(pattern);
    first.is_none() && last.is_none()
}

pub(crate) fn format_number(value: f64, pattern: &str, auto_negative_sign: bool, options: &FormatOptions) -> String {
    if pattern.trim().eq_ignore_ascii_case("general") {
        return format_general(value, options);
    }

    if !value.is_finite() {
        return value.to_string();
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
        let mut s = format!("{prefix}{out}{suffix}");
        if value < 0.0 && auto_negative_sign {
            s.insert(0, '-');
        }
        return s;
    }

    if let Some(spec) = parse_fraction(number_raw) {
        let out = format_fraction(v, &spec, options);
        let mut s = format!("{prefix}{out}{suffix}");
        if value < 0.0 && auto_negative_sign {
            s.insert(0, '-');
        }
        return s;
    }

    let spec = parse_fixed(number_raw);
    let out = format_fixed(v, &spec, options);
    let mut s = format!("{prefix}{out}{suffix}");
    if value < 0.0 && auto_negative_sign {
        s.insert(0, '-');
    }
    s
}

fn format_general(value: f64, options: &FormatOptions) -> String {
    if value == 0.0 {
        return "0".to_string();
    }

    let mut s = value.to_string();
    if s.contains('e') {
        s = s.replace('e', "E");
    }
    if options.locale.decimal_sep != '.' {
        s = s.replace('.', &options.locale.decimal_sep.to_string());
    }
    s
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
    min_int: usize,
    int_placeholders: usize,
    min_frac: usize,
    max_frac: usize,
    grouping: bool,
    scale_commas: usize,
    has_decimal_point: bool,
}

fn parse_fixed(number_raw: &str) -> FixedSpec {
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

    let (int_pat, frac_pat) = match decimal_pos {
        Some(pos) => (&raw[..pos], &raw[pos + 1..]),
        None => (raw.as_str(), ""),
    };

    let grouping = int_pat.contains(',');
    let int_placeholders = int_pat.chars().filter(|c| matches!(c, '0' | '#' | '?')).count();
    let min_int = int_pat.chars().filter(|c| *c == '0').count();
    let max_frac = frac_pat.chars().filter(|c| matches!(c, '0' | '#' | '?')).count();
    let min_frac = frac_pat.chars().filter(|c| *c == '0').count();

    FixedSpec {
        min_int,
        int_placeholders,
        min_frac,
        max_frac,
        grouping,
        scale_commas,
        has_decimal_point: decimal_pos.is_some(),
    }
}

fn format_fixed(mut value: f64, spec: &FixedSpec, options: &FormatOptions) -> String {
    // Scaling commas divide by 1000 per comma.
    for _ in 0..spec.scale_commas {
        value /= 1000.0;
    }

    // Round to max fraction digits.
    let rounded = round_to(value, spec.max_frac);

    let (mut int_part, mut frac_part) = if spec.max_frac == 0 {
        (format!("{:.0}", rounded), String::new())
    } else {
        let s = format!("{:.*}", spec.max_frac, rounded);
        let mut split = s.splitn(2, '.');
        (
            split.next().unwrap_or("").to_string(),
            split.next().unwrap_or("").to_string(),
        )
    };

    // Optional integer digits.
    if spec.int_placeholders == 0 {
        // No integer placeholders at all; unusual but legal.
        int_part.clear();
    } else if spec.min_int == 0 && int_part == "0" && value.abs() < 1.0 {
        // Common optional-digit pattern: "#" should show blank for values < 1.
        // (Excel behaviour for 0 is blank; for 0.5 it shows nothing before the decimal.)
        int_part.clear();
    }

    while int_part.len() < spec.min_int {
        int_part.insert(0, '0');
    }

    if spec.grouping && !int_part.is_empty() {
        int_part = group_thousands(&int_part, options.locale.thousands_sep);
    }

    if spec.max_frac > 0 {
        while frac_part.len() > spec.min_frac && frac_part.ends_with('0') {
            frac_part.pop();
        }
    }

    let mut out = String::new();
    out.push_str(&int_part);

    if spec.has_decimal_point && (spec.max_frac == 0 || !frac_part.is_empty()) {
        out.push(options.locale.decimal_sep);
    }
    if !frac_part.is_empty() {
        out.push_str(&frac_part);
    }

    out
}

fn round_to(value: f64, decimals: usize) -> f64 {
    if decimals == 0 {
        return value.round();
    }
    let factor = 10_f64.powi(decimals as i32);
    (value * factor).round() / factor
}

fn group_thousands(int_part: &str, sep: char) -> String {
    let mut out = String::new();
    let bytes = int_part.as_bytes();
    let len = bytes.len();
    for (i, ch) in int_part.chars().enumerate() {
        let pos_from_end = len - i;
        out.push(ch);
        if pos_from_end > 1 && pos_from_end % 3 == 1 {
            out.push(sep);
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
    mantissa = round_to(mantissa, spec.mantissa.max_frac);

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
