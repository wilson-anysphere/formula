use crate::{literal, FormatOptions};

pub(crate) fn format_number(value: f64, pattern: &str, auto_negative_sign: bool, options: &FormatOptions) -> String {
    if pattern.trim().eq_ignore_ascii_case("general") {
        return format_general(value, options);
    }

    if !value.is_finite() {
        return value.to_string();
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

    // Fraction (optional): currently not implemented; fall back to fixed number formatting.

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
        match ch {
            '"' => in_quotes = true,
            '\\' => escape = true,
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
        match ch {
            '"' => in_quotes = true,
            '\\' => escape = true,
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
        match ch {
            '"' => in_quotes = true,
            '\\' => escape = true,
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
        match ch {
            '"' => in_quotes = true,
            '\\' => escape = true,
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
