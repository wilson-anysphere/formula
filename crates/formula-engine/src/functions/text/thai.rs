use crate::error::{ExcelError, ExcelResult};
use unicode_segmentation::UnicodeSegmentation;

const THAI_DIGITS: [char; 10] = ['๐', '๑', '๒', '๓', '๔', '๕', '๖', '๗', '๘', '๙'];

const THAI_DIGIT_WORDS: [&str; 10] = [
    "ศูนย์",
    "หนึ่ง",
    "สอง",
    "สาม",
    "สี่",
    "ห้า",
    "หก",
    "เจ็ด",
    "แปด",
    "เก้า",
];

/// BAHTTEXT(number)
///
/// Convert a numeric value to Thai Baht text (baht + satang), using Thai number words and the
/// fixed suffixes `บาท`, `สตางค์`, and `ถ้วน`.
///
/// Locale note: This function is Thai-specific and does not depend on workbook locale.
pub fn bahttext(number: f64) -> ExcelResult<String> {
    if !number.is_finite() {
        return Err(ExcelError::Num);
    }

    // Excel's BAHTTEXT has a practical range limit; keep it bounded so we can safely convert
    // to integer satang for deterministic formatting.
    const MAX: f64 = 999_999_999_999.99;
    if number.abs() > MAX {
        return Err(ExcelError::Value);
    }

    let negative = number.is_sign_negative();
    let abs = number.abs();

    let satang_total = (abs * 100.0).round();
    if satang_total < 0.0 || satang_total > (i64::MAX as f64) {
        return Err(ExcelError::Value);
    }
    let satang_total = satang_total as i64;

    let baht = satang_total / 100;
    let satang = (satang_total % 100) as u8;

    let mut out = String::new();
    if negative && (baht != 0 || satang != 0) {
        out.push_str("ลบ");
    }

    out.push_str(&thai_integer_to_words(baht as u64));
    out.push_str("บาท");

    if satang == 0 {
        out.push_str("ถ้วน");
    } else {
        out.push_str(&thai_group_to_words(satang as u32));
        out.push_str("สตางค์");
    }

    Ok(out)
}

/// THAIDIGIT(text)
///
/// Replace ASCII digits 0-9 with Thai digits ๐-๙.
pub fn thai_digit(text: &str) -> String {
    text.chars()
        .map(|ch| match ch {
            '0'..='9' => {
                let idx = (ch as u8 - b'0') as usize;
                THAI_DIGITS[idx]
            }
            _ => ch,
        })
        .collect()
}

/// ISTHAIDIGIT(text)
pub fn is_thai_digit(text: &str) -> bool {
    if text.is_empty() {
        return false;
    }
    text.chars().all(|ch| matches!(ch, '๐'..='๙'))
}

/// THAISTRINGLENGTH(text)
///
/// Return the display-length of a Thai string. This uses Unicode grapheme cluster counting,
/// which matches how Thai combining marks (tone marks / diacritics) are rendered without taking
/// additional character cells.
pub fn thai_string_length(text: &str) -> usize {
    UnicodeSegmentation::graphemes(text, true).count()
}

/// THAINUMSTRING(number)
///
/// Convert a number into its Thai-digit string representation.
pub fn thainumstring(number: f64) -> ExcelResult<String> {
    if !number.is_finite() {
        return Err(ExcelError::Num);
    }
    Ok(thai_digit(&format_number_fixed_trim(number)))
}

/// THAINUMSOUND(number)
///
/// Convert a number to its Thai reading ("sound") using Thai number words.
pub fn thainumsound(number: f64) -> ExcelResult<String> {
    if !number.is_finite() {
        return Err(ExcelError::Num);
    }

    let negative = number.is_sign_negative();
    let abs = number.abs();

    let repr = format_number_fixed_trim(abs);
    let (int_str, frac_str) = match repr.split_once('.') {
        Some((a, b)) => (a, Some(b)),
        None => (repr.as_str(), None),
    };

    let int_val: u64 = int_str.parse().map_err(|_| ExcelError::Value)?;
    let mut out = String::new();
    if negative && abs != 0.0 {
        out.push_str("ลบ");
    }

    out.push_str(&thai_integer_to_words(int_val));

    if let Some(frac) = frac_str {
        if !frac.is_empty() {
            out.push_str("จุด");
            for ch in frac.chars() {
                let digit = ch.to_digit(10).ok_or(ExcelError::Value)? as usize;
                out.push_str(THAI_DIGIT_WORDS[digit]);
            }
        }
    }

    Ok(out)
}

/// ROUNDBAHTDOWN(number)
pub fn roundbahtdown(number: f64) -> ExcelResult<f64> {
    if !number.is_finite() {
        return Err(ExcelError::Num);
    }
    if number == 0.0 {
        return Ok(0.0);
    }
    let scaled = number * 4.0;
    let q = if number.is_sign_negative() {
        // Match Excel ROUNDDOWN semantics: toward zero.
        scaled.ceil()
    } else {
        scaled.floor()
    };
    Ok(q / 4.0)
}

/// ROUNDBAHTUP(number)
pub fn roundbahtup(number: f64) -> ExcelResult<f64> {
    if !number.is_finite() {
        return Err(ExcelError::Num);
    }
    if number == 0.0 {
        return Ok(0.0);
    }
    let scaled = number * 4.0;
    let q = if number.is_sign_negative() {
        // Match Excel ROUNDUP semantics: away from zero.
        scaled.floor()
    } else {
        scaled.ceil()
    };
    Ok(q / 4.0)
}

fn thai_integer_to_words(n: u64) -> String {
    if n == 0 {
        return THAI_DIGIT_WORDS[0].to_string();
    }

    let mut groups = Vec::<u32>::new();
    let mut cur = n;
    while cur > 0 {
        groups.push((cur % 1_000_000) as u32);
        cur /= 1_000_000;
    }

    // groups are little-endian (lowest 6 digits first).
    let mut out = String::new();
    for (idx, group) in groups.iter().enumerate().rev() {
        let group_words = thai_group_to_words(*group);
        if !group_words.is_empty() {
            out.push_str(&group_words);
        }
        if idx > 0 {
            out.push_str("ล้าน");
        }
    }
    out
}

fn thai_group_to_words(group: u32) -> String {
    if group == 0 {
        return String::new();
    }

    let digits = [
        ((group / 100_000) % 10, "แสน"),
        ((group / 10_000) % 10, "หมื่น"),
        ((group / 1_000) % 10, "พัน"),
        ((group / 100) % 10, "ร้อย"),
        ((group / 10) % 10, "สิบ"),
        (group % 10, ""),
    ];

    let mut out = String::new();
    for (pos, unit) in digits {
        if pos == 0 {
            continue;
        }

        if unit == "สิบ" {
            match pos {
                1 => out.push_str("สิบ"),
                2 => out.push_str("ยี่สิบ"),
                _ => {
                    out.push_str(THAI_DIGIT_WORDS[pos as usize]);
                    out.push_str("สิบ");
                }
            }
            continue;
        }

        if unit.is_empty() {
            match pos {
                1 if !out.is_empty() => out.push_str("เอ็ด"),
                _ => out.push_str(THAI_DIGIT_WORDS[pos as usize]),
            }
            continue;
        }

        out.push_str(THAI_DIGIT_WORDS[pos as usize]);
        out.push_str(unit);
    }

    out
}

fn format_number_fixed_trim(number: f64) -> String {
    // Produce a non-scientific representation (up to 15 decimal places) and trim trailing zeros.
    let mut out = format!("{number:.15}");
    if let Some(dot) = out.find('.') {
        while out.ends_with('0') {
            out.pop();
        }
        if out.len() == dot + 1 {
            out.pop();
        }
    }
    if out == "-0" {
        out = "0".to_string();
    }
    out
}
