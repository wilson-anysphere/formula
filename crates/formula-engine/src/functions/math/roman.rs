use crate::error::{ExcelError, ExcelResult};

/// ROMAN(number, [form])
///
/// Convert an integer in `[0, 3999]` to an Excel-compatible Roman numeral string.
///
/// Excel supports a `form` argument (0..=4) that generates progressively more "concise"
/// (relaxed) Roman numeral forms. These match the documented Excel outputs, e.g.:
/// - `ROMAN(499,0) == "CDXCIX"`
/// - `ROMAN(499,1) == "LDVLIV"`
/// - `ROMAN(499,2) == "XDIX"`
/// - `ROMAN(499,3) == "VDIV"`
/// - `ROMAN(499,4) == "ID"`
pub fn roman(number: i64, form: Option<i64>) -> ExcelResult<String> {
    let form = form.unwrap_or(0);
    let form_u8 = u8::try_from(form).map_err(|_| ExcelError::Value)?;
    let tokens = roman_tokens(form_u8)?;

    if !(0..=3999).contains(&number) {
        return Err(ExcelError::Value);
    }
    if number == 0 {
        return Ok(String::new());
    }

    let mut remaining = number as i32;
    let mut out = String::new();
    for &(value, symbol) in tokens {
        while remaining >= value {
            out.push_str(symbol);
            remaining -= value;
        }
    }

    debug_assert_eq!(remaining, 0);
    Ok(out)
}

/// ARABIC(text)
///
/// Parse an Excel-compatible Roman numeral and return its value in `[0, 3999]`.
///
/// The parser is case-insensitive and accepts any Roman numeral string that can be
/// produced by [`roman`] for `form` 0..=4.
pub fn arabic(text: &str) -> ExcelResult<i64> {
    let trimmed = text.trim();
    if trimmed.is_empty() {
        // Excel's ROMAN(0) yields an empty string; treat empty as 0 for round-tripping.
        return Ok(0);
    }

    let mut total: i64 = 0;
    let bytes = trimmed.as_bytes();
    for i in 0..bytes.len() {
        let v = match bytes[i].to_ascii_uppercase() {
            b'I' => 1,
            b'V' => 5,
            b'X' => 10,
            b'L' => 50,
            b'C' => 100,
            b'D' => 500,
            b'M' => 1000,
            _ => return Err(ExcelError::Value),
        };
        let next = if let Some(&b) = bytes.get(i + 1) {
            match b.to_ascii_uppercase() {
                b'I' => 1,
                b'V' => 5,
                b'X' => 10,
                b'L' => 50,
                b'C' => 100,
                b'D' => 500,
                b'M' => 1000,
                _ => return Err(ExcelError::Value),
            }
        } else {
            0
        };
        if v < next {
            total -= v as i64;
        } else {
            total += v as i64;
        }
    }

    if !(0..=3999).contains(&total) {
        return Err(ExcelError::Value);
    }

    // Validate that the input is a canonical Excel Roman numeral for at least one
    // supported form.
    for form in 0..=4 {
        if roman(total, Some(form))?.eq_ignore_ascii_case(trimmed) {
            return Ok(total);
        }
    }

    Err(ExcelError::Value)
}

type Token = (i32, &'static str);

fn roman_tokens(form: u8) -> ExcelResult<&'static [Token]> {
    match form {
        0 => Ok(&TOKENS_FORM0),
        1 => Ok(&TOKENS_FORM1),
        2 => Ok(&TOKENS_FORM2),
        3 => Ok(&TOKENS_FORM3),
        4 => Ok(&TOKENS_FORM4),
        _ => Err(ExcelError::Value),
    }
}

// Form 0 ("classic") uses the standard modern Roman numeral rules.
const TOKENS_FORM0: [Token; 13] = [
    (1000, "M"),
    (900, "CM"),
    (500, "D"),
    (400, "CD"),
    (100, "C"),
    (90, "XC"),
    (50, "L"),
    (40, "XL"),
    (10, "X"),
    (9, "IX"),
    (5, "V"),
    (4, "IV"),
    (1, "I"),
];

// Form 1 adds the "V before L/C" and "L before D/M" relaxed subtractive forms.
const TOKENS_FORM1: [Token; 17] = [
    (1000, "M"),
    (950, "LM"),
    (900, "CM"),
    (500, "D"),
    (450, "LD"),
    (400, "CD"),
    (100, "C"),
    (95, "VC"),
    (90, "XC"),
    (50, "L"),
    (45, "VL"),
    (40, "XL"),
    (10, "X"),
    (9, "IX"),
    (5, "V"),
    (4, "IV"),
    (1, "I"),
];

// Form 2 adds the "X before D/M" relaxed subtractive forms.
const TOKENS_FORM2: [Token; 19] = [
    (1000, "M"),
    (990, "XM"),
    (950, "LM"),
    (900, "CM"),
    (500, "D"),
    (490, "XD"),
    (450, "LD"),
    (400, "CD"),
    (100, "C"),
    (95, "VC"),
    (90, "XC"),
    (50, "L"),
    (45, "VL"),
    (40, "XL"),
    (10, "X"),
    (9, "IX"),
    (5, "V"),
    (4, "IV"),
    (1, "I"),
];

// Form 3 adds the "V before D/M" relaxed subtractive forms.
const TOKENS_FORM3: [Token; 21] = [
    (1000, "M"),
    (995, "VM"),
    (990, "XM"),
    (950, "LM"),
    (900, "CM"),
    (500, "D"),
    (495, "VD"),
    (490, "XD"),
    (450, "LD"),
    (400, "CD"),
    (100, "C"),
    (95, "VC"),
    (90, "XC"),
    (50, "L"),
    (45, "VL"),
    (40, "XL"),
    (10, "X"),
    (9, "IX"),
    (5, "V"),
    (4, "IV"),
    (1, "I"),
];

// Form 4 adds the "I before L/C/D/M" relaxed subtractive forms.
const TOKENS_FORM4: [Token; 25] = [
    (1000, "M"),
    (999, "IM"),
    (995, "VM"),
    (990, "XM"),
    (950, "LM"),
    (900, "CM"),
    (500, "D"),
    (499, "ID"),
    (495, "VD"),
    (490, "XD"),
    (450, "LD"),
    (400, "CD"),
    (100, "C"),
    (99, "IC"),
    (95, "VC"),
    (90, "XC"),
    (50, "L"),
    (49, "IL"),
    (45, "VL"),
    (40, "XL"),
    (10, "X"),
    (9, "IX"),
    (5, "V"),
    (4, "IV"),
    (1, "I"),
];
