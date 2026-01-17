use crate::value::ErrorKind;

/// Excel's legacy base conversion functions (BIN2*, OCT2*, HEX2*, DEC2*) use fixed-width two's
/// complement representations:
/// - BIN*: 10 bits
/// - OCT*: 30 bits
/// - HEX*: 40 bits
///
/// Excel limits these strings to 10 digits for each radix. When the input contains exactly 10
/// digits, Excel interprets it as a two's complement signed integer with the corresponding bit
/// width. Shorter strings are interpreted as positive numbers.
#[derive(Debug, Clone, Copy, PartialEq, Eq)]
pub(crate) enum FixedBase {
    Bin,
    Oct,
    Hex,
}

impl FixedBase {
    pub(crate) const fn radix(self) -> u32 {
        match self {
            FixedBase::Bin => 2,
            FixedBase::Oct => 8,
            FixedBase::Hex => 16,
        }
    }

    pub(crate) const fn bits(self) -> u32 {
        match self {
            FixedBase::Bin => 10,
            FixedBase::Oct => 30,
            FixedBase::Hex => 40,
        }
    }

    pub(crate) const fn max_digits(self) -> usize {
        // Excel uses 10 digits for BIN/OCT/HEX in these conversion functions.
        10
    }

    pub(crate) const fn min_signed(self) -> i64 {
        -(1i64 << (self.bits() - 1))
    }

    pub(crate) const fn max_signed(self) -> i64 {
        (1i64 << (self.bits() - 1)) - 1
    }
}

pub(crate) fn fixed_base_to_decimal(text: &str, base: FixedBase) -> Result<i64, ErrorKind> {
    let trimmed = text.trim();
    if trimmed.is_empty() {
        return Err(ErrorKind::Num);
    }

    if trimmed.len() > base.max_digits() {
        return Err(ErrorKind::Num);
    }

    let radix = base.radix();
    let unsigned = parse_unsigned_fixed(trimmed, radix)?;
    let signed = if trimmed.len() == base.max_digits() {
        twos_complement_to_i64(unsigned, base.bits())
    } else {
        unsigned as i64
    };

    Ok(signed)
}

pub(crate) fn fixed_decimal_to_fixed_base(
    value: i64,
    base: FixedBase,
    places: Option<usize>,
) -> Result<String, ErrorKind> {
    if value < base.min_signed() || value > base.max_signed() {
        return Err(ErrorKind::Num);
    }

    if let Some(p) = places {
        if p == 0 || p > base.max_digits() {
            return Err(ErrorKind::Num);
        }
    }

    let radix = base.radix();
    if value < 0 {
        // Excel ignores the requested `places` and always returns a full-width (10 digit) two's
        // complement representation for negative values.
        let unsigned = i64_to_twos_complement(value, base.bits());
        let raw = to_radix_upper(unsigned, radix);
        return Ok(pad_left(&raw, base.max_digits(), '0'));
    }

    let unsigned = value as u64;
    let raw = to_radix_upper(unsigned, radix);
    match places {
        None => Ok(raw),
        Some(p) => {
            if raw.len() > p {
                return Err(ErrorKind::Num);
            }
            Ok(pad_left(&raw, p, '0'))
        }
    }
}

pub(crate) fn fixed_base_to_fixed_base(
    text: &str,
    src: FixedBase,
    dst: FixedBase,
    places: Option<usize>,
) -> Result<String, ErrorKind> {
    let value = fixed_base_to_decimal(text, src)?;
    fixed_decimal_to_fixed_base(value, dst, places)
}

pub(crate) fn decimal_from_text(text: &str, radix: u32) -> Result<f64, ErrorKind> {
    let radix = validate_radix(radix)?;

    let trimmed = text.trim();
    if trimmed.is_empty() {
        return Err(ErrorKind::Num);
    }

    let (negative, digits) = if let Some(rest) = trimmed.strip_prefix('-') {
        (true, rest.trim_start())
    } else if let Some(rest) = trimmed.strip_prefix('+') {
        (false, rest.trim_start())
    } else {
        (false, trimmed)
    };

    if digits.is_empty() {
        return Err(ErrorKind::Num);
    }

    let mut acc: u128 = 0;
    let radix_u128 = u128::from(radix);
    for ch in digits.chars() {
        let v = match digit_value(ch) {
            Some(v) => v,
            None => return Err(ErrorKind::Num),
        };
        if v >= radix {
            return Err(ErrorKind::Num);
        }
        acc = acc
            .checked_mul(radix_u128)
            .and_then(|v2| v2.checked_add(u128::from(v)))
            .ok_or(ErrorKind::Num)?;
    }

    // Excel numbers are IEEE 754 doubles; above 2^53 the integer result would not be exact.
    // Excel's documentation for BASE/DECIMAL constrains inputs so results remain representable.
    const MAX_SAFE_INT: u128 = (1u128 << 53) - 1;
    if acc > MAX_SAFE_INT {
        return Err(ErrorKind::Num);
    }

    let out = acc as f64;
    Ok(if negative { -out } else { out })
}

pub(crate) fn validate_radix(radix: u32) -> Result<u32, ErrorKind> {
    if (2..=36).contains(&radix) {
        Ok(radix)
    } else {
        Err(ErrorKind::Num)
    }
}

pub(crate) fn base_from_decimal(
    number: u64,
    radix: u32,
    min_length: Option<usize>,
) -> Result<String, ErrorKind> {
    let radix = validate_radix(radix)?;
    let mut out = to_radix_upper(number, radix);

    if let Some(min_len) = min_length {
        if min_len > 255 {
            return Err(ErrorKind::Num);
        }
        out = pad_left(&out, min_len, '0');
    }

    Ok(out)
}

fn twos_complement_to_i64(unsigned: u64, bits: u32) -> i64 {
    debug_assert!(bits > 0 && bits <= 63);
    let sign_bit = 1u64 << (bits - 1);
    if unsigned & sign_bit == 0 {
        unsigned as i64
    } else {
        (unsigned as i64) - (1i64 << bits)
    }
}

fn i64_to_twos_complement(value: i64, bits: u32) -> u64 {
    debug_assert!(bits > 0 && bits <= 63);
    if value >= 0 {
        value as u64
    } else {
        // value is already range-checked against the signed `bits` domain.
        let modulus = 1u64 << bits;
        (modulus as i64 + value) as u64
    }
}

fn parse_unsigned_fixed(text: &str, radix: u32) -> Result<u64, ErrorKind> {
    // `from_str_radix` accepts a leading '+'/'-' which Excel does not for these fixed-base
    // functions. Reject any sign explicitly.
    if text.starts_with('+') || text.starts_with('-') {
        return Err(ErrorKind::Num);
    }
    u64::from_str_radix(text, radix).map_err(|_| ErrorKind::Num)
}

fn pad_left(text: &str, width: usize, pad: char) -> String {
    if text.len() >= width {
        return text.to_string();
    }
    let mut out = String::with_capacity(width);
    for _ in 0..(width - text.len()) {
        out.push(pad);
    }
    out.push_str(text);
    out
}

fn digit_value(ch: char) -> Option<u32> {
    match ch {
        '0'..='9' => Some((ch as u32) - ('0' as u32)),
        'A'..='Z' => Some(10 + (ch as u32) - ('A' as u32)),
        'a'..='z' => Some(10 + (ch as u32) - ('a' as u32)),
        _ => None,
    }
}

fn digit_char(value: u32) -> char {
    debug_assert!(value < 36);
    match value {
        0..=9 => (b'0' + (value as u8)) as char,
        10..=35 => (b'A' + ((value - 10) as u8)) as char,
        _ => {
            debug_assert!(false, "digit value out of range: {value}");
            '?'
        }
    }
}

fn to_radix_upper(mut value: u64, radix: u32) -> String {
    debug_assert!((2..=36).contains(&radix));

    if value == 0 {
        return "0".to_string();
    }

    let mut buf = Vec::<char>::new();
    while value > 0 {
        let digit = (value % (radix as u64)) as u32;
        buf.push(digit_char(digit));
        value /= radix as u64;
    }
    buf.iter().rev().collect()
}
