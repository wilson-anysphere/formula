use crate::value::ErrorKind;

/// Excel's BIT* functions operate on 48-bit unsigned integers (0..2^48-1).
pub(crate) const BIT_MAX: u64 = (1u64 << 48) - 1;

pub(crate) fn bitand(a: u64, b: u64) -> u64 {
    a & b
}

pub(crate) fn bitor(a: u64, b: u64) -> u64 {
    a | b
}

pub(crate) fn bitxor(a: u64, b: u64) -> u64 {
    a ^ b
}

pub(crate) fn bitlshift(value: u64, shift: i32) -> Result<u64, ErrorKind> {
    bit_shift(value, shift, ShiftDirection::Left)
}

pub(crate) fn bitrshift(value: u64, shift: i32) -> Result<u64, ErrorKind> {
    bit_shift(value, shift, ShiftDirection::Right)
}

#[derive(Debug, Clone, Copy, PartialEq, Eq)]
enum ShiftDirection {
    Left,
    Right,
}

fn bit_shift(value: u64, shift: i32, dir: ShiftDirection) -> Result<u64, ErrorKind> {
    // Excel documentation states shift_amount can be between -53 and 53.
    if shift < -53 || shift > 53 {
        return Err(ErrorKind::Num);
    }

    let (shift_dir, amount) = if shift >= 0 {
        (dir, shift as u32)
    } else {
        (
            match dir {
                ShiftDirection::Left => ShiftDirection::Right,
                ShiftDirection::Right => ShiftDirection::Left,
            },
            (-shift) as u32,
        )
    };

    let out = match shift_dir {
        ShiftDirection::Left => {
            let shifted = (value as u128).checked_shl(amount).ok_or(ErrorKind::Num)?;
            if shifted > (BIT_MAX as u128) {
                return Err(ErrorKind::Num);
            }
            shifted as u64
        }
        ShiftDirection::Right => value.checked_shr(amount).unwrap_or(0),
    };

    Ok(out)
}
