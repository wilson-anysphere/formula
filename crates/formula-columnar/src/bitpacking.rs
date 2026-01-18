#![forbid(unsafe_code)]

/// Returns the number of bits required to represent `value`.
pub fn bit_width_u64(value: u64) -> u8 {
    if value == 0 {
        0
    } else {
        (64 - value.leading_zeros()) as u8
    }
}

pub fn bit_width_u32(value: u32) -> u8 {
    if value == 0 {
        0
    } else {
        (32 - value.leading_zeros()) as u8
    }
}

/// Packs `values` into a byte vector using `bit_width` bits per value.
///
/// Bits are packed little-endian (lowest bits first).
pub fn pack_u64(values: &[u64], bit_width: u8) -> Vec<u8> {
    if bit_width == 0 || values.is_empty() {
        return Vec::new();
    }

    let total_bits = values.len().saturating_mul(bit_width as usize);
    let mut out = Vec::new();
    let _ = out.try_reserve_exact((total_bits + 7) / 8);

    let mut acc: u128 = 0;
    let mut acc_bits: u32 = 0;
    let mask: u128 = (1u128 << bit_width) - 1;

    for &v in values {
        acc |= ((v as u128) & mask) << acc_bits;
        acc_bits += bit_width as u32;

        while acc_bits >= 8 {
            out.push((acc & 0xFF) as u8);
            acc >>= 8;
            acc_bits -= 8;
        }
    }

    if acc_bits > 0 {
        out.push((acc & 0xFF) as u8);
    }

    out
}

pub fn pack_u32(values: &[u32], bit_width: u8) -> Vec<u8> {
    if bit_width == 0 || values.is_empty() {
        return Vec::new();
    }

    let total_bits = values.len().saturating_mul(bit_width as usize);
    let mut out = Vec::new();
    let _ = out.try_reserve_exact((total_bits + 7) / 8);

    let mut acc: u128 = 0;
    let mut acc_bits: u32 = 0;
    let mask: u128 = (1u128 << bit_width) - 1;

    for &v in values {
        acc |= ((v as u128) & mask) << acc_bits;
        acc_bits += bit_width as u32;

        while acc_bits >= 8 {
            out.push((acc & 0xFF) as u8);
            acc >>= 8;
            acc_bits -= 8;
        }
    }

    if acc_bits > 0 {
        out.push((acc & 0xFF) as u8);
    }

    out
}

pub fn unpack_u64(data: &[u8], bit_width: u8, count: usize) -> Vec<u64> {
    if count == 0 {
        return Vec::new();
    }
    if bit_width == 0 {
        return vec![0; count];
    }

    let mask: u128 = (1u128 << bit_width) - 1;
    let mut out = Vec::new();
    let _ = out.try_reserve_exact(count);

    let mut acc: u128 = 0;
    let mut acc_bits: u32 = 0;
    let mut byte_idx: usize = 0;

    for _ in 0..count {
        while acc_bits < bit_width as u32 {
            let b = data.get(byte_idx).copied().unwrap_or(0) as u128;
            acc |= b << acc_bits;
            acc_bits += 8;
            byte_idx += 1;
        }

        let v = (acc & mask) as u64;
        out.push(v);
        acc >>= bit_width;
        acc_bits -= bit_width as u32;
    }

    out
}

pub fn unpack_u32(data: &[u8], bit_width: u8, count: usize) -> Vec<u32> {
    if count == 0 {
        return Vec::new();
    }
    if bit_width == 0 {
        return vec![0; count];
    }

    let mask: u128 = (1u128 << bit_width) - 1;
    let mut out = Vec::new();
    let _ = out.try_reserve_exact(count);

    let mut acc: u128 = 0;
    let mut acc_bits: u32 = 0;
    let mut byte_idx: usize = 0;

    for _ in 0..count {
        while acc_bits < bit_width as u32 {
            let b = data.get(byte_idx).copied().unwrap_or(0) as u128;
            acc |= b << acc_bits;
            acc_bits += 8;
            byte_idx += 1;
        }

        let v = (acc & mask) as u32;
        out.push(v);
        acc >>= bit_width;
        acc_bits -= bit_width as u32;
    }

    out
}

pub fn get_u64_at(data: &[u8], bit_width: u8, index: usize) -> u64 {
    if bit_width == 0 {
        return 0;
    }

    let start_bit = index.saturating_mul(bit_width as usize);
    let byte_start = start_bit / 8;
    let bit_start = start_bit % 8;

    let needed_bits = bit_start + bit_width as usize;
    let needed_bytes = (needed_bits + 7) / 8;

    let mut acc: u128 = 0;
    for i in 0..needed_bytes {
        let b = data.get(byte_start + i).copied().unwrap_or(0) as u128;
        acc |= b << (8 * i);
    }

    acc >>= bit_start;
    let mask: u128 = (1u128 << bit_width) - 1;
    (acc & mask) as u64
}

#[cfg(test)]
mod tests {
    use super::*;

    #[test]
    fn pack_unpack_roundtrip() {
        let values: Vec<u64> = (0..1_000).map(|v| v * 3).collect();
        let max = *values.iter().max().unwrap();
        let bw = bit_width_u64(max);
        let packed = pack_u64(&values, bw);
        let decoded = unpack_u64(&packed, bw, values.len());
        assert_eq!(decoded, values);
    }

    #[test]
    fn get_u64_at_matches_unpack() {
        let values: Vec<u64> = (0..257).map(|v| (v * 13) as u64).collect();
        let max = *values.iter().max().unwrap();
        let bw = bit_width_u64(max);
        let packed = pack_u64(&values, bw);

        for (idx, expected) in values.iter().enumerate() {
            let got = get_u64_at(&packed, bw, idx);
            assert_eq!(got, *expected);
        }
    }
}
