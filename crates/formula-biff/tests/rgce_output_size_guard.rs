use formula_biff::{decode_rgce, DecodeRgceError};

fn ptg_str_ascii(byte: u8, len: u16) -> Vec<u8> {
    // BIFF12 PtgStr:
    // [0x17][cch: u16][utf16 units...]
    let len_usize = len as usize;
    let mut out = Vec::with_capacity(1 + 2 + (len_usize * 2));
    out.push(0x17);
    out.extend_from_slice(&len.to_le_bytes());
    out.resize(3 + (len_usize * 2), 0);
    for i in 0..len_usize {
        out[3 + (i * 2)] = byte;
        out[3 + (i * 2) + 1] = 0;
    }
    out
}

#[test]
fn decode_rgce_rejects_pathological_output_size() {
    // Build a pathological rgce stream that would decode to a very large formula string.
    //
    // We intentionally avoid `encode_rgce` here so we can generate an oversized token stream
    // efficiently.
    const STR_LEN: u16 = u16::MAX; // 65535 UTF-16 units
    const STR_COUNT: usize = 16;

    let ptg_str = ptg_str_ascii(b'A', STR_LEN);

    let mut rgce = Vec::with_capacity(ptg_str.len() * STR_COUNT + (STR_COUNT - 1));
    for _ in 0..STR_COUNT {
        rgce.extend_from_slice(&ptg_str);
    }
    // Concatenate all strings: produce a very large formula text like:
    //   "AAAA...A"&"AAAA...A"&...
    rgce.extend(std::iter::repeat(0x08).take(STR_COUNT - 1)); // PtgConcat

    let err = decode_rgce(&rgce).expect_err("expected OutputTooLarge");
    match err {
        DecodeRgceError::OutputTooLarge {
            offset,
            ptg,
            max_len,
        } => {
            assert!(offset < rgce.len(), "offset should be within rgce");
            assert_eq!(ptg, 0x08, "expected to fail on PtgConcat");
            assert_eq!(max_len, rgce.len().saturating_mul(10).min(1_000_000));
        }
        other => panic!("expected OutputTooLarge, got: {other:?}"),
    }
}

#[test]
fn decode_rgce_still_decodes_small_formulas() {
    // 1 + 1 (RPN: 1, 1, +)
    let rgce = [0x1E, 0x01, 0x00, 0x1E, 0x01, 0x00, 0x03];
    let text = decode_rgce(&rgce).expect("decode");
    assert_eq!(text, "1+1");
}

