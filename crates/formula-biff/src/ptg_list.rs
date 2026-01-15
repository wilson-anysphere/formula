/// Best-effort decoding helpers for BIFF12 `PtgList` payloads (structured references / tables).
///
/// MS-XLSB documents a fixed 12-byte payload for `etpg=0x19` (`PtgList`), but there are multiple
/// observed field layouts "in the wild". This module provides a shared, allocation-free way to
/// interpret a payload under several plausible layouts so higher-level decoders (XLS, XLSB, and
/// BIFF12 formula decoding) can apply their own scoring heuristics.

#[derive(Clone, Copy, Debug, PartialEq, Eq)]
pub struct PtgListDecoded {
    pub table_id: u32,
    pub flags: u32,
    pub col_first: u32,
    pub col_last: u32,
}

/// Decode a 12-byte `PtgList` payload under several plausible layouts.
///
/// Returned candidates are ordered by likelihood in well-formed files:
/// - Layout A (canonical): `[table_id: u32][flags: u16][col_first: u16][col_last: u16][reserved: u16]`
/// - Layout B: `[table_id: u32][col_first_raw: u32][col_last_raw: u32]` with flags packed into the high u16
/// - Layout C: `[table_id: u32][flags: u32][col_spec: u32]` with columns packed into `col_spec`
/// - Layout D: treat the middle/end u32s as raw column ids with no flags
pub fn decode_ptg_list_payload_candidates(payload: &[u8; 12]) -> [PtgListDecoded; 4] {
    let table_id = u32::from_le_bytes([payload[0], payload[1], payload[2], payload[3]]);

    let flags_a = u16::from_le_bytes([payload[4], payload[5]]) as u32;
    let col_first_a = u16::from_le_bytes([payload[6], payload[7]]) as u32;
    let col_last_a = u16::from_le_bytes([payload[8], payload[9]]) as u32;

    let col_first_raw = u32::from_le_bytes([payload[4], payload[5], payload[6], payload[7]]);
    let col_last_raw = u32::from_le_bytes([payload[8], payload[9], payload[10], payload[11]]);
    let col_first_b = (col_first_raw & 0xFFFF) as u32;
    let flags_b = (col_first_raw >> 16) & 0xFFFF;
    let col_last_b = (col_last_raw & 0xFFFF) as u32;

    let flags_c = u32::from_le_bytes([payload[4], payload[5], payload[6], payload[7]]);
    let col_spec_c = u32::from_le_bytes([payload[8], payload[9], payload[10], payload[11]]);
    let col_first_c = (col_spec_c & 0xFFFF) as u32;
    let col_last_c = ((col_spec_c >> 16) & 0xFFFF) as u32;

    [
        PtgListDecoded {
            table_id,
            flags: flags_a,
            col_first: col_first_a,
            col_last: col_last_a,
        },
        PtgListDecoded {
            table_id,
            flags: flags_b,
            col_first: col_first_b,
            col_last: col_last_b,
        },
        PtgListDecoded {
            table_id,
            flags: flags_c,
            col_first: col_first_c,
            col_last: col_last_c,
        },
        PtgListDecoded {
            table_id,
            flags: 0,
            col_first: col_first_raw,
            col_last: col_last_raw,
        },
    ]
}

