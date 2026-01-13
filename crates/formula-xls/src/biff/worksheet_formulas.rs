//! BIFF8 worksheet formula record parsing helpers.
//!
//! BIFF8 worksheet formulas are stored as `rgce` token streams inside `FORMULA`, `SHRFMLA`, and
//! `ARRAY` records. These records can be split across `CONTINUE` boundaries.
//!
//! When a `PtgStr` (ShortXLUnicodeString) payload is continued into a `CONTINUE` record, Excel
//! inserts an extra 1-byte "continued segment" option flags prefix at the fragment boundary.
//! Naively concatenating record payload bytes therefore corrupts the rgce stream (token alignment),
//! typically producing string literals containing an embedded NUL and leaving trailing bytes.
//!
//! This module implements a fragment-aware `rgce` reader that tokenizes the stream and skips those
//! continuation flag bytes so downstream formula decoding sees the canonical rgce bytes.

#![allow(dead_code)]

use super::records;

// Worksheet record ids (BIFF8).
// See [MS-XLS]:
// - FORMULA: 2.4.127 (0x0006)
// - ARRAY: 2.4.19 (0x0221)
// - SHRFMLA: 2.4.276 (0x04BC)
pub(crate) const RECORD_FORMULA: u16 = 0x0006;
pub(crate) const RECORD_ARRAY: u16 = 0x0221;
pub(crate) const RECORD_SHRFMLA: u16 = 0x04BC;

// BIFF8 string option flags used by ShortXLUnicodeString.
// See [MS-XLS] 2.5.293.
const STR_FLAG_HIGH_BYTE: u8 = 0x01;
const STR_FLAG_EXT: u8 = 0x04;
const STR_FLAG_RICH_TEXT: u8 = 0x08;

#[derive(Debug, Clone, PartialEq, Eq)]
pub(crate) struct ParsedFormulaRecord {
    pub(crate) row: u16,
    pub(crate) col: u16,
    pub(crate) xf: u16,
    pub(crate) rgce: Vec<u8>,
}

#[derive(Debug, Clone, PartialEq, Eq)]
pub(crate) struct ParsedSharedFormulaRecord {
    pub(crate) rgce: Vec<u8>,
}

#[derive(Debug, Clone, PartialEq, Eq)]
pub(crate) struct ParsedArrayRecord {
    pub(crate) rgce: Vec<u8>,
}

pub(crate) fn parse_biff8_formula_record(
    record: &records::LogicalBiffRecord<'_>,
) -> Result<ParsedFormulaRecord, String> {
    let fragments: Vec<&[u8]> = record.fragments().collect();
    let mut cursor = FragmentCursor::new(&fragments, 0, 0);

    // FORMULA [MS-XLS 2.4.127]
    let row = cursor.read_u16_le()?;
    let col = cursor.read_u16_le()?;
    let xf = cursor.read_u16_le()?;

    // Skip cached result (8), flags (2), and calc chain (4).
    cursor.skip_bytes(8 + 2 + 4)?;

    let cce = cursor.read_u16_le()? as usize;
    let rgce = cursor.read_biff8_rgce(cce)?;

    Ok(ParsedFormulaRecord { row, col, xf, rgce })
}

pub(crate) fn parse_biff8_shrfmla_record(
    record: &records::LogicalBiffRecord<'_>,
) -> Result<ParsedSharedFormulaRecord, String> {
    let fragments: Vec<&[u8]> = record.fragments().collect();
    let cursor = FragmentCursor::new(&fragments, 0, 0);

    // SHRFMLA layouts vary slightly between producers (RefU vs Ref8 for the shared range). Try a
    // small set of plausible BIFF8 layouts.
    // Layout A: RefU (6) + cUse (2) + cce (2).
    let mut c = cursor.clone();
    if let Ok(rgce) = parse_shrfmla_with_refu(&mut c) {
        return Ok(ParsedSharedFormulaRecord { rgce });
    }
    // Layout B: Ref8 (8) + cUse (2) + cce (2).
    let mut c = cursor;
    if let Ok(rgce) = parse_shrfmla_with_ref8(&mut c) {
        return Ok(ParsedSharedFormulaRecord { rgce });
    }

    Err("unrecognized SHRFMLA record layout".to_string())
}

pub(crate) fn parse_biff8_array_record(
    record: &records::LogicalBiffRecord<'_>,
) -> Result<ParsedArrayRecord, String> {
    let fragments: Vec<&[u8]> = record.fragments().collect();
    let cursor = FragmentCursor::new(&fragments, 0, 0);

    // ARRAY layouts vary slightly (RefU vs Ref8). Try both.
    {
        let mut c = cursor.clone();
        if let Ok(rgce) = parse_array_with_refu(&mut c) {
            return Ok(ParsedArrayRecord { rgce });
        }
    }
    {
        let mut c = cursor;
        if let Ok(rgce) = parse_array_with_ref8(&mut c) {
            return Ok(ParsedArrayRecord { rgce });
        }
    }

    Err("unrecognized ARRAY record layout".to_string())
}

fn parse_shrfmla_with_refu(cursor: &mut FragmentCursor<'_>) -> Result<Vec<u8>, String> {
    // ref (rwFirst:u16, rwLast:u16, colFirst:u8, colLast:u8)
    cursor.skip_bytes(2 + 2 + 1 + 1)?;
    // cUse
    cursor.skip_bytes(2)?;
    let cce = cursor.read_u16_le()? as usize;
    cursor.read_biff8_rgce(cce)
}

fn parse_shrfmla_with_ref8(cursor: &mut FragmentCursor<'_>) -> Result<Vec<u8>, String> {
    // ref (rwFirst:u16, rwLast:u16, colFirst:u16, colLast:u16)
    cursor.skip_bytes(8)?;
    // cUse
    cursor.skip_bytes(2)?;
    let cce = cursor.read_u16_le()? as usize;
    cursor.read_biff8_rgce(cce)
}

fn parse_array_with_refu(cursor: &mut FragmentCursor<'_>) -> Result<Vec<u8>, String> {
    // ref (rwFirst:u16, rwLast:u16, colFirst:u8, colLast:u8)
    cursor.skip_bytes(2 + 2 + 1 + 1)?;
    // reserved
    cursor.skip_bytes(2)?;
    let cce = cursor.read_u16_le()? as usize;
    cursor.read_biff8_rgce(cce)
}

fn parse_array_with_ref8(cursor: &mut FragmentCursor<'_>) -> Result<Vec<u8>, String> {
    // ref (rwFirst:u16, rwLast:u16, colFirst:u16, colLast:u16)
    cursor.skip_bytes(8)?;
    // reserved
    cursor.skip_bytes(2)?;
    let cce = cursor.read_u16_le()? as usize;
    cursor.read_biff8_rgce(cce)
}

#[derive(Debug, Clone)]
struct FragmentCursor<'a> {
    fragments: &'a [&'a [u8]],
    frag_idx: usize,
    offset: usize,
}

impl<'a> FragmentCursor<'a> {
    fn new(fragments: &'a [&'a [u8]], frag_idx: usize, offset: usize) -> Self {
        Self {
            fragments,
            frag_idx,
            offset,
        }
    }

    fn remaining_in_fragment(&self) -> usize {
        self.fragments
            .get(self.frag_idx)
            .map(|f| f.len().saturating_sub(self.offset))
            .unwrap_or(0)
    }

    fn advance_fragment(&mut self) -> Result<(), String> {
        self.frag_idx = self
            .frag_idx
            .checked_add(1)
            .ok_or_else(|| "fragment index overflow".to_string())?;
        self.offset = 0;
        if self.frag_idx >= self.fragments.len() {
            return Err("unexpected end of record".to_string());
        }
        Ok(())
    }

    fn read_u8(&mut self) -> Result<u8, String> {
        loop {
            let frag = self
                .fragments
                .get(self.frag_idx)
                .ok_or_else(|| "unexpected end of record".to_string())?;
            if self.offset < frag.len() {
                let b = frag[self.offset];
                self.offset += 1;
                return Ok(b);
            }
            self.advance_fragment()?;
        }
    }

    fn read_u16_le(&mut self) -> Result<u16, String> {
        let lo = self.read_u8()?;
        let hi = self.read_u8()?;
        Ok(u16::from_le_bytes([lo, hi]))
    }

    fn read_u32_le(&mut self) -> Result<u32, String> {
        let b0 = self.read_u8()?;
        let b1 = self.read_u8()?;
        let b2 = self.read_u8()?;
        let b3 = self.read_u8()?;
        Ok(u32::from_le_bytes([b0, b1, b2, b3]))
    }

    fn read_exact_from_current(&mut self, n: usize) -> Result<&'a [u8], String> {
        let frag = self
            .fragments
            .get(self.frag_idx)
            .ok_or_else(|| "unexpected end of record".to_string())?;
        let end = self
            .offset
            .checked_add(n)
            .ok_or_else(|| "offset overflow".to_string())?;
        if end > frag.len() {
            return Err("unexpected end of record".to_string());
        }
        let out = &frag[self.offset..end];
        self.offset = end;
        Ok(out)
    }

    fn read_bytes(&mut self, mut n: usize) -> Result<Vec<u8>, String> {
        let mut out = Vec::with_capacity(n);
        while n > 0 {
            let available = self.remaining_in_fragment();
            if available == 0 {
                self.advance_fragment()?;
                continue;
            }
            let take = n.min(available);
            let bytes = self.read_exact_from_current(take)?;
            out.extend_from_slice(bytes);
            n -= take;
        }
        Ok(out)
    }

    fn skip_bytes(&mut self, mut n: usize) -> Result<(), String> {
        while n > 0 {
            let available = self.remaining_in_fragment();
            if available == 0 {
                self.advance_fragment()?;
                continue;
            }
            let take = n.min(available);
            self.offset += take;
            n -= take;
        }
        Ok(())
    }

    fn read_biff8_rgce(&mut self, cce: usize) -> Result<Vec<u8>, String> {
        // Best-effort: parse BIFF8 ptg tokens so we can skip the continuation flags byte injected
        // at fragment boundaries when a `PtgStr` (ShortXLUnicodeString) payload is split across
        // `CONTINUE` records.
        //
        // If we encounter an unsupported token, fall back to raw byte copying for the remainder of
        // the `rgce` stream (without special continuation handling).
        let mut out = Vec::with_capacity(cce);

        while out.len() < cce {
            let ptg = self.read_u8()?;
            out.push(ptg);

            match ptg {
                // PtgExp / PtgTbl: shared/array formula tokens.
                0x01 | 0x02 => {
                    let bytes = self.read_bytes(4)?;
                    out.extend_from_slice(&bytes);
                }
                // Binary operators.
                0x03..=0x11
                // Unary +/- and postfix/paren/missarg.
                | 0x12
                | 0x13
                | 0x14
                | 0x15
                | 0x16 => {}
                // Spill range postfix (`#`).
                0x2F => {}
                // PtgStr (ShortXLUnicodeString) [MS-XLS 2.5.293]
                0x17 => {
                    let cch = self.read_u8()? as usize;
                    let flags = self.read_u8()?;
                    out.push(cch as u8);
                    out.push(flags);

                    let richtext_runs = if (flags & STR_FLAG_RICH_TEXT) != 0 {
                        let v = self.read_u16_le()?;
                        out.extend_from_slice(&v.to_le_bytes());
                        v as usize
                    } else {
                        0
                    };

                    let ext_size = if (flags & STR_FLAG_EXT) != 0 {
                        let v = self.read_u32_le()?;
                        out.extend_from_slice(&v.to_le_bytes());
                        v as usize
                    } else {
                        0
                    };

                    let mut is_unicode = (flags & STR_FLAG_HIGH_BYTE) != 0;
                    let mut remaining_chars = cch;

                    while remaining_chars > 0 {
                        if self.remaining_in_fragment() == 0 {
                            self.advance_fragment()?;
                            // Continued-segment option flags byte (fHighByte).
                            let cont_flags = self.read_u8()?;
                            is_unicode = (cont_flags & STR_FLAG_HIGH_BYTE) != 0;
                            continue;
                        }

                        let bytes_per_char = if is_unicode { 2 } else { 1 };
                        let available_bytes = self.remaining_in_fragment();
                        let available_chars = available_bytes / bytes_per_char;
                        if available_chars == 0 {
                            return Err("string continuation split mid-character".to_string());
                        }

                        let take_chars = remaining_chars.min(available_chars);
                        let take_bytes = take_chars * bytes_per_char;
                        let bytes = self.read_exact_from_current(take_bytes)?;
                        out.extend_from_slice(bytes);
                        remaining_chars -= take_chars;
                    }

                    let richtext_bytes = richtext_runs
                        .checked_mul(4)
                        .ok_or_else(|| "rich text run count overflow".to_string())?;
                    if richtext_bytes + ext_size > 0 {
                        let extra = self.read_bytes(richtext_bytes + ext_size)?;
                        out.extend_from_slice(&extra);
                    }
                }
                // PtgExtend* token 0x18 (and class variants).
                0x18 | 0x38 | 0x58 | 0x78 => {
                    let etpg = self.read_u8()?;
                    out.push(etpg);
                    if etpg == 0x19 {
                        let bytes = self.read_bytes(12)?;
                        out.extend_from_slice(&bytes);
                    } else {
                        let bytes = self.read_bytes(4)?;
                        out.extend_from_slice(&bytes);
                    }
                }
                // PtgAttr (evaluation hints / jump tables).
                0x19 => {
                    let grbit = self.read_u8()?;
                    let w_attr = self.read_u16_le()?;
                    out.push(grbit);
                    out.extend_from_slice(&w_attr.to_le_bytes());

                    // tAttrChoose includes a jump table of `u16` offsets (wAttr entries).
                    const T_ATTR_CHOOSE: u8 = 0x04;
                    if (grbit & T_ATTR_CHOOSE) != 0 {
                        let entries = w_attr as usize;
                        let bytes = entries
                            .checked_mul(2)
                            .ok_or_else(|| "tAttrChoose jump table length overflow".to_string())?;
                        let table = self.read_bytes(bytes)?;
                        out.extend_from_slice(&table);
                    }
                }
                // PtgErr / PtgBool (1 byte)
                0x1C | 0x1D => {
                    out.push(self.read_u8()?);
                }
                // PtgInt (2 bytes)
                0x1E => {
                    let bytes = self.read_bytes(2)?;
                    out.extend_from_slice(&bytes);
                }
                // PtgNum (8 bytes)
                0x1F => {
                    let bytes = self.read_bytes(8)?;
                    out.extend_from_slice(&bytes);
                }
                // PtgArray (7 bytes) [MS-XLS 2.5.198.8]
                0x20 | 0x40 | 0x60 => {
                    let bytes = self.read_bytes(7)?;
                    out.extend_from_slice(&bytes);
                }
                // PtgFunc (2 bytes)
                0x21 | 0x41 | 0x61 => {
                    let bytes = self.read_bytes(2)?;
                    out.extend_from_slice(&bytes);
                }
                // PtgFuncVar (3 bytes)
                0x22 | 0x42 | 0x62 => {
                    let bytes = self.read_bytes(3)?;
                    out.extend_from_slice(&bytes);
                }
                // PtgName (defined name reference) (6 bytes).
                0x23 | 0x43 | 0x63 => {
                    let bytes = self.read_bytes(6)?;
                    out.extend_from_slice(&bytes);
                }
                // PtgRef (4 bytes)
                0x24 | 0x44 | 0x64 => {
                    let bytes = self.read_bytes(4)?;
                    out.extend_from_slice(&bytes);
                }
                // PtgArea (8 bytes)
                0x25 | 0x45 | 0x65 => {
                    let bytes = self.read_bytes(8)?;
                    out.extend_from_slice(&bytes);
                }
                // PtgRefErr (4 bytes)
                0x2A | 0x4A | 0x6A => {
                    let bytes = self.read_bytes(4)?;
                    out.extend_from_slice(&bytes);
                }
                // PtgAreaErr (8 bytes)
                0x2B | 0x4B | 0x6B => {
                    let bytes = self.read_bytes(8)?;
                    out.extend_from_slice(&bytes);
                }
                // PtgRefN (4 bytes)
                0x2C | 0x4C | 0x6C => {
                    let bytes = self.read_bytes(4)?;
                    out.extend_from_slice(&bytes);
                }
                // PtgAreaN (8 bytes)
                0x2D | 0x4D | 0x6D => {
                    let bytes = self.read_bytes(8)?;
                    out.extend_from_slice(&bytes);
                }
                // PtgNameX (external name) [MS-XLS 2.5.198.41]
                0x39 | 0x59 | 0x79 => {
                    let bytes = self.read_bytes(6)?;
                    out.extend_from_slice(&bytes);
                }
                // 3D references: PtgRef3d / PtgArea3d.
                0x3A | 0x5A | 0x7A => {
                    let bytes = self.read_bytes(6)?;
                    out.extend_from_slice(&bytes);
                }
                0x3B | 0x5B | 0x7B => {
                    let bytes = self.read_bytes(10)?;
                    out.extend_from_slice(&bytes);
                }
                // 3D error references: PtgRefErr3d / PtgAreaErr3d.
                0x3C | 0x5C | 0x7C => {
                    let bytes = self.read_bytes(6)?;
                    out.extend_from_slice(&bytes);
                }
                0x3D | 0x5D | 0x7D => {
                    let bytes = self.read_bytes(10)?;
                    out.extend_from_slice(&bytes);
                }
                // 3D relative references: PtgRefN3d / PtgAreaN3d.
                0x3E | 0x5E | 0x7E => {
                    let bytes = self.read_bytes(6)?;
                    out.extend_from_slice(&bytes);
                }
                0x3F | 0x5F | 0x7F => {
                    let bytes = self.read_bytes(10)?;
                    out.extend_from_slice(&bytes);
                }
                // PtgMem* tokens: [ptg][cce: u16][rgce: cce bytes]
                0x26 | 0x46 | 0x66 | 0x27 | 0x47 | 0x67 | 0x28 | 0x48 | 0x68 | 0x29 | 0x49
                | 0x69 | 0x2E | 0x4E | 0x6E => {
                    let inner_cce = self.read_u16_le()? as usize;
                    out.extend_from_slice(&(inner_cce as u16).to_le_bytes());
                    let inner = self.read_biff8_rgce(inner_cce)?;
                    out.extend_from_slice(&inner);
                }
                _ => {
                    // Unsupported token: copy the remaining bytes as-is to satisfy the `cce`
                    // contract and avoid dropping the formula entirely.
                    let remaining = cce.saturating_sub(out.len());
                    if remaining > 0 {
                        let bytes = self.read_bytes(remaining)?;
                        out.extend_from_slice(&bytes);
                    }
                }
            }
        }

        if out.len() != cce {
            return Err(format!(
                "rgce length mismatch (expected {cce} bytes, got {})",
                out.len()
            ));
        }

        Ok(out)
    }
}

#[cfg(test)]
mod tests {
    use super::*;

    fn record(id: u16, payload: &[u8]) -> Vec<u8> {
        let mut out = Vec::with_capacity(4 + payload.len());
        out.extend_from_slice(&id.to_le_bytes());
        out.extend_from_slice(&(payload.len() as u16).to_le_bytes());
        out.extend_from_slice(payload);
        out
    }

    #[test]
    fn parses_formula_rgce_with_continued_ptgstr_token() {
        // Build a FORMULA record whose rgce contains a PtgStr token split across a CONTINUE
        // boundary. Excel inserts a 1-byte "continued segment" option flags prefix at the start of
        // the continued fragment; ensure we skip it so the recovered rgce matches the canonical
        // stream.
        let literal = "ABCDE";

        let rgce_expected: Vec<u8> = [
            vec![0x17, literal.len() as u8, 0u8], // PtgStr + cch + flags (compressed)
            literal.as_bytes().to_vec(),
        ]
        .concat();

        // Split after the first two characters ("AB"). The continued fragment begins with the
        // continued-segment option flags byte (fHighByte), then the remaining bytes.
        let first_rgce = &rgce_expected[..(3 + 2)]; // ptg + cch + flags + "AB"
        let remaining_chars = &literal.as_bytes()[2..]; // "CDE"
        let mut continue_payload = Vec::new();
        continue_payload.push(0); // continued segment option flags (compressed)
        continue_payload.extend_from_slice(remaining_chars);

        // Minimal BIFF8 FORMULA record header (matches `xls_fixture_builder::formula_cell`):
        // [row][col][xf][cached_result:f64][grbit][chn][cce][rgce]
        let row = 1u16;
        let col = 2u16;
        let xf = 3u16;
        let cached_result = 0f64;
        let cce = rgce_expected.len() as u16;

        let mut formula_payload_part1 = Vec::new();
        formula_payload_part1.extend_from_slice(&row.to_le_bytes());
        formula_payload_part1.extend_from_slice(&col.to_le_bytes());
        formula_payload_part1.extend_from_slice(&xf.to_le_bytes());
        formula_payload_part1.extend_from_slice(&cached_result.to_le_bytes());
        formula_payload_part1.extend_from_slice(&0u16.to_le_bytes()); // grbit
        formula_payload_part1.extend_from_slice(&0u32.to_le_bytes()); // chn
        formula_payload_part1.extend_from_slice(&cce.to_le_bytes());
        formula_payload_part1.extend_from_slice(first_rgce);

        let stream = [
            record(RECORD_FORMULA, &formula_payload_part1),
            record(records::RECORD_CONTINUE, &continue_payload),
        ]
        .concat();

        let allows_continuation = |id: u16| id == RECORD_FORMULA;
        let mut iter = records::LogicalBiffRecordIter::new(&stream, allows_continuation);
        let record = iter.next().expect("record").expect("logical record");
        assert_eq!(record.record_id, RECORD_FORMULA);
        assert!(record.is_continued());

        let parsed = parse_biff8_formula_record(&record).expect("parse formula");
        assert_eq!(parsed.row, row);
        assert_eq!(parsed.col, col);
        assert_eq!(parsed.xf, xf);
        assert_eq!(parsed.rgce, rgce_expected);
    }
}
