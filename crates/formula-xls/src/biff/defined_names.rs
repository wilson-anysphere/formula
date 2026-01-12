//! BIFF8 defined name (`NAME` / `0x0018`) parsing.
//!
//! This module implements a small, best-effort parser for workbook- and sheet-scoped defined
//! names (named ranges / constants) stored in the workbook-global substream.
//!
//! The parser:
//! - extracts `NAME` records (including scope + hidden + description/comment)
//! - extracts the workbook `EXTERNSHEET` table (for 3D reference rendering)
//! - decodes BIFF8 `rgce` token streams into formula text (no leading `=`)

#![allow(dead_code)]

use super::{records, rgce, strings, BiffVersion};

// Record ids used by workbook-global defined name parsing.
// See [MS-XLS] sections:
// - EXTERNSHEET: 2.4.103
// - NAME: 2.4.150
const RECORD_EXTERNSHEET: u16 = 0x0017;
const RECORD_NAME: u16 = 0x0018;

// NAME record flags (Lbl.grbit).
// See [MS-XLS] 2.4.150 (NAME) / 2.5.114 (Lbl).
const NAME_FLAG_HIDDEN: u16 = 0x0001;
// fBuiltin (bit 5) indicates the name is a built-in defined name (e.g. print area).
const NAME_FLAG_BUILTIN: u16 = 0x0020;

// BIFF8 string option flags used by `XLUnicodeStringNoCch`.
// See [MS-XLS] 2.5.277.
const STR_FLAG_HIGH_BYTE: u8 = 0x01;
const STR_FLAG_EXT: u8 = 0x04;
const STR_FLAG_RICH_TEXT: u8 = 0x08;

#[derive(Debug, Clone)]
pub(crate) struct BiffDefinedName {
    pub(crate) name: String,
    /// BIFF sheet index (0-based) for local names, or `None` for workbook scope.
    pub(crate) scope_sheet: Option<usize>,
    pub(crate) refers_to: String,
    pub(crate) hidden: bool,
    pub(crate) comment: Option<String>,
}

#[derive(Debug, Clone, Default)]
pub(crate) struct BiffDefinedNames {
    pub(crate) names: Vec<BiffDefinedName>,
    pub(crate) warnings: Vec<String>,
}

#[derive(Debug, Clone)]
struct RawDefinedName {
    name: String,
    scope_sheet: Option<usize>,
    hidden: bool,
    comment: Option<String>,
    rgce: Vec<u8>,
}

pub(crate) fn parse_biff_defined_names(
    workbook_stream: &[u8],
    biff: BiffVersion,
    codepage: u16,
    sheet_names: &[String],
) -> Result<BiffDefinedNames, String> {
    let mut out = BiffDefinedNames::default();

    if biff != BiffVersion::Biff8 {
        out.warnings
            .push("BIFF defined name import currently supports BIFF8 only".to_string());
        return Ok(out);
    }

    let allows_continuation = |id: u16| id == RECORD_EXTERNSHEET || id == RECORD_NAME;
    let iter = records::LogicalBiffRecordIter::new(workbook_stream, allows_continuation);

    let mut externsheet: Vec<rgce::ExternSheetRef> = Vec::new();
    let mut raw_names: Vec<RawDefinedName> = Vec::new();

    for record in iter {
        let record = match record {
            Ok(record) => record,
            Err(err) => {
                out.warnings.push(format!("malformed BIFF record: {err}"));
                break;
            }
        };

        let record_id = record.record_id;

        // Stop at the next substream BOF; workbook globals start at offset 0.
        if record.offset != 0 && records::is_bof_record(record_id) {
            break;
        }

        match record_id {
            RECORD_EXTERNSHEET => match parse_externsheet_record(record.data.as_ref()) {
                Ok(table) => externsheet = table,
                Err(err) => out
                    .warnings
                    .push(format!("failed to parse EXTERNSHEET record: {err}")),
            },
            RECORD_NAME => match parse_biff8_name_record(&record, codepage, sheet_names) {
                Ok(raw) => raw_names.push(raw),
                Err(err) => out.warnings.push(format!("failed to parse NAME record: {err}")),
            },
            records::RECORD_EOF => break,
            _ => {}
        }
    }

    // Build name table metadata (for PtgName resolution), then decode formulas.
    let metas: Vec<rgce::DefinedNameMeta> = raw_names
        .iter()
        .map(|n| rgce::DefinedNameMeta {
            name: n.name.clone(),
            scope_sheet: n.scope_sheet,
        })
        .collect();

    let ctx = rgce::RgceDecodeContext {
        codepage,
        sheet_names,
        externsheet: &externsheet,
        defined_names: &metas,
    };

    for raw in raw_names {
        let decoded = rgce::decode_biff8_rgce(&raw.rgce, &ctx);
        for warning in decoded.warnings {
            out.warnings.push(format!("defined name `{}`: {warning}", raw.name));
        }

        out.names.push(BiffDefinedName {
            name: raw.name,
            scope_sheet: raw.scope_sheet,
            refers_to: decoded.text,
            hidden: raw.hidden,
            comment: raw.comment,
        });
    }

    Ok(out)
}

fn parse_externsheet_record(data: &[u8]) -> Result<Vec<rgce::ExternSheetRef>, String> {
    if data.len() < 2 {
        return Err("EXTERNSHEET record too short".to_string());
    }

    let count = u16::from_le_bytes([data[0], data[1]]) as usize;
    let mut offset = 2usize;
    let mut out = Vec::with_capacity(count);

    for _ in 0..count {
        if data.len() < offset + 6 {
            return Err("EXTERNSHEET record truncated".to_string());
        }
        // iSupBook is currently ignored; we only support internal sheet refs.
        let _isupbook = u16::from_le_bytes([data[offset], data[offset + 1]]);
        let itab_first = u16::from_le_bytes([data[offset + 2], data[offset + 3]]);
        let itab_last = u16::from_le_bytes([data[offset + 4], data[offset + 5]]);
        offset += 6;
        out.push(rgce::ExternSheetRef {
            itab_first,
            itab_last,
        });
    }

    Ok(out)
}

fn parse_biff8_name_record(
    record: &records::LogicalBiffRecord<'_>,
    codepage: u16,
    sheet_names: &[String],
) -> Result<RawDefinedName, String> {
    let fragments: Vec<&[u8]> = record.fragments().collect();
    let mut cursor = FragmentCursor::new(&fragments, 0, 0);

    // Fixed-size `NAME` record header (14 bytes).
    // [MS-XLS] 2.4.150
    let grbit = cursor.read_u16_le()?;
    let _ch_key = cursor.read_u8()?;
    let cch = cursor.read_u8()? as usize;
    let cce = cursor.read_u16_le()? as usize;
    let _ixals = cursor.read_u16_le()?;
    let itab_raw = cursor.read_u16_le()?;
    let cch_cust_menu = cursor.read_u8()? as usize;
    let cch_description = cursor.read_u8()? as usize;
    let cch_help_topic = cursor.read_u8()? as usize;
    let cch_status_text = cursor.read_u8()? as usize;

    let hidden = (grbit & NAME_FLAG_HIDDEN) != 0;
    let builtin = (grbit & NAME_FLAG_BUILTIN) != 0;

    let scope_sheet = if itab_raw == 0 {
        None
    } else {
        Some(itab_raw as usize - 1)
    };

    let name = if builtin {
        let id = cursor.read_u8()?;
        builtin_name_to_string(id)
    } else {
        cursor.read_biff8_unicode_string_no_cch(cch, codepage)?
    };

    // `rgce`: parsed formula bytes.
    //
    // BIFF8 can insert an additional option-flags byte at the start of a `CONTINUE` fragment when
    // an in-record string (e.g. a `PtgStr` token) is split across fragments. We therefore parse the
    // rgce token stream in a fragment-aware way so those continuation flag bytes are not treated as
    // rgce payload bytes.
    let rgce = cursor.read_biff8_rgce(cce)?;

    // Optional strings.
    if cch_cust_menu > 0 {
        let _ = cursor.read_biff8_unicode_string_no_cch(cch_cust_menu, codepage)?;
    }
    let comment = if cch_description > 0 {
        Some(cursor.read_biff8_unicode_string_no_cch(
            cch_description,
            codepage,
        )?)
    } else {
        None
    };
    if cch_help_topic > 0 {
        let _ = cursor.read_biff8_unicode_string_no_cch(cch_help_topic, codepage)?;
    }
    if cch_status_text > 0 {
        let _ = cursor.read_biff8_unicode_string_no_cch(cch_status_text, codepage)?;
    }

    if let Some(scope) = scope_sheet {
        if scope >= sheet_names.len() {
            log::warn!(
                "NAME record `{name}` has out-of-range itab={itab_raw} (sheet count={})",
                sheet_names.len()
            );
        }
    }

    Ok(RawDefinedName {
        name,
        scope_sheet,
        hidden,
        comment,
        rgce,
    })
}

fn builtin_name_to_string(id: u8) -> String {
    match id {
        0x06 => formula_model::XLNM_PRINT_AREA.to_string(),
        0x07 => formula_model::XLNM_PRINT_TITLES.to_string(),
        0x0D => formula_model::XLNM_FILTER_DATABASE.to_string(),
        other => {
            log::warn!("unknown BIFF built-in defined name id 0x{other:02X}");
            // Must be a valid `DefinedName` identifier (`validate_defined_name`), so keep it
            // alphanumeric + underscore only.
            format!("__biff_builtin_name_0x{other:02X}")
        }
    }
}

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

    fn read_biff8_unicode_string_no_cch(
        &mut self,
        cch: usize,
        codepage: u16,
    ) -> Result<String, String> {
        let flags = self.read_u8()?;

        let richtext_runs = if flags & STR_FLAG_RICH_TEXT != 0 {
            self.read_u16_le()? as usize
        } else {
            0
        };

        let ext_size = if flags & STR_FLAG_EXT != 0 {
            self.read_u32_le()? as usize
        } else {
            0
        };

        let mut is_unicode = (flags & STR_FLAG_HIGH_BYTE) != 0;
        let mut remaining_chars = cch;
        let mut out = String::new();

        while remaining_chars > 0 {
            if self.remaining_in_fragment() == 0 {
                // Continuing character bytes into a new CONTINUE fragment: first byte is option
                // flags for the continued segment (fHighByte).
                self.advance_fragment()?;
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

            if is_unicode {
                let mut u16s = Vec::with_capacity(take_chars);
                for chunk in bytes.chunks_exact(2) {
                    u16s.push(u16::from_le_bytes([chunk[0], chunk[1]]));
                }
                out.push_str(&String::from_utf16_lossy(&u16s));
            } else {
                out.push_str(&strings::decode_ansi(codepage, bytes));
            }

            remaining_chars -= take_chars;
        }

        let richtext_bytes = richtext_runs
            .checked_mul(4)
            .ok_or_else(|| "rich text run count overflow".to_string())?;
        self.skip_bytes(richtext_bytes + ext_size)?;

        Ok(out)
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
                // Binary operators.
                0x03..=0x11
                // Unary +/- and postfix/paren/missarg.
                | 0x12
                | 0x13
                | 0x14
                | 0x15
                | 0x16 => {}
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
                // PtgAttr (evaluation hints / jump tables).
                //
                // Payload: [grbit: u8][wAttr: u16] + optional jump table for tAttrChoose.
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
                // PtgErr (1 byte)
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
                // PtgName (6 bytes)
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
                // 3D references: PtgRef3d / PtgArea3d.
                0x3A | 0x5A | 0x7A => {
                    let bytes = self.read_bytes(6)?;
                    out.extend_from_slice(&bytes);
                }
                0x3B | 0x5B | 0x7B => {
                    let bytes = self.read_bytes(10)?;
                    out.extend_from_slice(&bytes);
                }
                // PtgMem* tokens: consume the nested rgce payload. These tokens have the form:
                //   [ptg][cce: u16][rgce: cce bytes]
                //
                // The nested rgce stream itself can contain continued strings, so we parse it via
                // `read_biff8_rgce` recursively.
                0x26 | 0x46 | 0x66 | 0x27 | 0x47 | 0x67 | 0x28 | 0x48 | 0x68 | 0x29 | 0x49
                | 0x69 | 0x2E | 0x4E | 0x6E => {
                    let inner_cce = self.read_u16_le()? as usize;
                    out.extend_from_slice(&(inner_cce as u16).to_le_bytes());
                    let inner = self.read_biff8_rgce(inner_cce)?;
                    out.extend_from_slice(&inner);
                }
                _ => {
                    // Unsupported token: copy the remaining bytes as-is to satisfy the `cce`
                    // contract and avoid dropping the defined name entirely.
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
    fn parses_defined_name_with_continued_rgce_bytes() {
        let name = "Name";

        // 1+2
        let rgce: Vec<u8> = vec![
            0x1E, 0x01, 0x00, // PtgInt 1
            0x1E, 0x02, 0x00, // PtgInt 2
            0x03, // PtgAdd
        ];

        let mut header = Vec::new();
        header.extend_from_slice(&0u16.to_le_bytes()); // grbit
        header.push(0); // chKey
        header.push(name.len() as u8); // cch
        header.extend_from_slice(&(rgce.len() as u16).to_le_bytes()); // cce
        header.extend_from_slice(&0u16.to_le_bytes()); // ixals
        header.extend_from_slice(&0u16.to_le_bytes()); // itab
        header.extend_from_slice(&[0, 0, 0, 0]); // cchCustMenu, cchDescription, cchHelpTopic, cchStatusText

        let mut name_str = Vec::new();
        name_str.push(0); // flags (compressed)
        name_str.extend_from_slice(name.as_bytes());

        let first_rgce = &rgce[..4];
        let second_rgce = &rgce[4..];

        let r_bof = record(records::RECORD_BOF_BIFF8, &[0u8; 16]);
        let r_name = record(RECORD_NAME, &[header.clone(), name_str.clone(), first_rgce.to_vec()].concat());
        let r_continue = record(records::RECORD_CONTINUE, second_rgce);
        let r_eof = record(records::RECORD_EOF, &[]);
        let stream = [r_bof, r_name, r_continue, r_eof].concat();

        let parsed =
            parse_biff_defined_names(&stream, BiffVersion::Biff8, 1252, &[]).expect("parse names");
        assert_eq!(parsed.names.len(), 1);
        assert_eq!(parsed.names[0].name, name);
        assert_eq!(parsed.names[0].refers_to, "1+2");
        assert!(parsed.warnings.is_empty(), "warnings={:?}", parsed.warnings);
    }

    #[test]
    fn parses_defined_name_with_continued_name_string() {
        let name = "ABCDE";
        let rgce: Vec<u8> = vec![0x1E, 0x2A, 0x00]; // PtgInt 42

        let mut header = Vec::new();
        header.extend_from_slice(&0u16.to_le_bytes()); // grbit
        header.push(0); // chKey
        header.push(name.len() as u8); // cch
        header.extend_from_slice(&(rgce.len() as u16).to_le_bytes()); // cce
        header.extend_from_slice(&0u16.to_le_bytes()); // ixals
        header.extend_from_slice(&0u16.to_le_bytes()); // itab
        header.extend_from_slice(&[0, 0, 0, 0]); // cchCustMenu, cchDescription, cchHelpTopic, cchStatusText

        // Split the name string across records after 2 characters.
        let mut first = Vec::new();
        first.extend_from_slice(&header);
        first.push(0); // string flags (compressed)
        first.extend_from_slice(&name.as_bytes()[..2]); // "AB"

        let mut second = Vec::new();
        second.push(0); // continued segment option flags (fHighByte=0)
        second.extend_from_slice(&name.as_bytes()[2..]); // "CDE"
        second.extend_from_slice(&rgce);

        let stream = [
            record(records::RECORD_BOF_BIFF8, &[0u8; 16]),
            record(RECORD_NAME, &first),
            record(records::RECORD_CONTINUE, &second),
            record(records::RECORD_EOF, &[]),
        ]
        .concat();

        let parsed =
            parse_biff_defined_names(&stream, BiffVersion::Biff8, 1252, &[]).expect("parse names");
        assert_eq!(parsed.names.len(), 1);
        assert_eq!(parsed.names[0].name, name);
        assert_eq!(parsed.names[0].refers_to, "42");
        assert!(parsed.warnings.is_empty(), "warnings={:?}", parsed.warnings);
    }

    #[test]
    fn parses_defined_name_with_continued_ptgstr_token() {
        let name = "StrName";
        let literal = "ABCDE";

        // rgce containing a single PtgStr token (string literal).
        let rgce: Vec<u8> = [
            vec![0x17, literal.len() as u8, 0u8], // PtgStr + cch + flags (compressed)
            literal.as_bytes().to_vec(),
        ]
        .concat();

        let mut header = Vec::new();
        header.extend_from_slice(&0u16.to_le_bytes()); // grbit
        header.push(0); // chKey
        header.push(name.len() as u8); // cch
        header.extend_from_slice(&(rgce.len() as u16).to_le_bytes()); // cce
        header.extend_from_slice(&0u16.to_le_bytes()); // ixals
        header.extend_from_slice(&0u16.to_le_bytes()); // itab
        header.extend_from_slice(&[0, 0, 0, 0]); // cchCustMenu, cchDescription, cchHelpTopic, cchStatusText

        let mut name_str = Vec::new();
        name_str.push(0); // flags (compressed)
        name_str.extend_from_slice(name.as_bytes());

        // Split the PtgStr character bytes across the CONTINUE boundary after "AB".
        let first_rgce = &rgce[..5]; // ptg + cch + flags + "AB"
        let second_chars = &literal.as_bytes()[2..]; // "CDE"

        let mut continue_payload = Vec::new();
        continue_payload.push(0); // continued segment option flags (fHighByte=0)
        continue_payload.extend_from_slice(second_chars);

        let stream = [
            record(records::RECORD_BOF_BIFF8, &[0u8; 16]),
            record(RECORD_NAME, &[header, name_str, first_rgce.to_vec()].concat()),
            record(records::RECORD_CONTINUE, &continue_payload),
            record(records::RECORD_EOF, &[]),
        ]
        .concat();

        let parsed =
            parse_biff_defined_names(&stream, BiffVersion::Biff8, 1252, &[]).expect("parse names");
        assert_eq!(parsed.names.len(), 1);
        assert_eq!(parsed.names[0].name, name);
        assert_eq!(parsed.names[0].refers_to, "\"ABCDE\"");
        assert!(parsed.warnings.is_empty(), "warnings={:?}", parsed.warnings);
    }

    #[test]
    fn parses_defined_name_with_refn_before_continued_ptgstr_token() {
        let name = "RefStrName";
        let literal = "ABCDE";

        // rgce for `A1&"ABCDE"`, using a relative reference token (PtgRefN).
        let rgce: Vec<u8> = [
            vec![0x2C, 0x00, 0x00, 0x00, 0x00], // PtgRefN row_off=0 col_off=0 => A1 (best-effort base)
            vec![0x17, literal.len() as u8, 0u8], // PtgStr + cch + flags (compressed)
            literal.as_bytes().to_vec(),
            vec![0x08], // PtgConcat
        ]
        .concat();

        let mut header = Vec::new();
        header.extend_from_slice(&0u16.to_le_bytes()); // grbit
        header.push(0); // chKey
        header.push(name.len() as u8); // cch
        header.extend_from_slice(&(rgce.len() as u16).to_le_bytes()); // cce
        header.extend_from_slice(&0u16.to_le_bytes()); // ixals
        header.extend_from_slice(&0u16.to_le_bytes()); // itab
        header.extend_from_slice(&[0, 0, 0, 0]); // cchCustMenu, cchDescription, cchHelpTopic, cchStatusText

        let mut name_str = Vec::new();
        name_str.push(0); // flags (compressed)
        name_str.extend_from_slice(name.as_bytes());

        // Split the PtgStr character bytes across the CONTINUE boundary after "AB".
        let first_rgce = &rgce[..10]; // PtgRefN (5) + PtgStr header (3) + "AB" (2)
        let second_chars = &literal.as_bytes()[2..]; // "CDE"

        let mut continue_payload = Vec::new();
        continue_payload.push(0); // continued segment option flags (fHighByte=0)
        continue_payload.extend_from_slice(second_chars);
        continue_payload.push(0x08); // PtgConcat

        let stream = [
            record(records::RECORD_BOF_BIFF8, &[0u8; 16]),
            record(RECORD_NAME, &[header, name_str, first_rgce.to_vec()].concat()),
            record(records::RECORD_CONTINUE, &continue_payload),
            record(records::RECORD_EOF, &[]),
        ]
        .concat();

        let parsed =
            parse_biff_defined_names(&stream, BiffVersion::Biff8, 1252, &[]).expect("parse names");
        assert_eq!(parsed.names.len(), 1);
        assert_eq!(parsed.names[0].name, name);
        assert_eq!(parsed.names[0].refers_to, "A1&\"ABCDE\"");
        assert!(parsed.warnings.is_empty(), "warnings={:?}", parsed.warnings);
    }

    #[test]
    fn parses_defined_name_with_continued_description_string() {
        let name = "DescName";
        // rgce for `1` (PtgInt 1).
        let rgce: Vec<u8> = vec![0x1E, 0x01, 0x00];

        let description = "ABCDE";

        let mut header = Vec::new();
        header.extend_from_slice(&0u16.to_le_bytes()); // grbit
        header.push(0); // chKey
        header.push(name.len() as u8); // cch
        header.extend_from_slice(&(rgce.len() as u16).to_le_bytes()); // cce
        header.extend_from_slice(&0u16.to_le_bytes()); // ixals
        header.extend_from_slice(&0u16.to_le_bytes()); // itab
        header.push(0); // cchCustMenu
        header.push(description.len() as u8); // cchDescription
        header.push(0); // cchHelpTopic
        header.push(0); // cchStatusText

        let mut name_str = Vec::new();
        name_str.push(0); // flags (compressed)
        name_str.extend_from_slice(name.as_bytes());

        // Description string (XLUnicodeStringNoCch) split across fragments after "AB".
        let mut desc_part1 = Vec::new();
        desc_part1.push(0); // flags (compressed)
        desc_part1.extend_from_slice(&description.as_bytes()[..2]); // "AB"

        let mut desc_part2 = Vec::new();
        desc_part2.push(0); // continued segment option flags (fHighByte=0)
        desc_part2.extend_from_slice(&description.as_bytes()[2..]); // "CDE"

        let stream = [
            record(records::RECORD_BOF_BIFF8, &[0u8; 16]),
            record(
                RECORD_NAME,
                &[header, name_str, rgce.clone(), desc_part1].concat(),
            ),
            record(records::RECORD_CONTINUE, &desc_part2),
            record(records::RECORD_EOF, &[]),
        ]
        .concat();

        let parsed =
            parse_biff_defined_names(&stream, BiffVersion::Biff8, 1252, &[]).expect("parse names");
        assert_eq!(parsed.names.len(), 1);
        assert_eq!(parsed.names[0].name, name);
        assert_eq!(parsed.names[0].refers_to, "1");
        assert_eq!(parsed.names[0].comment.as_deref(), Some(description));
        assert!(parsed.warnings.is_empty(), "warnings={:?}", parsed.warnings);
    }
}
