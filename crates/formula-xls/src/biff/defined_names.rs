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

use super::{records, rgce, strings, workbook_context, BiffVersion};

// Record ids used by workbook-global defined name parsing.
// See [MS-XLS] sections:
// - NAME: 2.4.150
const RECORD_NAME: u16 = 0x0018;

// Avoid unbounded warning growth when parsing corrupt/hostile workbook-global NAME record streams.
const MAX_DEFINED_NAME_WARNINGS: usize = 200;

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
    /// Raw BIFF `itab` value from the NAME record.
    ///
    /// - `0` => workbook scope
    /// - `>0` => worksheet scope (`itab-1` is the BIFF sheet index)
    pub(crate) itab: u16,
    pub(crate) refers_to: String,
    pub(crate) hidden: bool,
    pub(crate) comment: Option<String>,
    /// Built-in defined name id when `fBuiltin` is set.
    ///
    /// In BIFF8 this is typically stored as a single-byte id in `rgchName` (with `cch=1`).
    ///
    /// Some writers also populate `chKey`; empirically Excel appears to prefer the id from
    /// `rgchName` and treat `chKey` as a keyboard shortcut, so we only fall back to `chKey` when
    /// the `rgchName` payload is missing.
    pub(crate) builtin_id: Option<u8>,
    /// Raw BIFF8 `rgce` bytes for the defined name formula.
    pub(crate) rgce: Vec<u8>,
}

#[derive(Debug, Clone, Default)]
pub(crate) struct BiffDefinedNames {
    pub(crate) names: Vec<BiffDefinedName>,
    pub(crate) warnings: Vec<String>,
}

fn push_warning(out: &mut BiffDefinedNames, msg: String) {
    if out.warnings.len() < MAX_DEFINED_NAME_WARNINGS {
        out.warnings.push(msg);
        return;
    }

    // When the warning cap is exceeded, append a single suppression message and then ignore any
    // subsequent warnings so the vector size remains bounded.
    if out.warnings.len() == MAX_DEFINED_NAME_WARNINGS {
        out.warnings
            .push("additional defined-name warnings suppressed".to_string());
    }
}

#[derive(Debug, Clone)]
pub(super) struct RawDefinedName {
    pub(super) name: String,
    pub(super) scope_sheet: Option<usize>,
    pub(super) itab: u16,
    pub(super) hidden: bool,
    pub(super) comment: Option<String>,
    pub(super) builtin_id: Option<u8>,
    pub(super) rgce: Vec<u8>,
    /// Trailing data blocks (`rgcb`) referenced by certain ptgs (notably `PtgArray`).
    pub(super) rgcb: Vec<u8>,
}

pub(crate) fn parse_biff_defined_names(
    workbook_stream: &[u8],
    biff: BiffVersion,
    codepage: u16,
    sheet_names: &[String],
) -> Result<BiffDefinedNames, String> {
    let mut out = BiffDefinedNames::default();

    if biff != BiffVersion::Biff8 {
        push_warning(
            &mut out,
            "BIFF defined name import currently supports BIFF8 only".to_string(),
        );
        return Ok(out);
    }

    let mut ctx_tables = workbook_context::build_biff_workbook_context_tables(
        workbook_stream,
        biff,
        codepage,
        sheet_names,
    );
    for warning in ctx_tables.warnings.drain(..) {
        push_warning(&mut out, warning);
    }

    // `ctx` borrows from `ctx_tables`, so move the NAME records out up-front.
    let name_records = std::mem::take(&mut ctx_tables.name_records);
    let ctx = ctx_tables.rgce_decode_context(sheet_names);

    for raw in name_records.into_iter().flatten() {
        let decoded = if raw.rgcb.is_empty() {
            rgce::decode_defined_name_rgce_with_context(&raw.rgce, codepage, &ctx)
        } else {
            rgce::decode_defined_name_rgce_with_context_and_rgcb(
                &raw.rgce,
                &raw.rgcb,
                codepage,
                &ctx,
            )
        };
        for warning in decoded.warnings {
            push_warning(&mut out, format!("defined name `{}`: {warning}", raw.name));
        }

        out.names.push(BiffDefinedName {
            name: raw.name,
            scope_sheet: raw.scope_sheet,
            itab: raw.itab,
            refers_to: decoded.text,
            hidden: raw.hidden,
            comment: raw.comment,
            builtin_id: raw.builtin_id,
            rgce: raw.rgce,
        });
    }

    Ok(out)
}

/// Parse the workbook `NAME` table into a metadata vector suitable for resolving `PtgName` tokens
/// in worksheet formulas.
///
/// BIFF8 `PtgName` stores a 1-based index into the workbook-global `NAME` record order. To preserve
/// those indices even when individual `NAME` records are malformed, this helper inserts a
/// placeholder `#NAME?` entry for any `NAME` record that fails to parse.
///
/// The returned vector is indexed by `iname-1` and can be passed directly as
/// [`rgce::RgceDecodeContext::defined_names`].
///
/// Best-effort semantics:
/// - Only BIFF8 is currently supported; BIFF5 yields an empty table.
/// - Stops scanning at `EOF` or the next `BOF` record (start of the next substream).
/// - Malformed records produce warnings but do not hard-fail.
pub(crate) fn parse_biff_defined_name_metas(
    workbook_stream: &[u8],
    biff: BiffVersion,
    codepage: u16,
    sheet_names: &[String],
) -> (Vec<rgce::DefinedNameMeta>, Vec<String>) {
    let mut warnings: Vec<String> = Vec::new();
    let mut metas: Vec<rgce::DefinedNameMeta> = Vec::new();

    if biff != BiffVersion::Biff8 {
        return (metas, warnings);
    }

    let allows_continuation = |id: u16| id == RECORD_NAME;
    let iter = records::LogicalBiffRecordIter::new(workbook_stream, allows_continuation);

    for record in iter {
        let record = match record {
            Ok(record) => record,
            Err(err) => {
                warnings.push(format!("malformed BIFF record: {err}"));
                break;
            }
        };

        if record.offset != 0 && records::is_bof_record(record.record_id) {
            break;
        }

        match record.record_id {
            RECORD_NAME => match parse_biff8_name_record(&record, codepage, sheet_names) {
                Ok(raw) => metas.push(rgce::DefinedNameMeta {
                    name: raw.name,
                    scope_sheet: raw.scope_sheet,
                }),
                Err(err) => {
                    warnings.push(format!("failed to parse NAME record: {err}"));
                    metas.push(rgce::DefinedNameMeta {
                        name: "#NAME?".to_string(),
                        scope_sheet: None,
                    });
                }
            },
            records::RECORD_EOF => break,
            _ => {}
        }
    }

    (metas, warnings)
}

pub(super) fn parse_biff8_name_record(
    record: &records::LogicalBiffRecord<'_>,
    codepage: u16,
    sheet_names: &[String],
) -> Result<RawDefinedName, String> {
    let fragments: Vec<&[u8]> = record.fragments().collect();
    let mut cursor = FragmentCursor::new(&fragments, 0, 0);

    // Fixed-size `NAME` record header (14 bytes).
    // [MS-XLS] 2.4.150
    let grbit = cursor.read_u16_le()?;
    let ch_key = cursor.read_u8()?;
    let cch = cursor.read_u8()? as usize;
    let cce = cursor.read_u16_le()? as usize;
    let _ixals = cursor.read_u16_le()?;
    let itab = cursor.read_u16_le()?;
    let cch_cust_menu = cursor.read_u8()? as usize;
    let cch_description = cursor.read_u8()? as usize;
    let cch_help_topic = cursor.read_u8()? as usize;
    let cch_status_text = cursor.read_u8()? as usize;

    let hidden = (grbit & NAME_FLAG_HIDDEN) != 0;
    let builtin = (grbit & NAME_FLAG_BUILTIN) != 0;

    let scope_sheet = if itab == 0 {
        None
    } else {
        Some(itab as usize - 1)
    };

    let (builtin_id, name) = if builtin {
        // Built-in names are special:
        //
        // In BIFF8, `rgchName` is stored as a *single byte* built-in name id (no XLUnicodeString
        // option flags), and `cch` MUST be 1. Some producers also populate `chKey` (documented as a
        // keyboard shortcut), but Excel appears to prefer `rgchName` when present.
        //
        // We still consume `rgchName` so `rgce` parsing stays aligned.
        let id_from_name = if cch > 0 {
            Some(cursor.read_u8()?)
        } else {
            None
        };
        if cch > 1 {
            cursor.skip_bytes(cch - 1)?;
        }

        let id = match (id_from_name, ch_key) {
            (Some(id_from_name), ch_key) if ch_key != 0 && id_from_name != ch_key => {
                // [MS-XLS] defines `chKey` as a keyboard shortcut for user-defined names. When
                // `fBuiltin` is set, empirically Excel still prefers the built-in id stored in
                // `rgchName` and treats `chKey` as a shortcut; some producers incorrectly store a
                // built-in id in `chKey` instead.
                log::debug!(
                    "NAME record built-in id mismatch: rgchName=0x{id_from_name:02X} chKey=0x{ch_key:02X} (using rgchName)"
                );
                id_from_name
            }
            (Some(id_from_name), _) => id_from_name,
            (None, ch_key) => {
                // Best-effort fallback for malformed NAME records that omit `rgchName`.
                if ch_key != 0 {
                    log::debug!(
                        "NAME record missing built-in id in rgchName; using chKey=0x{ch_key:02X}"
                    );
                }
                ch_key
            }
        };

        (Some(id), builtin_name_to_string(id))
    } else {
        // `rgchName` (XLUnicodeStringNoCch): flags byte + character bytes.
        let raw_name = cursor.read_biff8_unicode_string_no_cch(cch, codepage)?;
        // Best-effort: BIFF Unicode strings can contain embedded NUL bytes in the wild; strip them
        // so the name matches Excel’s visible name semantics and can pass `formula_model` name
        // validation.
        let stripped = raw_name.replace('\0', "");
        if stripped.is_empty() {
            return Err("NAME record has empty name after stripping NULs".to_string());
        }
        (None, stripped)
    };

    // `rgce`: parsed formula bytes.
    //
    // BIFF8 can insert an additional option-flags byte at the start of a `CONTINUE` fragment when
    // an in-record string (e.g. a `PtgStr` token) is split across fragments. We therefore parse the
    // rgce token stream in a fragment-aware way so those continuation flag bytes are not treated as
    // rgce payload bytes.
    let rgce = cursor.read_biff8_rgce(cce)?;

    let parse_optional_strings =
        |cursor: &mut FragmentCursor<'_>| -> Result<Option<String>, String> {
            if cch_cust_menu > 0 {
                let _ = cursor.read_biff8_unicode_string_no_cch(cch_cust_menu, codepage)?;
            }
            let comment = if cch_description > 0 {
                let raw = cursor.read_biff8_unicode_string_no_cch(cch_description, codepage)?;
                // Best-effort: Excel UIs generally treat embedded NULs as invalid; strip them so the value
                // is usable as `formula_model::DefinedName.comment`.
                let stripped = raw.replace('\0', "");
                (!stripped.is_empty()).then_some(stripped)
            } else {
                None
            };
            if cch_help_topic > 0 {
                let _ = cursor.read_biff8_unicode_string_no_cch(cch_help_topic, codepage)?;
            }
            if cch_status_text > 0 {
                let _ = cursor.read_biff8_unicode_string_no_cch(cch_status_text, codepage)?;
            }
            Ok(comment)
        };

    // Optional trailing `rgcb` blocks referenced by `PtgArray` tokens.
    //
    // BIFF8 stores array constant payloads in the record data immediately after the `rgce` token
    // stream. These `rgcb` bytes are not counted in `cce` and must be consumed before any optional
    // strings (custom menu/description/help/status) that may follow.
    //
    // Best-effort: try parsing `rgcb` when `rgce` likely contains `PtgArray`; if that fails, fall
    // back to the legacy layout that assumes no `rgcb` is present.
    let mut rgcb: Vec<u8> = Vec::new();
    let may_have_ptgarray = rgce.iter().any(|&b| matches!(b, 0x20 | 0x40 | 0x60));
    let comment = if may_have_ptgarray {
        let mut with_rgcb = cursor.clone();
        match read_biff8_rgcb_for_rgce_arrays(&rgce, &mut with_rgcb) {
            Ok(rgcb_candidate) => match parse_optional_strings(&mut with_rgcb) {
                Ok(comment) => {
                    rgcb = rgcb_candidate;
                    comment
                }
                Err(_) => parse_optional_strings(&mut cursor)?,
            },
            Err(_) => parse_optional_strings(&mut cursor)?,
        }
    } else {
        parse_optional_strings(&mut cursor)?
    };

    if let Some(scope) = scope_sheet {
        if scope >= sheet_names.len() {
            log::warn!(
                "NAME record `{name}` has out-of-range itab={itab} (sheet count={})",
                sheet_names.len()
            );
        }
    }

    Ok(RawDefinedName {
        name,
        scope_sheet,
        itab,
        hidden,
        comment,
        builtin_id,
        rgce,
        rgcb,
    })
}

fn read_biff8_rgcb_for_rgce_arrays(
    rgce: &[u8],
    cursor: &mut FragmentCursor<'_>,
) -> Result<Vec<u8>, String> {
    let mut out = Vec::<u8>::new();
    scan_biff8_rgce_for_array_constants(rgce, cursor, &mut out)?;
    Ok(out)
}

fn scan_biff8_rgce_for_array_constants(
    rgce: &[u8],
    cursor: &mut FragmentCursor<'_>,
    rgcb_out: &mut Vec<u8>,
) -> Result<(), String> {
    fn inner(
        input: &[u8],
        cursor: &mut FragmentCursor<'_>,
        rgcb_out: &mut Vec<u8>,
    ) -> Result<(), String> {
        let mut i = 0usize;
        while i < input.len() {
            let ptg = *input.get(i).ok_or_else(|| "unexpected end of rgce stream".to_string())?;
            i = i.saturating_add(1);

            match ptg {
                // PtgExp / PtgTbl: [rw: u16][col: u16]
                0x01 | 0x02 => {
                    if i + 4 > input.len() {
                        return Err("unexpected end of rgce stream".to_string());
                    }
                    i += 4;
                }
                // Binary operators and simple operators with no payload.
                0x03..=0x16 | 0x2F => {}
                // PtgStr (ShortXLUnicodeString): variable.
                0x17 => {
                    let remaining = input
                        .get(i..)
                        .ok_or_else(|| "unexpected end of rgce stream".to_string())?;
                    let (_s, consumed) = strings::parse_biff8_short_string(remaining, 1252)
                        .map_err(|e| format!("failed to parse PtgStr: {e}"))?;
                    i = i
                        .checked_add(consumed)
                        .ok_or_else(|| "rgce offset overflow".to_string())?;
                    if i > input.len() {
                        return Err("unexpected end of rgce stream".to_string());
                    }
                }
                // PtgExtend* tokens (ptg=0x18 variants): [etpg: u8][payload...]
                0x18 | 0x38 | 0x58 | 0x78 => {
                    let etpg = *input
                        .get(i)
                        .ok_or_else(|| "unexpected end of rgce stream".to_string())?;
                    i += 1;
                    match etpg {
                        0x19 => {
                            // PtgList: fixed 12-byte payload.
                            if i + 12 > input.len() {
                                return Err("unexpected end of rgce stream".to_string());
                            }
                            i += 12;
                        }
                        _ => {
                            // Opaque 5-byte payload (see decoder heuristics).
                            //
                            // The ptg itself is followed by 5 bytes; since we consumed the first
                            // one as the "etpg" discriminator above, skip the remaining 4 bytes.
                            if i + 4 > input.len() {
                                return Err("unexpected end of rgce stream".to_string());
                            }
                            i += 4;
                        }
                    }
                }
                // PtgAttr: [grbit: u8][wAttr: u16] + optional jump table for tAttrChoose.
                0x19 => {
                    if i + 3 > input.len() {
                        return Err("unexpected end of rgce stream".to_string());
                    }
                    let grbit = input[i];
                    let w_attr = u16::from_le_bytes([input[i + 1], input[i + 2]]) as usize;
                    i += 3;

                    const T_ATTR_CHOOSE: u8 = 0x04;
                    if grbit & T_ATTR_CHOOSE != 0 {
                        let needed = w_attr
                            .checked_mul(2)
                            .ok_or_else(|| "PtgAttr jump table length overflow".to_string())?;
                        if i + needed > input.len() {
                            return Err("unexpected end of rgce stream".to_string());
                        }
                        i += needed;
                    }
                }
                // PtgErr / PtgBool: 1 byte.
                0x1C | 0x1D => {
                    if i + 1 > input.len() {
                        return Err("unexpected end of rgce stream".to_string());
                    }
                    i += 1;
                }
                // PtgInt: 2 bytes.
                0x1E => {
                    if i + 2 > input.len() {
                        return Err("unexpected end of rgce stream".to_string());
                    }
                    i += 2;
                }
                // PtgNum: 8 bytes.
                0x1F => {
                    if i + 8 > input.len() {
                        return Err("unexpected end of rgce stream".to_string());
                    }
                    i += 8;
                }
                // PtgArray: [unused: 7 bytes] + array constant values stored in rgcb.
                0x20 | 0x40 | 0x60 => {
                    if i + 7 > input.len() {
                        return Err("unexpected end of rgce stream".to_string());
                    }
                    i += 7;
                    read_biff8_array_constant_from_record(cursor, rgcb_out)?;
                }
                // PtgFunc: 2 bytes.
                0x21 | 0x41 | 0x61 => {
                    if i + 2 > input.len() {
                        return Err("unexpected end of rgce stream".to_string());
                    }
                    i += 2;
                }
                // PtgFuncVar: 3 bytes.
                0x22 | 0x42 | 0x62 => {
                    if i + 3 > input.len() {
                        return Err("unexpected end of rgce stream".to_string());
                    }
                    i += 3;
                }
                // PtgName: 6 bytes.
                0x23 | 0x43 | 0x63 => {
                    if i + 6 > input.len() {
                        return Err("unexpected end of rgce stream".to_string());
                    }
                    i += 6;
                }
                // PtgRef: 4 bytes.
                0x24 | 0x44 | 0x64 => {
                    if i + 4 > input.len() {
                        return Err("unexpected end of rgce stream".to_string());
                    }
                    i += 4;
                }
                // PtgArea: 8 bytes.
                0x25 | 0x45 | 0x65 => {
                    if i + 8 > input.len() {
                        return Err("unexpected end of rgce stream".to_string());
                    }
                    i += 8;
                }
                // PtgMem* tokens: [cce: u16][rgce: cce bytes]
                0x26 | 0x46 | 0x66 | 0x27 | 0x47 | 0x67 | 0x28 | 0x48 | 0x68 | 0x29 | 0x49
                | 0x69 | 0x2E | 0x4E | 0x6E => {
                    if i + 2 > input.len() {
                        return Err("unexpected end of rgce stream".to_string());
                    }
                    let cce = u16::from_le_bytes([input[i], input[i + 1]]) as usize;
                    i += 2;
                    let sub = input
                        .get(i..i + cce)
                        .ok_or_else(|| "unexpected end of rgce stream".to_string())?;
                    inner(sub, cursor, rgcb_out)?;
                    i += cce;
                }
                // PtgRefErr: 4 bytes.
                0x2A | 0x4A | 0x6A => {
                    if i + 4 > input.len() {
                        return Err("unexpected end of rgce stream".to_string());
                    }
                    i += 4;
                }
                // PtgAreaErr: 8 bytes.
                0x2B | 0x4B | 0x6B => {
                    if i + 8 > input.len() {
                        return Err("unexpected end of rgce stream".to_string());
                    }
                    i += 8;
                }
                // PtgRefN: 4 bytes.
                0x2C | 0x4C | 0x6C => {
                    if i + 4 > input.len() {
                        return Err("unexpected end of rgce stream".to_string());
                    }
                    i += 4;
                }
                // PtgAreaN: 8 bytes.
                0x2D | 0x4D | 0x6D => {
                    if i + 8 > input.len() {
                        return Err("unexpected end of rgce stream".to_string());
                    }
                    i += 8;
                }
                // PtgNameX: 6 bytes.
                0x39 | 0x59 | 0x79 => {
                    if i + 6 > input.len() {
                        return Err("unexpected end of rgce stream".to_string());
                    }
                    i += 6;
                }
                // PtgRef3d: 6 bytes.
                0x3A | 0x5A | 0x7A => {
                    if i + 6 > input.len() {
                        return Err("unexpected end of rgce stream".to_string());
                    }
                    i += 6;
                }
                // PtgArea3d: 10 bytes.
                0x3B | 0x5B | 0x7B => {
                    if i + 10 > input.len() {
                        return Err("unexpected end of rgce stream".to_string());
                    }
                    i += 10;
                }
                // PtgRefErr3d: 6 bytes.
                0x3C | 0x5C | 0x7C => {
                    if i + 6 > input.len() {
                        return Err("unexpected end of rgce stream".to_string());
                    }
                    i += 6;
                }
                // PtgAreaErr3d: 10 bytes.
                0x3D | 0x5D | 0x7D => {
                    if i + 10 > input.len() {
                        return Err("unexpected end of rgce stream".to_string());
                    }
                    i += 10;
                }
                // PtgRefN3d: 6 bytes.
                0x3E | 0x5E | 0x7E => {
                    if i + 6 > input.len() {
                        return Err("unexpected end of rgce stream".to_string());
                    }
                    i += 6;
                }
                // PtgAreaN3d: 10 bytes.
                0x3F | 0x5F | 0x7F => {
                    if i + 10 > input.len() {
                        return Err("unexpected end of rgce stream".to_string());
                    }
                    i += 10;
                }
                other => {
                    return Err(format!(
                        "unsupported rgce token 0x{other:02X} while scanning for PtgArray constants"
                    ));
                }
            }
        }
        Ok(())
    }

    inner(rgce, cursor, rgcb_out)
}

fn read_biff8_array_constant_from_record(
    cursor: &mut FragmentCursor<'_>,
    out: &mut Vec<u8>,
) -> Result<(), String> {
    // BIFF8 array constant payload stream stored as trailing `rgcb` bytes.
    // See MS-XLS 2.5.198.8 PtgArray.
    //
    // Layout:
    //   [cols_minus1: u16][rows_minus1: u16][values...]
    //
    // Values are stored row-major and each starts with a type byte:
    //   0x00 = empty
    //   0x01 = number (f64)
    //   0x02 = string ([cch: u16][utf16 chars...])
    //   0x04 = bool ([b: u8])
    //   0x10 = error ([code: u8])
    const MAX_ARRAY_CELLS: usize = 4096;

    let cols_minus1 = cursor.read_u16_le()?;
    let rows_minus1 = cursor.read_u16_le()?;
    out.extend_from_slice(&cols_minus1.to_le_bytes());
    out.extend_from_slice(&rows_minus1.to_le_bytes());

    let cols = cols_minus1 as usize + 1;
    let rows = rows_minus1 as usize + 1;
    let total = cols.saturating_mul(rows);
    if total > MAX_ARRAY_CELLS {
        return Err(format!(
            "array constant is too large to parse (rows={rows}, cols={cols})"
        ));
    }

    for _ in 0..total {
        let ty = cursor.read_u8()?;
        out.push(ty);
        match ty {
            0x00 => {}
            0x01 => {
                let bytes = cursor.read_bytes(8)?;
                out.extend_from_slice(&bytes);
            }
            0x02 => {
                let cch = cursor.read_u16_le()?;
                out.extend_from_slice(&cch.to_le_bytes());
                let byte_len = (cch as usize)
                    .checked_mul(2)
                    .ok_or_else(|| "array string length overflow".to_string())?;
                let bytes = cursor.read_bytes(byte_len)?;
                out.extend_from_slice(&bytes);
            }
            0x04 | 0x10 => {
                out.push(cursor.read_u8()?);
            }
            other => {
                return Err(format!(
                    "unsupported array constant element type 0x{other:02X}"
                ));
            }
        }
    }

    Ok(())
}

fn builtin_name_to_string(id: u8) -> String {
    match id {
        // Built-in name ids from [MS-XLS] 2.5.114 (Lbl.chKey when fBuiltin is set).
        0x00 => "_xlnm.Consolidate_Area".to_string(),
        0x01 => "_xlnm.Auto_Open".to_string(),
        0x02 => "_xlnm.Auto_Close".to_string(),
        0x03 => "_xlnm.Extract".to_string(),
        0x04 => "_xlnm.Database".to_string(),
        0x05 => "_xlnm.Criteria".to_string(),
        0x06 => formula_model::XLNM_PRINT_AREA.to_string(),
        0x07 => formula_model::XLNM_PRINT_TITLES.to_string(),
        0x08 => "_xlnm.Recorder".to_string(),
        0x09 => "_xlnm.Data_Form".to_string(),
        0x0A => "_xlnm.Auto_Activate".to_string(),
        0x0B => "_xlnm.Auto_Deactivate".to_string(),
        0x0C => "_xlnm.Sheet_Title".to_string(),
        0x0D => formula_model::XLNM_FILTER_DATABASE.to_string(),
        other => {
            log::warn!("unsupported BIFF8 built-in NAME id 0x{other:02X}");
            format!("_xlnm.Builtin_0x{other:02X}")
        }
    }
}

#[derive(Clone)]
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

    fn advance_fragment_in_biff8_string(&mut self, is_unicode: &mut bool) -> Result<(), String> {
        self.advance_fragment()?;
        // When a BIFF8 string spans a CONTINUE boundary, Excel inserts a 1-byte option flags prefix
        // at the start of the continued fragment. The only relevant bit for formula string tokens is
        // `fHighByte` (unicode vs compressed).
        let cont_flags = self.read_u8()?;
        *is_unicode = (cont_flags & STR_FLAG_HIGH_BYTE) != 0;
        Ok(())
    }

    fn read_biff8_string_bytes(
        &mut self,
        mut n: usize,
        is_unicode: &mut bool,
    ) -> Result<Vec<u8>, String> {
        // Read `n` canonical bytes from a BIFF8 continued string payload, skipping the 1-byte
        // continuation flags prefix that appears at the start of each continued fragment.
        let mut out = Vec::with_capacity(n);
        while n > 0 {
            if self.remaining_in_fragment() == 0 {
                self.advance_fragment_in_biff8_string(is_unicode)?;
                continue;
            }
            let available = self.remaining_in_fragment();
            let take = n.min(available);
            let bytes = self.read_exact_from_current(take)?;
            out.extend_from_slice(bytes);
            n -= take;
        }
        Ok(out)
    }

    fn skip_biff8_string_bytes(
        &mut self,
        mut n: usize,
        is_unicode: &mut bool,
    ) -> Result<(), String> {
        // Skip `n` canonical bytes from a BIFF8 continued string payload, consuming any inserted
        // continuation flags bytes at fragment boundaries.
        while n > 0 {
            if self.remaining_in_fragment() == 0 {
                self.advance_fragment_in_biff8_string(is_unicode)?;
                continue;
            }
            let available = self.remaining_in_fragment();
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

        let mut is_unicode = (flags & STR_FLAG_HIGH_BYTE) != 0;

        let richtext_runs = if flags & STR_FLAG_RICH_TEXT != 0 {
            let bytes = self.read_biff8_string_bytes(2, &mut is_unicode)?;
            u16::from_le_bytes([bytes[0], bytes[1]]) as usize
        } else {
            0
        };

        let ext_size = if flags & STR_FLAG_EXT != 0 {
            let bytes = self.read_biff8_string_bytes(4, &mut is_unicode)?;
            u32::from_le_bytes([bytes[0], bytes[1], bytes[2], bytes[3]]) as usize
        } else {
            0
        };

        let mut remaining_chars = cch;
        let mut out = String::new();

        while remaining_chars > 0 {
            if self.remaining_in_fragment() == 0 {
                self.advance_fragment_in_biff8_string(&mut is_unicode)?;
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
        let extra_len = richtext_bytes
            .checked_add(ext_size)
            .ok_or_else(|| "string ext payload length overflow".to_string())?;
        self.skip_biff8_string_bytes(extra_len, &mut is_unicode)?;

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
                // PtgExp / PtgTbl: shared/array formula tokens (not expected in NAME, but consume
                // payload to keep the stream aligned if they appear).
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

                    let mut is_unicode = (flags & STR_FLAG_HIGH_BYTE) != 0;

                    let richtext_runs = if (flags & STR_FLAG_RICH_TEXT) != 0 {
                        let bytes = self.read_biff8_string_bytes(2, &mut is_unicode)?;
                        out.extend_from_slice(&bytes);
                        u16::from_le_bytes([bytes[0], bytes[1]]) as usize
                    } else {
                        0
                    };

                    let ext_size = if (flags & STR_FLAG_EXT) != 0 {
                        let bytes = self.read_biff8_string_bytes(4, &mut is_unicode)?;
                        out.extend_from_slice(&bytes);
                        u32::from_le_bytes([bytes[0], bytes[1], bytes[2], bytes[3]]) as usize
                    } else {
                        0
                    };

                    let mut remaining_chars = cch;

                    while remaining_chars > 0 {
                        if self.remaining_in_fragment() == 0 {
                            self.advance_fragment_in_biff8_string(&mut is_unicode)?;
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
                        if out.len() + take_bytes > cce {
                            return Err(
                                "PtgStr character payload exceeds declared rgce length"
                                    .to_string(),
                            );
                        }
                        let bytes = self.read_exact_from_current(take_bytes)?;
                        out.extend_from_slice(bytes);
                        remaining_chars -= take_chars;
                    }

                    let richtext_bytes = richtext_runs
                        .checked_mul(4)
                        .ok_or_else(|| "rich text run count overflow".to_string())?;
                    let extra_len = richtext_bytes
                        .checked_add(ext_size)
                        .ok_or_else(|| "PtgStr extra payload length overflow".to_string())?;
                    if extra_len > 0 {
                        let remaining = cce.saturating_sub(out.len());
                        if extra_len > remaining {
                            return Err(
                                "PtgStr extra payload exceeds declared rgce length".to_string(),
                            );
                        }
                        let extra = self.read_biff8_string_bytes(extra_len, &mut is_unicode)?;
                        out.extend_from_slice(&extra);
                    }
                }
                // PtgExtend* token 0x18 (and class variants).
                //
                // Excel can embed newer operand subtypes behind `PtgExtend` using an `etpg` subtype
                // byte. Structured references (tables) use `etpg=0x19` (PtgList) and include a
                // 12-byte payload.
                //
                // Some `.xls` files in the wild also include a 5-byte opaque token with this ptg
                // value (calamine treats it as a 5-byte payload and skips it). To remain compatible
                // while keeping the token stream aligned, we parse `etpg=0x19` as a 13-byte payload
                // (etpg + 12 bytes) and treat all other subtypes as a 5-byte opaque payload.
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
                // PtgArray (7 bytes) [MS-XLS 2.5.198.8]
                //
                // The actual array constant values live in `rgcb` (trailing data blocks), but the
                // token itself contains a 7-byte header we must preserve to keep `rgce` aligned.
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

    fn xl_unicode_string_no_cch_compressed(s: &str) -> Vec<u8> {
        let mut out = Vec::with_capacity(1 + s.len());
        out.push(0); // flags (compressed)
        out.extend_from_slice(s.as_bytes());
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
        let r_name = record(
            RECORD_NAME,
            &[header.clone(), name_str.clone(), first_rgce.to_vec()].concat(),
        );
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
            record(
                RECORD_NAME,
                &[header, name_str, first_rgce.to_vec()].concat(),
            ),
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
            // PtgRefN: rw=0, col=0xC000 (row+col relative) => A1 (base A1).
            vec![0x2C, 0x00, 0x00, 0x00, 0xC0],
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
            record(
                RECORD_NAME,
                &[header, name_str, first_rgce.to_vec()].concat(),
            ),
            record(records::RECORD_CONTINUE, &continue_payload),
            record(records::RECORD_EOF, &[]),
        ]
        .concat();

        let parsed =
            parse_biff_defined_names(&stream, BiffVersion::Biff8, 1252, &[]).expect("parse names");
        assert_eq!(parsed.names.len(), 1);
        assert_eq!(parsed.names[0].name, name);
        assert_eq!(parsed.names[0].refers_to, "A1&\"ABCDE\"");
        assert!(
            parsed
                .warnings
                .iter()
                .any(|w| w.contains("interpreted relative to A1")),
            "warnings={:?}",
            parsed.warnings
        );
    }

    #[test]
    fn parses_defined_name_with_referr_before_continued_ptgstr_token() {
        let name = "ErrStrName";
        let literal = "ABCDE";

        // rgce for `#REF!&"ABCDE"`.
        let rgce: Vec<u8> = [
            vec![0x2A, 0x00, 0x00, 0x00, 0x00],   // PtgRefErr + payload
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
        let first_rgce = &rgce[..10]; // PtgRefErr (5) + PtgStr header (3) + "AB" (2)
        let second_chars = &literal.as_bytes()[2..]; // "CDE"

        let mut continue_payload = Vec::new();
        continue_payload.push(0); // continued segment option flags (fHighByte=0)
        continue_payload.extend_from_slice(second_chars);
        continue_payload.push(0x08); // PtgConcat

        let stream = [
            record(records::RECORD_BOF_BIFF8, &[0u8; 16]),
            record(
                RECORD_NAME,
                &[header, name_str, first_rgce.to_vec()].concat(),
            ),
            record(records::RECORD_CONTINUE, &continue_payload),
            record(records::RECORD_EOF, &[]),
        ]
        .concat();

        let parsed =
            parse_biff_defined_names(&stream, BiffVersion::Biff8, 1252, &[]).expect("parse names");
        assert_eq!(parsed.names.len(), 1);
        assert_eq!(parsed.names[0].name, name);
        assert_eq!(parsed.names[0].refers_to, "#REF!&\"ABCDE\"");
        assert!(parsed.warnings.is_empty(), "warnings={:?}", parsed.warnings);
    }

    #[test]
    fn parses_name_record_with_namex_before_continued_ptgstr_token() {
        let name = "NameXStr";
        let literal = "ABCDE";

        // `rgce` containing:
        //   PtgNameX (6 bytes)
        //   PtgStr (string literal)
        //   PtgConcat
        //
        // If `PtgNameX` isn't recognized, the parser falls back to raw byte copying, which can
        // corrupt a continued `PtgStr` by treating the continued-segment option flags byte as part
        // of the formula payload.
        let rgce: Vec<u8> = [
            vec![0x39, 0x00, 0x00, 0x01, 0x00, 0x00, 0x00], // PtgNameX + payload (ixti=0, iname=1)
            vec![0x17, literal.len() as u8, 0u8],           // PtgStr + cch + flags (compressed)
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
        let first_rgce = &rgce[..12]; // PtgNameX (7) + PtgStr header (3) + "AB" (2)
        let second_chars = &literal.as_bytes()[2..]; // "CDE"

        let mut continue_payload = Vec::new();
        continue_payload.push(0); // continued segment option flags (fHighByte=0)
        continue_payload.extend_from_slice(second_chars);
        continue_payload.push(0x08); // PtgConcat

        let stream = [
            record(records::RECORD_BOF_BIFF8, &[0u8; 16]),
            record(
                RECORD_NAME,
                &[header, name_str, first_rgce.to_vec()].concat(),
            ),
            record(records::RECORD_CONTINUE, &continue_payload),
            record(records::RECORD_EOF, &[]),
        ]
        .concat();

        let allows_continuation = |id: u16| id == RECORD_NAME;
        let iter = records::LogicalBiffRecordIter::new(&stream, allows_continuation);
        let record = iter
            .filter_map(Result::ok)
            .find(|r| r.record_id == RECORD_NAME)
            .expect("NAME record");
        let raw = parse_biff8_name_record(&record, 1252, &[]).expect("parse NAME");
        assert_eq!(raw.name, name);
        assert_eq!(raw.rgce, rgce);
    }

    #[test]
    fn parses_name_record_with_ptg18_before_continued_ptgstr_token() {
        let name = "Ptg18Str";
        let literal = "ABCDE";

        let rgce: Vec<u8> = [
            vec![0x18, 0x11, 0x22, 0x33, 0x44, 0x55], // ptg=0x18 + 5-byte opaque payload
            vec![0x17, literal.len() as u8, 0u8],     // PtgStr + cch + flags (compressed)
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

        let name_str = xl_unicode_string_no_cch_compressed(name);

        // Split the PtgStr character bytes across the CONTINUE boundary after "AB".
        let first_rgce = &rgce[..(6 + 3 + 2)]; // ptg18 (6) + ptgstr header (3) + "AB" (2)
        let second_chars = &literal.as_bytes()[2..]; // "CDE"

        let mut continue_payload = Vec::new();
        continue_payload.push(0); // continued segment option flags (fHighByte=0)
        continue_payload.extend_from_slice(second_chars);

        let stream = [
            record(records::RECORD_BOF_BIFF8, &[0u8; 16]),
            record(
                RECORD_NAME,
                &[header, name_str, first_rgce.to_vec()].concat(),
            ),
            record(records::RECORD_CONTINUE, &continue_payload),
            record(records::RECORD_EOF, &[]),
        ]
        .concat();

        let allows_continuation = |id: u16| id == RECORD_NAME;
        let iter = records::LogicalBiffRecordIter::new(&stream, allows_continuation);
        let record = iter
            .filter_map(Result::ok)
            .find(|r| r.record_id == RECORD_NAME)
            .expect("NAME record");
        let raw = parse_biff8_name_record(&record, 1252, &[]).expect("parse NAME");
        assert_eq!(raw.name, name);
        assert_eq!(raw.rgce, rgce);
    }

    #[test]
    fn parses_name_record_with_ptglist_before_continued_ptgstr_token() {
        let name = "PtgListStr";
        let literal = "ABCDE";

        // rgce: PtgExtend(etpg=0x19 PtgList) + PtgStr.
        let table_id = 1u32;
        let flags = 0u16;
        let col_first = 2u16;
        let col_last = 2u16;
        let reserved = 0u16;
        let rgce: Vec<u8> = [
            vec![0x18, 0x19], // PtgExtend + etpg=PtgList
            table_id.to_le_bytes().to_vec(),
            flags.to_le_bytes().to_vec(),
            col_first.to_le_bytes().to_vec(),
            col_last.to_le_bytes().to_vec(),
            reserved.to_le_bytes().to_vec(),
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

        let name_str = xl_unicode_string_no_cch_compressed(name);

        // Split the PtgStr character bytes across the CONTINUE boundary after "AB".
        let ptglist_len = 1 + 1 + 12; // ptg + etpg + payload
        let first_rgce = &rgce[..(ptglist_len + 3 + 2)]; // PtgList (14) + PtgStr header (3) + "AB" (2)
        let second_chars = &literal.as_bytes()[2..]; // "CDE"

        let mut continue_payload = Vec::new();
        continue_payload.push(0); // continued segment option flags (fHighByte=0)
        continue_payload.extend_from_slice(second_chars);

        let stream = [
            record(records::RECORD_BOF_BIFF8, &[0u8; 16]),
            record(
                RECORD_NAME,
                &[header, name_str, first_rgce.to_vec()].concat(),
            ),
            record(records::RECORD_CONTINUE, &continue_payload),
            record(records::RECORD_EOF, &[]),
        ]
        .concat();

        let allows_continuation = |id: u16| id == RECORD_NAME;
        let iter = records::LogicalBiffRecordIter::new(&stream, allows_continuation);
        let record = iter
            .filter_map(Result::ok)
            .find(|r| r.record_id == RECORD_NAME)
            .expect("NAME record");
        let raw = parse_biff8_name_record(&record, 1252, &[]).expect("parse NAME");
        assert_eq!(raw.name, name);
        assert_eq!(raw.rgce, rgce);
    }

    #[test]
    fn parses_name_record_with_spill_before_continued_ptgstr_token() {
        let name = "SpillStr";
        let literal = "ABCDE";

        let rgce: Vec<u8> = [
            // PtgRef (A1) with relative row/col flags so the decoder prints `A1`.
            vec![0x24, 0x00, 0x00, 0x00, 0xC0],
            vec![0x2F],                           // spill postfix
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

        let name_str = xl_unicode_string_no_cch_compressed(name);

        // Split the PtgStr character bytes across the CONTINUE boundary after "AB".
        let first_rgce = &rgce[..(5 + 1 + 3 + 2)]; // ptgref (5) + spill (1) + ptgstr header (3) + "AB" (2)
        let second_chars = &literal.as_bytes()[2..]; // "CDE"

        let mut continue_payload = Vec::new();
        continue_payload.push(0); // continued segment option flags (fHighByte=0)
        continue_payload.extend_from_slice(second_chars);

        let stream = [
            record(records::RECORD_BOF_BIFF8, &[0u8; 16]),
            record(
                RECORD_NAME,
                &[header, name_str, first_rgce.to_vec()].concat(),
            ),
            record(records::RECORD_CONTINUE, &continue_payload),
            record(records::RECORD_EOF, &[]),
        ]
        .concat();

        let allows_continuation = |id: u16| id == RECORD_NAME;
        let iter = records::LogicalBiffRecordIter::new(&stream, allows_continuation);
        let record = iter
            .filter_map(Result::ok)
            .find(|r| r.record_id == RECORD_NAME)
            .expect("NAME record");
        let raw = parse_biff8_name_record(&record, 1252, &[]).expect("parse NAME");
        assert_eq!(raw.name, name);
        assert_eq!(raw.rgce, rgce);
    }

    #[test]
    fn parses_name_record_with_richtext_ptgstr_split_inside_char_bytes() {
        let name = "RichStr";
        let literal = "ABCDE";
        let c_run = 1u16;
        let rg_run = [0x11, 0x22, 0x33, 0x44];

        let rgce: Vec<u8> = [
            vec![0x17, literal.len() as u8, STR_FLAG_RICH_TEXT],
            c_run.to_le_bytes().to_vec(),
            literal.as_bytes().to_vec(),
            rg_run.to_vec(),
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

        let name_str = xl_unicode_string_no_cch_compressed(name);

        // Split after "AB" inside the character bytes.
        let first_rgce_len = 3 + 2 + 2; // header + cRun + "AB"
        let first_rgce = &rgce[..first_rgce_len];
        let remaining = &rgce[first_rgce_len..];

        let mut continue_payload = Vec::new();
        continue_payload.push(0); // continued segment option flags (fHighByte=0)
        continue_payload.extend_from_slice(remaining);

        let stream = [
            record(records::RECORD_BOF_BIFF8, &[0u8; 16]),
            record(
                RECORD_NAME,
                &[header, name_str, first_rgce.to_vec()].concat(),
            ),
            record(records::RECORD_CONTINUE, &continue_payload),
            record(records::RECORD_EOF, &[]),
        ]
        .concat();

        let allows_continuation = |id: u16| id == RECORD_NAME;
        let iter = records::LogicalBiffRecordIter::new(&stream, allows_continuation);
        let record = iter
            .filter_map(Result::ok)
            .find(|r| r.record_id == RECORD_NAME)
            .expect("NAME record");
        let raw = parse_biff8_name_record(&record, 1252, &[]).expect("parse NAME");
        assert_eq!(raw.name, name);
        assert_eq!(raw.rgce, rgce);
    }

    #[test]
    fn parses_name_record_with_richtext_ptgstr_split_between_crun_bytes() {
        let name = "RichCrun";
        let literal = "ABCDE";
        let c_run = 1u16;
        let rg_run = [0x11, 0x22, 0x33, 0x44];

        let rgce: Vec<u8> = [
            vec![0x17, literal.len() as u8, STR_FLAG_RICH_TEXT],
            c_run.to_le_bytes().to_vec(),
            literal.as_bytes().to_vec(),
            rg_run.to_vec(),
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

        let name_str = xl_unicode_string_no_cch_compressed(name);

        // Split between the two bytes of `cRun`.
        let first_rgce_len = 3 + 1; // header + low byte of cRun
        let first_rgce = &rgce[..first_rgce_len];
        let remaining = &rgce[first_rgce_len..];

        let mut continue_payload = Vec::new();
        continue_payload.push(0); // continued segment option flags (fHighByte=0)
        continue_payload.extend_from_slice(remaining);

        let stream = [
            record(records::RECORD_BOF_BIFF8, &[0u8; 16]),
            record(
                RECORD_NAME,
                &[header, name_str, first_rgce.to_vec()].concat(),
            ),
            record(records::RECORD_CONTINUE, &continue_payload),
            record(records::RECORD_EOF, &[]),
        ]
        .concat();

        let allows_continuation = |id: u16| id == RECORD_NAME;
        let iter = records::LogicalBiffRecordIter::new(&stream, allows_continuation);
        let record = iter
            .filter_map(Result::ok)
            .find(|r| r.record_id == RECORD_NAME)
            .expect("NAME record");
        let raw = parse_biff8_name_record(&record, 1252, &[]).expect("parse NAME");
        assert_eq!(raw.name, name);
        assert_eq!(raw.rgce, rgce);
    }

    #[test]
    fn parses_name_record_with_ext_ptgstr_split_inside_ext_bytes() {
        let name = "ExtStr";
        let literal = "ABCDE";
        let ext = [0xDE, 0xAD, 0xBE, 0xEF];

        let rgce: Vec<u8> = [
            vec![0x17, literal.len() as u8, STR_FLAG_EXT],
            (ext.len() as u32).to_le_bytes().to_vec(),
            literal.as_bytes().to_vec(),
            ext.to_vec(),
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

        let name_str = xl_unicode_string_no_cch_compressed(name);

        // Split inside the ext bytes.
        let first_rgce_len = 3 + 4 + literal.len() + 2; // header + cbExtRst + chars + first 2 ext bytes
        let first_rgce = &rgce[..first_rgce_len];
        let remaining = &rgce[first_rgce_len..];

        let mut continue_payload = Vec::new();
        continue_payload.push(0); // continued segment option flags (fHighByte=0)
        continue_payload.extend_from_slice(remaining);

        let stream = [
            record(records::RECORD_BOF_BIFF8, &[0u8; 16]),
            record(
                RECORD_NAME,
                &[header, name_str, first_rgce.to_vec()].concat(),
            ),
            record(records::RECORD_CONTINUE, &continue_payload),
            record(records::RECORD_EOF, &[]),
        ]
        .concat();

        let allows_continuation = |id: u16| id == RECORD_NAME;
        let iter = records::LogicalBiffRecordIter::new(&stream, allows_continuation);
        let record = iter
            .filter_map(Result::ok)
            .find(|r| r.record_id == RECORD_NAME)
            .expect("NAME record");
        let raw = parse_biff8_name_record(&record, 1252, &[]).expect("parse NAME");
        assert_eq!(raw.name, name);
        assert_eq!(raw.rgce, rgce);
    }

    #[test]
    fn parses_name_record_with_richtext_and_ext_ptgstr_split_inside_rgrun_bytes() {
        let name = "RichExt";
        let literal = "ABCDE";
        let c_run = 1u16;
        let rg_run = [0x11, 0x22, 0x33, 0x44];
        let ext = [0x55, 0x66, 0x77, 0x88];

        let rgce: Vec<u8> = [
            vec![0x17, literal.len() as u8, STR_FLAG_RICH_TEXT | STR_FLAG_EXT],
            c_run.to_le_bytes().to_vec(),
            (ext.len() as u32).to_le_bytes().to_vec(),
            literal.as_bytes().to_vec(),
            rg_run.to_vec(),
            ext.to_vec(),
        ]
        .concat();

        // Split inside `rgRun` bytes.
        let first_rgce_len = 3 + 2 + 4 + literal.len() + 2; // header + cRun + cbExtRst + chars + first 2 rgRun bytes
        let first_rgce = &rgce[..first_rgce_len];
        let remaining = &rgce[first_rgce_len..];

        let mut header = Vec::new();
        header.extend_from_slice(&0u16.to_le_bytes()); // grbit
        header.push(0); // chKey
        header.push(name.len() as u8); // cch
        header.extend_from_slice(&(rgce.len() as u16).to_le_bytes()); // cce
        header.extend_from_slice(&0u16.to_le_bytes()); // ixals
        header.extend_from_slice(&0u16.to_le_bytes()); // itab
        header.extend_from_slice(&[0, 0, 0, 0]); // cchCustMenu, cchDescription, cchHelpTopic, cchStatusText

        let name_str = xl_unicode_string_no_cch_compressed(name);

        let mut continue_payload = Vec::new();
        continue_payload.push(0); // continued segment option flags (fHighByte=0)
        continue_payload.extend_from_slice(remaining);

        let stream = [
            record(records::RECORD_BOF_BIFF8, &[0u8; 16]),
            record(
                RECORD_NAME,
                &[header, name_str, first_rgce.to_vec()].concat(),
            ),
            record(records::RECORD_CONTINUE, &continue_payload),
            record(records::RECORD_EOF, &[]),
        ]
        .concat();

        let allows_continuation = |id: u16| id == RECORD_NAME;
        let iter = records::LogicalBiffRecordIter::new(&stream, allows_continuation);
        let record = iter
            .filter_map(Result::ok)
            .find(|r| r.record_id == RECORD_NAME)
            .expect("NAME record");
        let raw = parse_biff8_name_record(&record, 1252, &[]).expect("parse NAME");
        assert_eq!(raw.name, name);
        assert_eq!(raw.rgce, rgce);
    }

    #[test]
    fn ptgstr_ext_size_out_of_bounds_errors_without_allocating_unbounded() {
        // Crafted PtgStr that sets fExtSt with a huge cbExtRst. The parser should not try to
        // allocate cbExtRst bytes; it should fail fast and return an error.
        let name = "BigExt";
        let rgce = vec![0x17, 0u8, STR_FLAG_EXT, 0xFF, 0xFF, 0xFF, 0xFF];

        let mut header = Vec::new();
        header.extend_from_slice(&0u16.to_le_bytes()); // grbit
        header.push(0); // chKey
        header.push(name.len() as u8); // cch
        header.extend_from_slice(&(rgce.len() as u16).to_le_bytes()); // cce
        header.extend_from_slice(&0u16.to_le_bytes()); // ixals
        header.extend_from_slice(&0u16.to_le_bytes()); // itab
        header.extend_from_slice(&[0, 0, 0, 0]); // cchCustMenu, cchDescription, cchHelpTopic, cchStatusText

        let name_str = xl_unicode_string_no_cch_compressed(name);

        let stream = [
            record(records::RECORD_BOF_BIFF8, &[0u8; 16]),
            record(RECORD_NAME, &[header, name_str, rgce].concat()),
            record(records::RECORD_EOF, &[]),
        ]
        .concat();

        let allows_continuation = |id: u16| id == RECORD_NAME;
        let iter = records::LogicalBiffRecordIter::new(&stream, allows_continuation);
        let record = iter
            .filter_map(Result::ok)
            .find(|r| r.record_id == RECORD_NAME)
            .expect("NAME record");

        let err = parse_biff8_name_record(&record, 1252, &[]).unwrap_err();
        assert!(err.contains("PtgStr"), "err={err}");
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

    #[test]
    fn parses_defined_name_with_continued_richtext_description_string_crun_split_across_fragments() {
        let name = "DescRich";
        let rgce: Vec<u8> = vec![0x1E, 0x01, 0x00]; // PtgInt 1
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

        let name_str = xl_unicode_string_no_cch_compressed(name);

        // Rich-text description split so the u16 `cRun` spans the CONTINUE boundary.
        let mut desc_part1 = Vec::new();
        desc_part1.push(STR_FLAG_RICH_TEXT); // flags (compressed + rich text)
        desc_part1.push(0x01); // cRun low byte (cRun=1)

        let mut desc_part2 = Vec::new();
        desc_part2.push(0); // continued segment option flags (fHighByte=0)
        desc_part2.push(0x00); // cRun high byte
        desc_part2.extend_from_slice(description.as_bytes());
        desc_part2.extend_from_slice(&[0x11, 0x22, 0x33, 0x44]); // rgRun (4 bytes)

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

    #[test]
    fn parses_defined_name_with_continued_ext_description_string_ext_payload_split_across_fragments() {
        let name = "DescExt";
        let rgce: Vec<u8> = vec![0x1E, 0x01, 0x00]; // PtgInt 1

        let description = "abc";
        let status = "Z";
        let ext = [0xDEu8, 0xAD, 0xBE, 0xEF];

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
        header.push(status.len() as u8); // cchStatusText

        let name_str = xl_unicode_string_no_cch_compressed(name);

        // Description string: fExtSt set, and its ext payload is split across CONTINUE such that
        // the status text string follows in the continued fragment.
        let mut desc_part1 = Vec::new();
        desc_part1.push(STR_FLAG_EXT); // flags (compressed + ext)
        desc_part1.extend_from_slice(&(ext.len() as u32).to_le_bytes()); // cbExtRst
        desc_part1.extend_from_slice(description.as_bytes());
        desc_part1.extend_from_slice(&ext[..2]);

        let mut desc_part2 = Vec::new();
        desc_part2.push(0); // continued segment option flags (fHighByte=0)
        desc_part2.extend_from_slice(&ext[2..]);
        desc_part2.extend_from_slice(&xl_unicode_string_no_cch_compressed(status));

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

    #[test]
    fn strips_embedded_nuls_from_description_string() {
        let name = "NulDesc";
        let rgce: Vec<u8> = vec![0x1E, 0x01, 0x00]; // PtgInt 1
        let description = "Hello\0World";
        let expected = "HelloWorld";

        let mut header = Vec::new();
        header.extend_from_slice(&0u16.to_le_bytes()); // grbit
        header.push(0); // chKey
        header.push(name.len() as u8); // cch
        header.extend_from_slice(&(rgce.len() as u16).to_le_bytes()); // cce
        header.extend_from_slice(&0u16.to_le_bytes()); // ixals
        header.extend_from_slice(&0u16.to_le_bytes()); // itab
        header.push(0); // cchCustMenu
        header.push(description.len() as u8); // cchDescription (includes NUL)
        header.push(0); // cchHelpTopic
        header.push(0); // cchStatusText

        let name_str = xl_unicode_string_no_cch_compressed(name);
        let desc_str = xl_unicode_string_no_cch_compressed(description);

        let stream = [
            record(records::RECORD_BOF_BIFF8, &[0u8; 16]),
            record(
                RECORD_NAME,
                &[header, name_str, rgce.clone(), desc_str].concat(),
            ),
            record(records::RECORD_EOF, &[]),
        ]
        .concat();

        let parsed =
            parse_biff_defined_names(&stream, BiffVersion::Biff8, 1252, &[]).expect("parse names");
        assert_eq!(parsed.names.len(), 1);
        assert_eq!(parsed.names[0].name, name);
        assert_eq!(parsed.names[0].refers_to, "1");
        assert_eq!(parsed.names[0].comment.as_deref(), Some(expected));
        assert!(parsed.warnings.is_empty(), "warnings={:?}", parsed.warnings);
    }

    #[test]
    fn strips_embedded_nuls_from_defined_name_string() {
        let name = "Hello\0World";
        let expected = "HelloWorld";
        let rgce: Vec<u8> = vec![0x1E, 0x01, 0x00]; // PtgInt 1

        let mut header = Vec::new();
        header.extend_from_slice(&0u16.to_le_bytes()); // grbit
        header.push(0); // chKey
        header.push(name.len() as u8); // cch (includes NUL)
        header.extend_from_slice(&(rgce.len() as u16).to_le_bytes()); // cce
        header.extend_from_slice(&0u16.to_le_bytes()); // ixals
        header.extend_from_slice(&0u16.to_le_bytes()); // itab
        header.extend_from_slice(&[0, 0, 0, 0]); // cchCustMenu, cchDescription, cchHelpTopic, cchStatusText

        let name_str = xl_unicode_string_no_cch_compressed(name);

        let stream = [
            record(records::RECORD_BOF_BIFF8, &[0u8; 16]),
            record(RECORD_NAME, &[header, name_str, rgce.clone()].concat()),
            record(records::RECORD_EOF, &[]),
        ]
        .concat();

        let parsed =
            parse_biff_defined_names(&stream, BiffVersion::Biff8, 1252, &[]).expect("parse names");
        assert_eq!(parsed.names.len(), 1);
        assert_eq!(parsed.names[0].name, expected);
        assert_eq!(parsed.names[0].refers_to, "1");
        assert!(parsed.warnings.is_empty(), "warnings={:?}", parsed.warnings);
    }

    #[test]
    fn skips_malformed_name_record_but_continues_scan() {
        // First record: malformed/truncated description string.
        let bad_name = "BadDesc";
        let rgce: Vec<u8> = vec![0x1E, 0x01, 0x00]; // PtgInt 1

        let mut bad_header = Vec::new();
        bad_header.extend_from_slice(&0u16.to_le_bytes()); // grbit
        bad_header.push(0); // chKey
        bad_header.push(bad_name.len() as u8); // cch
        bad_header.extend_from_slice(&(rgce.len() as u16).to_le_bytes()); // cce
        bad_header.extend_from_slice(&0u16.to_le_bytes()); // ixals
        bad_header.extend_from_slice(&0u16.to_le_bytes()); // itab
        bad_header.push(0); // cchCustMenu
        bad_header.push(5); // cchDescription (claims 5 chars, but we truncate below)
        bad_header.push(0); // cchHelpTopic
        bad_header.push(0); // cchStatusText

        let bad_name_str = xl_unicode_string_no_cch_compressed(bad_name);
        // Truncated description: flags + only 2 bytes ("AB"), but header says 5 chars.
        let bad_desc_partial: Vec<u8> = [vec![0u8], b"AB".to_vec()].concat();

        let bad_record_payload =
            [bad_header, bad_name_str, rgce.clone(), bad_desc_partial].concat();

        // Second record: valid defined name.
        let good_name = "Good";
        let mut good_header = Vec::new();
        good_header.extend_from_slice(&0u16.to_le_bytes()); // grbit
        good_header.push(0); // chKey
        good_header.push(good_name.len() as u8); // cch
        good_header.extend_from_slice(&(rgce.len() as u16).to_le_bytes()); // cce
        good_header.extend_from_slice(&0u16.to_le_bytes()); // ixals
        good_header.extend_from_slice(&0u16.to_le_bytes()); // itab
        good_header.extend_from_slice(&[0, 0, 0, 0]); // cchCustMenu, cchDescription, cchHelpTopic, cchStatusText

        let good_name_str = xl_unicode_string_no_cch_compressed(good_name);
        let good_record_payload = [good_header, good_name_str, rgce.clone()].concat();

        let stream = [
            record(records::RECORD_BOF_BIFF8, &[0u8; 16]),
            record(RECORD_NAME, &bad_record_payload),
            record(RECORD_NAME, &good_record_payload),
            record(records::RECORD_EOF, &[]),
        ]
        .concat();

        let parsed =
            parse_biff_defined_names(&stream, BiffVersion::Biff8, 1252, &[]).expect("parse names");
        assert_eq!(parsed.names.len(), 1);
        assert_eq!(parsed.names[0].name, good_name);
        assert_eq!(parsed.names[0].refers_to, "1");
        assert_eq!(parsed.warnings.len(), 1);
        assert!(
            parsed.warnings[0].contains("failed to parse NAME record"),
            "warnings={:?}",
            parsed.warnings
        );
    }

    #[test]
    fn skips_name_record_when_continued_unicode_string_splits_mid_code_unit() {
        // First record: description string is BIFF8 unicode (fHighByte=1) but is split mid UTF-16LE
        // code unit across the CONTINUE boundary. The parser should not panic and should skip the
        // bad NAME record.
        let bad_name = "Bad";

        let mut bad_header = Vec::new();
        bad_header.extend_from_slice(&0u16.to_le_bytes()); // grbit
        bad_header.push(0); // chKey
        bad_header.push(bad_name.len() as u8); // cch
        bad_header.extend_from_slice(&0u16.to_le_bytes()); // cce (empty rgce)
        bad_header.extend_from_slice(&0u16.to_le_bytes()); // ixals
        bad_header.extend_from_slice(&0u16.to_le_bytes()); // itab
        bad_header.push(0); // cchCustMenu
        bad_header.push(1); // cchDescription (1 char)
        bad_header.push(0); // cchHelpTopic
        bad_header.push(0); // cchStatusText

        let bad_name_str = xl_unicode_string_no_cch_compressed(bad_name);

        // Description begins with flags byte (unicode) and then only 1 byte of the 2-byte code unit.
        let bad_desc_partial = [0x01u8, b'A'].to_vec();

        let name_payload = [bad_header, bad_name_str, bad_desc_partial].concat();
        // CONTINUE fragment contains the continued-segment option flags byte (unicode), then the
        // missing second byte of the UTF-16LE code unit.
        let cont_payload = vec![0x01u8, 0x00u8];

        // Second record: valid defined name.
        let good_name = "Good";
        let rgce: Vec<u8> = vec![0x1E, 0x01, 0x00]; // PtgInt 1
        let mut good_header = Vec::new();
        good_header.extend_from_slice(&0u16.to_le_bytes()); // grbit
        good_header.push(0); // chKey
        good_header.push(good_name.len() as u8); // cch
        good_header.extend_from_slice(&(rgce.len() as u16).to_le_bytes()); // cce
        good_header.extend_from_slice(&0u16.to_le_bytes()); // ixals
        good_header.extend_from_slice(&0u16.to_le_bytes()); // itab
        good_header.extend_from_slice(&[0, 0, 0, 0]); // no optional strings
        let good_name_str = xl_unicode_string_no_cch_compressed(good_name);
        let good_payload = [good_header, good_name_str, rgce].concat();

        let stream = [
            record(records::RECORD_BOF_BIFF8, &[0u8; 16]),
            record(RECORD_NAME, &name_payload),
            record(records::RECORD_CONTINUE, &cont_payload),
            record(RECORD_NAME, &good_payload),
            record(records::RECORD_EOF, &[]),
        ]
        .concat();

        let parsed =
            parse_biff_defined_names(&stream, BiffVersion::Biff8, 1252, &[]).expect("parse names");
        assert_eq!(parsed.names.len(), 1);
        assert_eq!(parsed.names[0].name, good_name);
        assert!(
            parsed
                .warnings
                .iter()
                .any(|w| w.contains("string continuation split mid-character")),
            "warnings={:?}",
            parsed.warnings
        );
    }

    #[test]
    fn preserves_ptgname_indices_when_name_record_is_skipped() {
        // Two NAME records:
        //  - First is malformed (truncated description string) so it is skipped, but it still
        //    occupies name_id=1.
        //  - Second refers to name_id=1 via PtgName. The decoder should not mis-resolve that to the
        //    second name (index-shift bug); it should render as #NAME? via the placeholder meta.
        let bad_name = "BadDesc";
        let good_name = "RefBad";

        // NAME #1 formula: `1` (not important; record will be skipped).
        let bad_rgce: Vec<u8> = vec![0x1E, 0x01, 0x00]; // PtgInt 1

        let mut bad_header = Vec::new();
        bad_header.extend_from_slice(&0u16.to_le_bytes()); // grbit
        bad_header.push(0); // chKey
        bad_header.push(bad_name.len() as u8); // cch
        bad_header.extend_from_slice(&(bad_rgce.len() as u16).to_le_bytes()); // cce
        bad_header.extend_from_slice(&0u16.to_le_bytes()); // ixals
        bad_header.extend_from_slice(&0u16.to_le_bytes()); // itab
        bad_header.push(0); // cchCustMenu
        bad_header.push(5); // cchDescription (claims 5 chars, but we truncate below)
        bad_header.push(0); // cchHelpTopic
        bad_header.push(0); // cchStatusText

        let bad_name_str = xl_unicode_string_no_cch_compressed(bad_name);
        // Truncated description: flags + only 2 bytes ("AB"), but header says 5 chars.
        let bad_desc_partial: Vec<u8> = [vec![0u8], b"AB".to_vec()].concat();

        let bad_record_payload =
            [bad_header, bad_name_str, bad_rgce.clone(), bad_desc_partial].concat();

        // NAME #2 formula: PtgName(name_id=1).
        let ptgname_rgce: Vec<u8> = [
            vec![0x23],                  // PtgName
            1u32.to_le_bytes().to_vec(), // name_id=1
            0u16.to_le_bytes().to_vec(), // reserved
        ]
        .concat();

        let mut good_header = Vec::new();
        good_header.extend_from_slice(&0u16.to_le_bytes()); // grbit
        good_header.push(0); // chKey
        good_header.push(good_name.len() as u8); // cch
        good_header.extend_from_slice(&(ptgname_rgce.len() as u16).to_le_bytes()); // cce
        good_header.extend_from_slice(&0u16.to_le_bytes()); // ixals
        good_header.extend_from_slice(&0u16.to_le_bytes()); // itab
        good_header.extend_from_slice(&[0, 0, 0, 0]); // cchCustMenu, cchDescription, cchHelpTopic, cchStatusText

        let good_name_str = xl_unicode_string_no_cch_compressed(good_name);
        let good_record_payload = [good_header, good_name_str, ptgname_rgce].concat();

        let stream = [
            record(records::RECORD_BOF_BIFF8, &[0u8; 16]),
            record(RECORD_NAME, &bad_record_payload),
            record(RECORD_NAME, &good_record_payload),
            record(records::RECORD_EOF, &[]),
        ]
        .concat();

        let parsed =
            parse_biff_defined_names(&stream, BiffVersion::Biff8, 1252, &[]).expect("parse names");
        assert_eq!(parsed.names.len(), 1);
        assert_eq!(parsed.names[0].name, good_name);
        assert_eq!(parsed.names[0].refers_to, "#NAME?");
        assert_eq!(parsed.warnings.len(), 1, "warnings={:?}", parsed.warnings);
        assert!(parsed.warnings[0].contains("failed to parse NAME record"));
    }

    #[test]
    fn parses_name_metadata_and_builtin_name_ids() {
        let rgce: Vec<u8> = vec![0x1E, 0x01, 0x00]; // PtgInt 1
        let description = "My description";

        // Print_Area built-in, hidden, worksheet-scoped.
        let grbit = NAME_FLAG_HIDDEN | NAME_FLAG_BUILTIN;
        let builtin_id = 0x06u8;
        let itab = 3u16;

        let mut header = Vec::new();
        header.extend_from_slice(&grbit.to_le_bytes());
        header.push(0); // chKey (keyboard shortcut; built-in id is stored in rgchName)
        header.push(1); // cch (built-in name id length)
        header.extend_from_slice(&(rgce.len() as u16).to_le_bytes()); // cce
        header.extend_from_slice(&0u16.to_le_bytes()); // ixals
        header.extend_from_slice(&itab.to_le_bytes()); // itab
        header.extend_from_slice(&[
            0,                       // cchCustMenu
            description.len() as u8, // cchDescription
            0,                       // cchHelpTopic
            0,                       // cchStatusText
        ]);

        let r_name = record(
            RECORD_NAME,
            &[
                header,
                vec![builtin_id], // rgchName (built-in id)
                rgce.clone(),
                xl_unicode_string_no_cch_compressed(description),
            ]
            .concat(),
        );

        // Additional built-ins to validate the mapping table.
        fn builtin_record(id: u8) -> Vec<u8> {
            let mut header = Vec::new();
            header.extend_from_slice(&NAME_FLAG_BUILTIN.to_le_bytes());
            header.push(0); // chKey
            header.push(1); // cch (built-in id length)
            header.extend_from_slice(&0u16.to_le_bytes()); // cce (empty formula)
            header.extend_from_slice(&0u16.to_le_bytes()); // ixals
            header.extend_from_slice(&0u16.to_le_bytes()); // itab
            header.extend_from_slice(&[0, 0, 0, 0]); // no optional strings

            record(RECORD_NAME, &[header, vec![id]].concat())
        }

        let stream = [
            record(records::RECORD_BOF_BIFF8, &[0u8; 16]),
            r_name,
            builtin_record(0x07), // Print_Titles
            builtin_record(0x0D), // _FilterDatabase
            builtin_record(0xFF), // unknown => placeholder
            record(records::RECORD_EOF, &[]),
        ]
        .concat();

        // Provide at least 3 sheets so itab=3 is in range.
        let sheet_names: Vec<String> = ["S1", "S2", "S3"].iter().map(|s| s.to_string()).collect();

        let parsed = parse_biff_defined_names(&stream, BiffVersion::Biff8, 1252, &sheet_names)
            .expect("parse names");
        assert_eq!(parsed.names.len(), 4);

        assert_eq!(parsed.names[0].name, formula_model::XLNM_PRINT_AREA);
        assert_eq!(parsed.names[0].builtin_id, Some(builtin_id));
        assert_eq!(parsed.names[0].itab, itab);
        assert_eq!(parsed.names[0].scope_sheet, Some(2));
        assert!(parsed.names[0].hidden);
        assert_eq!(parsed.names[0].comment.as_deref(), Some(description));
        assert_eq!(parsed.names[0].rgce, rgce);
        assert_eq!(parsed.names[0].refers_to, "1");

        assert_eq!(parsed.names[1].name, formula_model::XLNM_PRINT_TITLES);
        assert_eq!(parsed.names[1].builtin_id, Some(0x07));
        assert_eq!(parsed.names[2].name, formula_model::XLNM_FILTER_DATABASE);
        assert_eq!(parsed.names[2].builtin_id, Some(0x0D));
        assert_eq!(parsed.names[3].name, "_xlnm.Builtin_0xFF");
        assert_eq!(parsed.names[3].builtin_id, Some(0xFF));

        assert!(parsed.warnings.is_empty(), "warnings={:?}", parsed.warnings);
    }

    #[test]
    fn maps_common_builtin_name_ids() {
        // Auto_Open (0x01) should map to its Excel-visible name.
        let id = 0x01u8;
        let expected = "_xlnm.Auto_Open";

        let mut header = Vec::new();
        header.extend_from_slice(&NAME_FLAG_BUILTIN.to_le_bytes());
        header.push(b'X'); // chKey (keyboard shortcut; should not affect built-in id)
        header.push(1); // cch (built-in id length)
        header.extend_from_slice(&0u16.to_le_bytes()); // cce (empty formula)
        header.extend_from_slice(&0u16.to_le_bytes()); // ixals
        header.extend_from_slice(&0u16.to_le_bytes()); // itab
        header.extend_from_slice(&[0, 0, 0, 0]); // no optional strings

        let stream = [
            record(records::RECORD_BOF_BIFF8, &[0u8; 16]),
            record(RECORD_NAME, &[header, vec![id]].concat()),
            record(records::RECORD_EOF, &[]),
        ]
        .concat();

        let parsed =
            parse_biff_defined_names(&stream, BiffVersion::Biff8, 1252, &[]).expect("parse names");
        assert_eq!(parsed.names.len(), 1);
        assert_eq!(parsed.names[0].name, expected);
        assert_eq!(parsed.names[0].builtin_id, Some(id));
        assert_eq!(parsed.names[0].refers_to, "");
        assert!(parsed.warnings.is_empty(), "warnings={:?}", parsed.warnings);
    }

    #[test]
    fn builtin_name_prefers_rgchname_when_chkey_is_nonzero() {
        // Some producers store mismatched built-in ids in `chKey` vs `rgchName`.
        // Empirically Excel prefers the built-in id stored in `rgchName` and treats `chKey` as a
        // shortcut; ensure we do the same.
        let rgce: Vec<u8> = vec![0x1E, 0x01, 0x00]; // PtgInt 1

        let grbit = NAME_FLAG_BUILTIN;
        let ch_key = 0x06u8; // Print_Area
        let rgch_builtin_id = 0x07u8; // Print_Titles (mismatch)

        let mut header = Vec::new();
        header.extend_from_slice(&grbit.to_le_bytes());
        header.push(ch_key); // chKey
        header.push(1); // cch (built-in id length)
        header.extend_from_slice(&(rgce.len() as u16).to_le_bytes()); // cce
        header.extend_from_slice(&0u16.to_le_bytes()); // ixals
        header.extend_from_slice(&0u16.to_le_bytes()); // itab (workbook scope)
        header.extend_from_slice(&[0, 0, 0, 0]); // no optional strings

        let stream = [
            record(records::RECORD_BOF_BIFF8, &[0u8; 16]),
            record(
                RECORD_NAME,
                &[header, vec![rgch_builtin_id], rgce.clone()].concat(),
            ),
            record(records::RECORD_EOF, &[]),
        ]
        .concat();

        let parsed =
            parse_biff_defined_names(&stream, BiffVersion::Biff8, 1252, &[]).expect("parse names");
        assert_eq!(parsed.names.len(), 1);
        assert_eq!(parsed.names[0].name, formula_model::XLNM_PRINT_TITLES);
        assert_eq!(parsed.names[0].builtin_id, Some(rgch_builtin_id));
        assert_eq!(parsed.names[0].itab, 0);
        assert_eq!(parsed.names[0].scope_sheet, None);
        assert_eq!(parsed.names[0].rgce, rgce);
        assert_eq!(parsed.names[0].refers_to, "1");
        assert!(parsed.warnings.is_empty(), "warnings={:?}", parsed.warnings);
    }

    #[test]
    fn builtin_name_falls_back_to_chkey_when_rgchname_missing() {
        // Malformed files in the wild may set `fBuiltin` but omit the single-byte built-in id that
        // should be stored in `rgchName` (by setting `cch=0`). In that case we fall back to `chKey`
        // as a best-effort source of the id so the name can still be recognized.
        let rgce: Vec<u8> = vec![0x1E, 0x01, 0x00]; // PtgInt 1

        let grbit = NAME_FLAG_BUILTIN;
        let ch_key = 0x06u8; // Print_Area
        let cch = 0u8; // missing built-in id in rgchName

        let mut header = Vec::new();
        header.extend_from_slice(&grbit.to_le_bytes());
        header.push(ch_key); // chKey
        header.push(cch); // cch
        header.extend_from_slice(&(rgce.len() as u16).to_le_bytes()); // cce
        header.extend_from_slice(&0u16.to_le_bytes()); // ixals
        header.extend_from_slice(&0u16.to_le_bytes()); // itab (workbook scope)
        header.extend_from_slice(&[0, 0, 0, 0]); // no optional strings

        let stream = [
            record(records::RECORD_BOF_BIFF8, &[0u8; 16]),
            record(RECORD_NAME, &[header, rgce.clone()].concat()),
            record(records::RECORD_EOF, &[]),
        ]
        .concat();

        let parsed =
            parse_biff_defined_names(&stream, BiffVersion::Biff8, 1252, &[]).expect("parse names");
        assert_eq!(parsed.names.len(), 1);
        assert_eq!(parsed.names[0].name, formula_model::XLNM_PRINT_AREA);
        assert_eq!(parsed.names[0].builtin_id, Some(ch_key));
        assert_eq!(parsed.names[0].itab, 0);
        assert_eq!(parsed.names[0].scope_sheet, None);
        assert_eq!(parsed.names[0].rgce, rgce);
        assert_eq!(parsed.names[0].refers_to, "1");
        assert!(parsed.warnings.is_empty(), "warnings={:?}", parsed.warnings);
    }

    #[test]
    fn imports_defined_name_array_constant_from_rgcb() {
        // NAME formula `={1,2;3,4}` encoded with PtgArray + trailing rgcb bytes.
        let name = "ArrConst";

        let rgce: Vec<u8> = vec![
            0x20, // PtgArray
            0, 0, 0, 0, 0, 0, 0, // 7-byte opaque header
        ];

        // BIFF8 array constant: [cols_minus1: u16][rows_minus1: u16] + values row-major.
        let mut rgcb = Vec::<u8>::new();
        rgcb.extend_from_slice(&1u16.to_le_bytes()); // 2 cols
        rgcb.extend_from_slice(&1u16.to_le_bytes()); // 2 rows
        for n in [1.0f64, 2.0, 3.0, 4.0] {
            rgcb.push(0x01); // number
            rgcb.extend_from_slice(&n.to_le_bytes());
        }

        let mut header = Vec::new();
        header.extend_from_slice(&0u16.to_le_bytes()); // grbit
        header.push(0); // chKey
        header.push(name.len() as u8); // cch
        header.extend_from_slice(&(rgce.len() as u16).to_le_bytes()); // cce
        header.extend_from_slice(&0u16.to_le_bytes()); // ixals
        header.extend_from_slice(&0u16.to_le_bytes()); // itab (workbook scope)
        header.extend_from_slice(&[0, 0, 0, 0]); // no optional strings

        let name_str = xl_unicode_string_no_cch_compressed(name);

        let stream = [
            record(records::RECORD_BOF_BIFF8, &[0u8; 16]),
            record(RECORD_NAME, &[header, name_str, rgce, rgcb].concat()),
            record(records::RECORD_EOF, &[]),
        ]
        .concat();

        let parsed =
            parse_biff_defined_names(&stream, BiffVersion::Biff8, 1252, &[]).expect("parse names");
        assert_eq!(parsed.names.len(), 1);
        assert_eq!(parsed.names[0].name, name);
        assert_eq!(parsed.names[0].refers_to, "{1,2;3,4}");
        assert!(parsed.warnings.is_empty(), "warnings={:?}", parsed.warnings);
    }

    #[test]
    fn imports_defined_name_array_constant_with_description_after_rgcb() {
        // Ensure `rgcb` bytes do not interfere with parsing the optional description string.
        let name = "ArrDesc";
        let desc = "Hi";

        let rgce: Vec<u8> = vec![
            0x20, // PtgArray
            0, 0, 0, 0, 0, 0, 0, // 7-byte opaque header
        ];

        // {1,2}
        let mut rgcb = Vec::<u8>::new();
        rgcb.extend_from_slice(&1u16.to_le_bytes()); // 2 cols
        rgcb.extend_from_slice(&0u16.to_le_bytes()); // 1 row
        for n in [1.0f64, 2.0] {
            rgcb.push(0x01); // number
            rgcb.extend_from_slice(&n.to_le_bytes());
        }

        let mut header = Vec::new();
        header.extend_from_slice(&0u16.to_le_bytes()); // grbit
        header.push(0); // chKey
        header.push(name.len() as u8); // cch
        header.extend_from_slice(&(rgce.len() as u16).to_le_bytes()); // cce
        header.extend_from_slice(&0u16.to_le_bytes()); // ixals
        header.extend_from_slice(&0u16.to_le_bytes()); // itab (workbook scope)
        header.extend_from_slice(&[0, desc.len() as u8, 0, 0]); // cchCustMenu, cchDescription, cchHelpTopic, cchStatusText

        let name_str = xl_unicode_string_no_cch_compressed(name);
        let desc_str = xl_unicode_string_no_cch_compressed(desc);

        let stream = [
            record(records::RECORD_BOF_BIFF8, &[0u8; 16]),
            record(RECORD_NAME, &[header, name_str, rgce, rgcb, desc_str].concat()),
            record(records::RECORD_EOF, &[]),
        ]
        .concat();

        let parsed =
            parse_biff_defined_names(&stream, BiffVersion::Biff8, 1252, &[]).expect("parse names");
        assert_eq!(parsed.names.len(), 1);
        assert_eq!(parsed.names[0].name, name);
        assert_eq!(parsed.names[0].refers_to, "{1,2}");
        assert_eq!(parsed.names[0].comment.as_deref(), Some(desc));
        assert!(parsed.warnings.is_empty(), "warnings={:?}", parsed.warnings);
    }

    #[test]
    fn caps_defined_name_warnings_and_emits_suppression_message() {
        // Build a workbook globals stream with many malformed NAME records (empty payloads).
        // Each record should trigger a warning, but the warning vector must remain bounded.
        let count = MAX_DEFINED_NAME_WARNINGS + 50;

        let mut stream_parts: Vec<Vec<u8>> = Vec::new();
        stream_parts.push(record(records::RECORD_BOF_BIFF8, &[0u8; 16]));
        for _ in 0..count {
            stream_parts.push(record(RECORD_NAME, &[]));
        }
        stream_parts.push(record(records::RECORD_EOF, &[]));
        let stream = stream_parts.concat();

        let parsed =
            parse_biff_defined_names(&stream, BiffVersion::Biff8, 1252, &[]).expect("parse names");

        assert!(parsed.names.is_empty());

        assert_eq!(parsed.warnings.len(), MAX_DEFINED_NAME_WARNINGS + 1);
        assert_eq!(
            parsed.warnings[MAX_DEFINED_NAME_WARNINGS],
            "additional defined-name warnings suppressed"
        );

        assert_eq!(
            parsed
                .warnings
                .iter()
                .filter(|w| w.contains("failed to parse NAME record"))
                .count(),
            MAX_DEFINED_NAME_WARNINGS,
            "warnings={:?}",
            parsed.warnings
        );
    }
}
