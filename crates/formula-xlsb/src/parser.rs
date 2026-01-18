use std::collections::HashMap;
use std::io::{self, BufReader, Read};
use std::ops::ControlFlow;

use crate::biff12_varint;
use crate::shared_strings::SharedString;
use crate::workbook_context::{
    display_supbook_name, ExternName, ExternSheet, SupBook, SupBookKind, WorkbookContext,
};
use formula_model::rich_text::{RichText, RichTextRunStyle};
use thiserror::Error;

use crate::rgce::DecodeWarning;
use crate::strings::{
    read_xl_wide_string_with_flags, FlagsWidth, ParsedXlsbString,
};

#[cfg(test)]
use std::cell::Cell as ThreadCell;

#[cfg(test)]
thread_local! {
    static FORMULA_DECODE_ATTEMPTS: ThreadCell<usize> = ThreadCell::new(0);
}

#[cfg(test)]
fn reset_formula_decode_attempts() {
    FORMULA_DECODE_ATTEMPTS.with(|c| c.set(0));
}

#[cfg(test)]
fn formula_decode_attempts() -> usize {
    FORMULA_DECODE_ATTEMPTS.with(|c| c.get())
}

fn decode_formula_text(
    rgce: &[u8],
    rgcb: &[u8],
    ctx: &WorkbookContext,
    base: crate::rgce::CellCoord,
) -> crate::rgce::DecodedFormula {
    // Formula decoding is best-effort and can be expensive. Avoid multi-pass fallback decoding
    // here because large worksheets can contain millions of formula cells.
    #[cfg(test)]
    FORMULA_DECODE_ATTEMPTS.with(|c| c.set(c.get().saturating_add(1)));
    crate::rgce::decode_formula_rgce_with_context_and_rgcb_and_base(rgce, rgcb, ctx, base)
}

// Record IDs (BIFF12 / MS-XLSB). Values taken from pyxlsb (public domain-ish) and MS-XLSB.
#[allow(dead_code)]
pub(crate) mod biff12 {
    pub const WB_PROP: u32 = 0x0099;
    pub const CALC_PROP: u32 = 0x009A;

    // Workbook defined names (named ranges / constants / formulas).
    pub const NAME: u32 = 0x0027;

    pub const SHEETS_END: u32 = 0x0090;
    pub const SHEET: u32 = 0x009C;

    pub const WORKSHEET: u32 = 0x0081;
    pub const WORKSHEET_END: u32 = 0x0082;
    pub const SHEETDATA: u32 = 0x0091;
    pub const SHEETDATA_END: u32 = 0x0092;
    pub const DIMENSION: u32 = 0x0094;

    pub const ROW: u32 = 0x0000;
    pub const BLANK: u32 = 0x0001;
    pub const NUM: u32 = 0x0002;
    pub const BOOLERR: u32 = 0x0003;
    pub const BOOL: u32 = 0x0004;
    pub const FLOAT: u32 = 0x0005;
    pub const CELL_ST: u32 = 0x0006;
    pub const STRING: u32 = 0x0007;
    pub const FORMULA_STRING: u32 = 0x0008;
    pub const FORMULA_FLOAT: u32 = 0x0009;
    pub const FORMULA_BOOL: u32 = 0x000A;
    pub const FORMULA_BOOLERR: u32 = 0x000B;

    // Shared formula definition (MS-XLSB BrtShrFmla).
    pub const SHR_FMLA: u32 = 0x0010;

    pub const SST: u32 = 0x009F;
    pub const SST_END: u32 = 0x00A0;
    pub const SI: u32 = 0x0013;
}

#[derive(Debug, Error)]
pub enum Error {
    #[error("I/O error: {0}")]
    Io(#[from] io::Error),
    #[error("ZIP error: {0}")]
    Zip(#[from] zip::result::ZipError),
    #[error("XML error: {0}")]
    Xml(#[from] quick_xml::Error),
    #[error("allocation failure: {0}")]
    AllocationFailure(&'static str),
    #[error("invalid password")]
    InvalidPassword,
    #[error("invalid XLSB: unexpected end of record")]
    UnexpectedEof,
    #[error("invalid UTF-16 string in record")]
    InvalidUtf16,
    #[error("invalid worksheet name: {0}")]
    InvalidSheetName(String),
    #[error("sheet index out of bounds: {0}")]
    SheetIndexOutOfBounds(usize),
    #[error("missing relationship target for sheet rId {0}")]
    MissingSheetRelationship(String),
    #[error("invalid Excel string literal: {0}")]
    InvalidExcelStringLiteral(String),
    #[error("unsupported formula text: {0}")]
    UnsupportedFormulaText(String),
    #[error("XLSB part too large: {part} size {size} bytes exceeds limit {max} bytes")]
    PartTooLarge { part: String, size: u64, max: u64 },
    #[error("XLSB preserved parts too large: total {total} bytes exceeds limit {max} bytes")]
    PreservedPartsTooLarge { total: u64, max: u64 },
    #[error("XLSB has too many ZIP entries: {count} exceeds limit {max}")]
    TooManyZipEntries { count: usize, max: usize },
    #[error("office encryption error: {0}")]
    OfficeCrypto(String),
}

#[derive(Debug, Clone, Copy, PartialEq, Eq)]
pub enum CalcMode {
    Auto,
    Manual,
    AutoExceptTables,
}

#[derive(Debug, Clone, PartialEq, Eq)]
pub struct WorkbookProperties {
    pub date_system_1904: bool,
    pub calc_mode: Option<CalcMode>,
    pub full_calc_on_load: Option<bool>,
}

impl Default for WorkbookProperties {
    fn default() -> Self {
        Self {
            date_system_1904: false,
            calc_mode: None,
            full_calc_on_load: None,
        }
    }
}

#[derive(Debug, Clone, Copy, PartialEq, Eq)]
pub enum SheetVisibility {
    Visible,
    Hidden,
    VeryHidden,
}

impl Default for SheetVisibility {
    fn default() -> Self {
        Self::Visible
    }
}

#[derive(Debug, Clone, PartialEq, Eq)]
pub struct SheetMeta {
    pub name: String,
    pub part_path: String,
    pub visibility: SheetVisibility,
}

#[derive(Debug, Clone, PartialEq)]
pub struct DefinedName {
    /// 1-based name index used by `PtgName` tokens.
    pub index: u32,
    pub name: String,
    /// `None` for workbook-scope, `Some(sheet_index)` for sheet-scope (0-based).
    pub scope_sheet: Option<u32>,
    pub hidden: bool,
    /// Definition formula payload (best-effort).
    pub formula: Option<Formula>,
    /// Optional comment (not commonly present in XLSB; best-effort).
    pub comment: Option<String>,
}

#[derive(Debug, Clone, PartialEq)]
pub struct Dimension {
    pub start_row: u32,
    pub start_col: u32,
    pub height: u32,
    pub width: u32,
}

#[derive(Debug, Clone, PartialEq)]
pub enum CellValue {
    Blank,
    Number(f64),
    Bool(bool),
    /// Raw error code (Excel internal).
    Error(u8),
    Text(String),
}

#[derive(Debug, Clone, PartialEq)]
pub struct Formula {
    /// Raw formula token stream (`rgce`) from the XLSB record.
    pub rgce: Vec<u8>,
    /// Best-effort decoded Excel formula text (without leading `=`).
    pub text: Option<String>,
    /// Raw `grbitFmla` flags from the `BrtFmla*` record.
    ///
    /// We currently treat this as opaque and preserve it for round-trip
    /// fidelity. Excel uses it for semantics like "always recalc", shared
    /// formula markers, array indicators, etc.
    pub flags: u16,
    /// Any trailing bytes in the `BrtFmla*` record that we don't currently
    /// interpret but must preserve for round-tripping.
    pub extra: Vec<u8>,
    /// Non-fatal decode warnings encountered while decoding [`Self::rgce`].
    pub warnings: Vec<DecodeWarning>,
}

impl Formula {
    /// Construct a new formula payload with default XLSB flags (`0`) and no
    /// extra bytes.
    ///
    /// When *writing* new formulas we currently do not know how to populate
    /// `flags` for advanced Excel features (shared/array formulas, etc), so we
    /// default to `0` and rely on Excel to fill them in if needed.
    pub fn new(rgce: Vec<u8>, text: Option<String>) -> Self {
        Self {
            rgce,
            text,
            flags: 0,
            extra: Vec::new(),
            warnings: Vec::new(),
        }
    }
}

#[derive(Debug, Clone, PartialEq)]
pub struct Cell {
    pub row: u32,
    pub col: u32,
    pub style: u32,
    pub value: CellValue,
    pub formula: Option<Formula>,
    pub preserved_string: Option<ParsedXlsbString>,
}

#[derive(Debug, Clone, PartialEq)]
pub struct SheetData {
    pub dimension: Option<Dimension>,
    pub cells: Vec<Cell>,
}

pub(crate) struct Biff12Reader<R: Read> {
    inner: BufReader<R>,
}

pub(crate) struct Biff12Record<'a> {
    pub id: u32,
    pub data: &'a [u8],
}

impl<R: Read> Biff12Reader<R> {
    pub fn new(inner: R) -> Self {
        Self {
            inner: BufReader::new(inner),
        }
    }

    /// Chunk size used when reading BIFF record payloads into memory.
    ///
    /// BIFF record lengths are attacker-controlled and the underlying reader may be truncated
    /// (e.g. a size-limited wrapper around a ZIP entry). Avoid a single large `Vec::resize(len)`
    /// allocation up front; instead grow the buffer as bytes are successfully read.
    const READ_CHUNK_BYTES: usize = 64 * 1024; // 64 KiB

    pub fn read_record<'a>(
        &mut self,
        buf: &'a mut Vec<u8>,
    ) -> Result<Option<Biff12Record<'a>>, Error> {
        let Some(id) = self.read_id()? else {
            return Ok(None);
        };
        let Some(len) = self.read_len()? else {
            return Ok(None);
        };

        let len = len as usize;
        buf.clear();
        let mut remaining = len;
        while remaining > 0 {
            let chunk = remaining.min(Self::READ_CHUNK_BYTES);
            let start = buf.len();
            buf.resize(start + chunk, 0);
            if let Err(err) = self.inner.read_exact(&mut buf[start..]) {
                // Don't leave trailing zero bytes in `buf` when the stream ends early.
                buf.truncate(start);
                return Err(err.into());
            }
            remaining = remaining.saturating_sub(chunk);
        }
        Ok(Some(Biff12Record { id, data: buf }))
    }

    fn read_id(&mut self) -> Result<Option<u32>, Error> {
        Ok(biff12_varint::read_record_id(&mut self.inner)?)
    }

    fn read_len(&mut self) -> Result<Option<u32>, Error> {
        Ok(biff12_varint::read_record_len(&mut self.inner)?)
    }
}

pub(crate) struct RecordReader<'a> {
    data: &'a [u8],
    offset: usize,
}

impl<'a> RecordReader<'a> {
    pub(crate) fn new(data: &'a [u8]) -> Self {
        Self { data, offset: 0 }
    }

    pub(crate) fn skip(&mut self, n: usize) -> Result<(), Error> {
        self.offset = self
            .offset
            .checked_add(n)
            .filter(|&o| o <= self.data.len())
            .ok_or(Error::UnexpectedEof)?;
        Ok(())
    }

    pub(crate) fn read_u8(&mut self) -> Result<u8, Error> {
        let b = *self.data.get(self.offset).ok_or(Error::UnexpectedEof)?;
        self.offset += 1;
        Ok(b)
    }

    pub(crate) fn read_u16(&mut self) -> Result<u16, Error> {
        let raw = self
            .data
            .get(self.offset..self.offset + 2)
            .ok_or(Error::UnexpectedEof)?;
        self.offset += 2;
        Ok(u16::from_le_bytes([raw[0], raw[1]]))
    }

    pub(crate) fn read_u32(&mut self) -> Result<u32, Error> {
        let raw = self
            .data
            .get(self.offset..self.offset + 4)
            .ok_or(Error::UnexpectedEof)?;
        self.offset += 4;
        Ok(u32::from_le_bytes([raw[0], raw[1], raw[2], raw[3]]))
    }

    fn read_f64(&mut self) -> Result<f64, Error> {
        let raw = self
            .data
            .get(self.offset..self.offset + 8)
            .ok_or(Error::UnexpectedEof)?;
        self.offset += 8;
        Ok(f64::from_le_bytes([
            raw[0], raw[1], raw[2], raw[3], raw[4], raw[5], raw[6], raw[7],
        ]))
    }

    /// BIFF RK-encoded number used by `BrtCellRk` / `NUM` records.
    fn read_rk_number(&mut self) -> Result<f64, Error> {
        let raw = self.read_u32()? as i32;
        let mut v = if raw & 0x02 != 0 {
            // Signed integer.
            (raw >> 2) as f64
        } else {
            // 30-bit fraction shifted into IEEE754 double.
            let shifted = (raw as u32) & 0xFFFFFFFC;
            let bytes = (shifted as u64) << 32;
            f64::from_bits(bytes)
        };
        if raw & 0x01 != 0 {
            v /= 100.0;
        }
        Ok(v)
    }

    fn read_utf16_string(&mut self) -> Result<String, Error> {
        let len_chars = self.read_u32()? as usize;
        self.read_utf16_chars(len_chars)
    }

    pub(crate) fn read_utf16_chars(&mut self, len_chars: usize) -> Result<String, Error> {
        let byte_len = len_chars.checked_mul(2).ok_or(Error::UnexpectedEof)?;
        let raw = self
            .data
            .get(self.offset..self.offset + byte_len)
            .ok_or(Error::UnexpectedEof)?;
        self.offset += byte_len;

        // Avoid allocating a full `Vec<u16>` for attacker-controlled string lengths; decode
        // UTF-16LE directly into a `String`. This keeps peak memory closer to the final UTF-8
        // output size (the record payload bytes are already buffered elsewhere).
        let mut out = String::new();
        let _ = out.try_reserve(raw.len());
        let iter = raw
            .chunks_exact(2)
            .map(|chunk| u16::from_le_bytes([chunk[0], chunk[1]]));
        for decoded in std::char::decode_utf16(iter) {
            match decoded {
                Ok(ch) => out.push(ch),
                Err(_) => out.push('\u{FFFD}'),
            }
        }
        Ok(out)
    }

    pub(crate) fn read_slice(&mut self, len: usize) -> Result<&'a [u8], Error> {
        let raw = self
            .data
            .get(self.offset..self.offset + len)
            .ok_or(Error::UnexpectedEof)?;
        self.offset += len;
        Ok(raw)
    }
}

pub(crate) fn parse_workbook<R: Read>(
    workbook_bin: &mut R,
    rels: &HashMap<String, String>,
    decode_formulas: bool,
) -> Result<
    (
        Vec<SheetMeta>,
        WorkbookContext,
        WorkbookProperties,
        Vec<DefinedName>,
    ),
    Error,
> {
    let mut reader = Biff12Reader::new(workbook_bin);
    let mut buf = Vec::new();
    let mut sheets = Vec::new();
    let mut ctx = WorkbookContext::default();
    let mut props = WorkbookProperties::default();
    let mut defined_names: Vec<DefinedName> = Vec::new();
    let mut next_defined_name_index: u32 = 1;

    // NameX / external-name tables (used for add-ins and external defined names).
    let mut supbooks: Vec<SupBook> = Vec::new();
    let mut supbook_sheets: Vec<Vec<String>> = Vec::new();
    let mut namex_extern_names: HashMap<(u16, u16), ExternName> = HashMap::new();
    let mut namex_ixti_supbooks: HashMap<u16, u16> = HashMap::new();
    let mut extern_sheet_entries: Option<Vec<ExternSheet>> = None;

    let mut current_supbook: Option<u16> = None;
    let mut current_extern_name_idx: u16 = 0;
    while let Some(rec) = reader.read_record(&mut buf)? {
        match rec.id {
            biff12::WB_PROP => {
                // BrtWbProp: workbook properties flags. We only care about `date1904` for now.
                let mut rr = RecordReader::new(rec.data);
                let flags = rr.read_u32()?;
                props.date_system_1904 = (flags & 0x0000_0001) != 0;
            }
            biff12::CALC_PROP => {
                // BrtCalcProp: calculation settings. The exact spec has many knobs; for now we
                // interpret the first few bits, matching `calcPr` in XLSX.
                let mut rr = RecordReader::new(rec.data);
                let _calc_id = rr.read_u32()?;
                let flags = rr.read_u16()?;

                props.calc_mode = match flags & 0x0003 {
                    0 => Some(CalcMode::Manual),
                    1 => Some(CalcMode::Auto),
                    2 => Some(CalcMode::AutoExceptTables),
                    _ => None,
                };
                props.full_calc_on_load = Some((flags & 0x0004) != 0);
            }
            biff12::SHEET => {
                let mut rr = RecordReader::new(rec.data);
                let state_flags = rr.read_u32()?; // includes visibility state
                let visibility = match state_flags & 0x0003 {
                    1 => SheetVisibility::Hidden,
                    2 => SheetVisibility::VeryHidden,
                    _ => SheetVisibility::Visible,
                };
                let _sheet_id = rr.read_u32()?;
                let rel_id = rr.read_utf16_string()?;
                let name = rr.read_utf16_string()?;
                let Some(target) = rels.get(&rel_id) else {
                    return Err(Error::MissingSheetRelationship(rel_id));
                };
                let part_path = normalize_sheet_target(target);
                sheets.push(SheetMeta {
                    name,
                    part_path,
                    visibility,
                });
            }
            id if is_defined_name_record(id) => {
                let index = next_defined_name_index;
                next_defined_name_index = next_defined_name_index.saturating_add(1);
                if let Some(mut parsed) = parse_defined_name_record(rec.data) {
                    parsed.index = index;
                    defined_names.push(parsed);
                }
            }
            // External references.
            id if is_supbook_record(id) => {
                if let Some((supbook, sheets)) = parse_supbook(rec.data) {
                    supbooks.push(supbook);
                    supbook_sheets.push(sheets);
                    current_supbook = Some((supbooks.len() - 1) as u16);
                    current_extern_name_idx = 0;
                }
            }
            id if is_end_supbook_record(id) => {
                current_supbook = None;
                current_extern_name_idx = 0;
            }
            id if is_extern_name_record(id) => {
                let Some(supbook_index) = current_supbook else {
                    continue;
                };
                let Some(extern_name) = parse_extern_name(rec.data) else {
                    continue;
                };

                current_extern_name_idx = current_extern_name_idx.saturating_add(1);
                namex_extern_names.insert((supbook_index, current_extern_name_idx), extern_name);
            }
            id if is_extern_sheet_record(id) => {
                if let Some(entries) = parse_extern_sheet(rec.data) {
                    for (ixti, entry) in entries.iter().enumerate() {
                        namex_ixti_supbooks.insert(ixti as u16, entry.supbook_index);
                    }
                    extern_sheet_entries = Some(entries);
                }
            }
            // Keep scanning after the sheets list; external tables often appear later.
            biff12::SHEETS_END => {}
            _ => {
                // Heuristic fallback: some writers use different record ids for external-link tables
                // (or we haven't enumerated them yet). Try to recognize structures by shape.
                if namex_ixti_supbooks.is_empty() {
                    if let Some(entries) = parse_extern_sheet(rec.data) {
                        for (ixti, entry) in entries.iter().enumerate() {
                            namex_ixti_supbooks.insert(ixti as u16, entry.supbook_index);
                        }
                        extern_sheet_entries = Some(entries);
                    }
                }

                if current_supbook.is_none() {
                    if let Some((supbook, sheets)) = parse_supbook(rec.data) {
                        if supbook_is_plausible(&supbook) {
                            supbooks.push(supbook);
                            supbook_sheets.push(sheets);
                            current_supbook = Some((supbooks.len() - 1) as u16);
                            current_extern_name_idx = 0;
                        }
                    }
                } else if let Some(supbook_index) = current_supbook {
                    if let Some(extern_name) = parse_extern_name(rec.data) {
                        current_extern_name_idx = current_extern_name_idx.saturating_add(1);
                        namex_extern_names
                            .insert((supbook_index, current_extern_name_idx), extern_name);
                    }
                }
            }
        }
    }

    // Populate the ExternSheet table for 3D reference decoding/encoding when the ExternSheet
    // entries refer back into the current workbook.
    if let Some(entries) = &extern_sheet_entries {
        for (ixti, entry) in entries.iter().enumerate() {
            let supbook_kind = supbooks
                .get(entry.supbook_index as usize)
                .map(|sb| &sb.kind);
            let is_internal = matches!(supbook_kind, Some(SupBookKind::Internal))
                || (supbook_kind.is_none() && entry.supbook_index == 0);
            if is_internal {
                let Some(first_sheet) = sheets.get(entry.sheet_first as usize) else {
                    continue;
                };
                let Some(last_sheet) = sheets.get(entry.sheet_last as usize) else {
                    continue;
                };
                ctx.add_extern_sheet(&first_sheet.name, &last_sheet.name, ixti as u16);
                continue;
            }

            if matches!(supbook_kind, Some(SupBookKind::ExternalWorkbook)) {
                let Some(supbook) = supbooks.get(entry.supbook_index as usize) else {
                    continue;
                };
                let book = display_supbook_name(&supbook.raw_name);
                let sheet_list = supbook_sheets.get(entry.supbook_index as usize);
                let first_sheet = resolve_supbook_sheet_name(sheet_list, entry.sheet_first);
                let last_sheet = resolve_supbook_sheet_name(sheet_list, entry.sheet_last);
                ctx.add_extern_sheet_external_workbook(book, first_sheet, last_sheet, ixti as u16);
            }
        }
    }

    ctx.set_namex_tables(
        supbooks,
        supbook_sheets,
        namex_extern_names,
        namex_ixti_supbooks,
    );

    // Register defined names after we know the full sheet list so sheet-scoped names can be
    // displayed as `Sheet1!Name` in decoded formulas.
    for name in &defined_names {
        match name.scope_sheet {
            None => ctx.add_workbook_name(name.name.clone(), name.index),
            Some(sheet_idx) => {
                if let Some(sheet) = sheets.get(sheet_idx as usize) {
                    ctx.add_sheet_name(sheet.name.clone(), name.name.clone(), name.index);
                } else {
                    // Malformed scope; fall back to workbook scope so formulas can still decode.
                    ctx.add_workbook_name(name.name.clone(), name.index);
                }
            }
        }
    }

    if decode_formulas {
        // Best-effort decode of name definition formulas now that the workbook context is populated.
        for name in &mut defined_names {
            let Some(formula) = name.formula.as_mut() else {
                continue;
            };

            // Some contexts (notably defined-name formulas) can contain relative reference ptgs
            // (`PtgRefN` / `PtgAreaN`) which require a base cell to decode. We don't have a real
            // origin cell for workbook-scoped names, so use `A1` as a best-effort base.
            let base = crate::rgce::CellCoord::new(0, 0);
            let decoded = crate::rgce::decode_formula_rgce_with_context_and_rgcb_and_base(
                &formula.rgce,
                &formula.extra,
                &ctx,
                base,
            );

            formula.text = decoded.text;
            formula.warnings = decoded.warnings;
        }
    }

    Ok((sheets, ctx, props, defined_names))
}

#[cfg(test)]
mod tests {
    use super::*;
    use crate::rgce::{decode_rgce_with_context, encode_rgce_with_context, CellCoord};
    use std::collections::HashMap;
    use std::io::Cursor;

    fn write_record(out: &mut Vec<u8>, id: u32, payload: &[u8]) {
        biff12_varint::write_record_id(out, id).expect("write record id");
        biff12_varint::write_record_len(out, payload.len() as u32).expect("write record len");
        out.extend_from_slice(payload);
    }

    fn write_utf16_string(out: &mut Vec<u8>, s: &str) {
        let utf16le: Vec<u8> = s
            .encode_utf16()
            .flat_map(u16::to_le_bytes)
            .collect::<Vec<u8>>();
        let cch: u32 = (utf16le.len() / 2) as u32;
        out.extend_from_slice(&cch.to_le_bytes());
        out.extend_from_slice(&utf16le);
    }

    #[test]
    fn supbook_filename_heuristic_treats_template_extensions_as_external_workbooks() {
        // SupBook records in external-link tables often store just the referenced workbook file
        // name. Treat template/add-in extensions as external workbooks too so 3D references like
        // `'[Book.xltx]Sheet1'!A1` decode correctly.
        assert_eq!(
            classify_supbook_name("Book.xltx"),
            SupBookKind::ExternalWorkbook
        );
        assert_eq!(
            classify_supbook_name("Book.xltm"),
            SupBookKind::ExternalWorkbook
        );
        assert_eq!(
            classify_supbook_name("Book.xlt"),
            SupBookKind::ExternalWorkbook
        );
        assert_eq!(
            classify_supbook_name("Book.xla"),
            SupBookKind::ExternalWorkbook
        );
    }

    #[test]
    fn parse_workbook_populates_intern_extern_sheet_table() {
        // workbook.bin containing two sheets + an internal SupBook + ExternSheet table.
        let mut workbook_bin = Vec::new();

        // Sheet1 (rId1).
        let mut sheet1 = Vec::new();
        sheet1.extend_from_slice(&0u32.to_le_bytes()); // flags/state
        sheet1.extend_from_slice(&1u32.to_le_bytes()); // sheet id
        write_utf16_string(&mut sheet1, "rId1");
        write_utf16_string(&mut sheet1, "Sheet1");
        write_record(&mut workbook_bin, biff12::SHEET, &sheet1);

        // Sheet2 (rId2).
        let mut sheet2 = Vec::new();
        sheet2.extend_from_slice(&0u32.to_le_bytes()); // flags/state
        sheet2.extend_from_slice(&2u32.to_le_bytes()); // sheet id
        write_utf16_string(&mut sheet2, "rId2");
        write_utf16_string(&mut sheet2, "Sheet2");
        write_record(&mut workbook_bin, biff12::SHEET, &sheet2);

        // End of sheets list (we keep scanning for context records).
        write_record(&mut workbook_bin, biff12::SHEETS_END, &[]);

        // Internal SupBook (some producers store the first sheet name here rather than an empty string).
        let mut supbook = Vec::new();
        supbook.extend_from_slice(&2u16.to_le_bytes()); // ctab
        write_utf16_string(&mut supbook, "Sheet1");
        write_record(&mut workbook_bin, 0x00AE, &supbook);

        // ExternSheet table mapping ixti 0 -> Sheet1, ixti 1 -> Sheet2.
        let mut extern_sheet = Vec::new();
        extern_sheet.extend_from_slice(&2u16.to_le_bytes()); // cxti
                                                             // ixti 0
        extern_sheet.extend_from_slice(&0u16.to_le_bytes()); // supbook index
        extern_sheet.extend_from_slice(&0u16.to_le_bytes()); // sheet first
        extern_sheet.extend_from_slice(&0u16.to_le_bytes()); // sheet last
                                                             // ixti 1
        extern_sheet.extend_from_slice(&0u16.to_le_bytes()); // supbook index
        extern_sheet.extend_from_slice(&1u16.to_le_bytes()); // sheet first
        extern_sheet.extend_from_slice(&1u16.to_le_bytes()); // sheet last
        write_record(&mut workbook_bin, 0x0017, &extern_sheet);

        let rels: HashMap<String, String> = HashMap::from([
            ("rId1".to_string(), "worksheets/sheet1.bin".to_string()),
            ("rId2".to_string(), "worksheets/sheet2.bin".to_string()),
        ]);

        let (_sheets, ctx, _props, _defined_names) =
            parse_workbook(&mut Cursor::new(&workbook_bin), &rels, true)
                .expect("parse workbook.bin");

        assert_eq!(ctx.extern_sheet_index("Sheet1"), Some(0));
        assert_eq!(ctx.extern_sheet_index("Sheet2"), Some(1));

        let encoded =
            encode_rgce_with_context("=Sheet2!A1", &ctx, CellCoord::new(0, 0)).expect("encode");
        assert_eq!(
            encoded.rgce,
            vec![
                0x3A, // PtgRef3d
                0x01, 0x00, // ixti (Sheet2)
                0x00, 0x00, 0x00, 0x00, // row (A1)
                0x00, 0xC0, // col+flags (A, relative row/col)
            ]
        );

        let decoded = decode_rgce_with_context(&encoded.rgce, &ctx).expect("decode");
        assert_eq!(decoded, "Sheet2!A1");
    }

    #[test]
    fn parse_workbook_parses_brt_extern_sheet_even_when_table_is_already_populated() {
        // Some XLSB files contain multiple ExternSheet-like tables. Our workbook parser should
        // recognize the MS-XLSB `BrtExternSheet` record id (`0x016A`) even when we have already
        // populated the ExternSheet table from an earlier BIFF8-style record (`0x0017`).
        //
        // This test intentionally writes an incorrect BIFF8 ExternSheet mapping first (both `ixti`
        // entries point at Sheet1), followed by a `BrtExternSheet` that corrects `ixti=1` to point
        // at Sheet2. If we fail to parse `0x016A` after the table has been populated, Sheet2 will
        // remain unresolved.
        let mut workbook_bin = Vec::new();

        // Sheet1 (rId1).
        let mut sheet1 = Vec::new();
        sheet1.extend_from_slice(&0u32.to_le_bytes()); // flags/state
        sheet1.extend_from_slice(&1u32.to_le_bytes()); // sheet id
        write_utf16_string(&mut sheet1, "rId1");
        write_utf16_string(&mut sheet1, "Sheet1");
        write_record(&mut workbook_bin, biff12::SHEET, &sheet1);

        // Sheet2 (rId2).
        let mut sheet2 = Vec::new();
        sheet2.extend_from_slice(&0u32.to_le_bytes()); // flags/state
        sheet2.extend_from_slice(&2u32.to_le_bytes()); // sheet id
        write_utf16_string(&mut sheet2, "rId2");
        write_utf16_string(&mut sheet2, "Sheet2");
        write_record(&mut workbook_bin, biff12::SHEET, &sheet2);

        // End of sheets list (we keep scanning for context records).
        write_record(&mut workbook_bin, biff12::SHEETS_END, &[]);

        // Internal SupBook.
        let mut supbook = Vec::new();
        supbook.extend_from_slice(&2u16.to_le_bytes()); // ctab
        write_utf16_string(&mut supbook, "Sheet1");
        write_record(&mut workbook_bin, 0x00AE, &supbook);

        // BIFF8 ExternSheet table (incorrectly maps both ixti entries to Sheet1).
        let mut extern_sheet_biff8 = Vec::new();
        extern_sheet_biff8.extend_from_slice(&2u16.to_le_bytes()); // cxti
                                                                   // ixti 0: Sheet1
        extern_sheet_biff8.extend_from_slice(&0u16.to_le_bytes()); // supbook index
        extern_sheet_biff8.extend_from_slice(&0u16.to_le_bytes()); // sheet first
        extern_sheet_biff8.extend_from_slice(&0u16.to_le_bytes()); // sheet last
                                                                   // ixti 1: (wrongly) also Sheet1
        extern_sheet_biff8.extend_from_slice(&0u16.to_le_bytes()); // supbook index
        extern_sheet_biff8.extend_from_slice(&0u16.to_le_bytes()); // sheet first
        extern_sheet_biff8.extend_from_slice(&0u16.to_le_bytes()); // sheet last
        write_record(&mut workbook_bin, 0x0017, &extern_sheet_biff8);

        // MS-XLSB BrtExternSheet table (corrects ixti 1 -> Sheet2).
        let mut extern_sheet_brt = Vec::new();
        extern_sheet_brt.extend_from_slice(&2u32.to_le_bytes()); // cxti
                                                                 // ixti 0: Sheet1
        extern_sheet_brt.extend_from_slice(&0u32.to_le_bytes()); // supbook index
        extern_sheet_brt.extend_from_slice(&0u32.to_le_bytes()); // sheet first
        extern_sheet_brt.extend_from_slice(&0u32.to_le_bytes()); // sheet last
                                                                 // ixti 1: Sheet2
        extern_sheet_brt.extend_from_slice(&0u32.to_le_bytes()); // supbook index
        extern_sheet_brt.extend_from_slice(&1u32.to_le_bytes()); // sheet first
        extern_sheet_brt.extend_from_slice(&1u32.to_le_bytes()); // sheet last
        write_record(&mut workbook_bin, 0x016A, &extern_sheet_brt);

        let rels: HashMap<String, String> = HashMap::from([
            ("rId1".to_string(), "worksheets/sheet1.bin".to_string()),
            ("rId2".to_string(), "worksheets/sheet2.bin".to_string()),
        ]);

        let (_sheets, ctx, _props, _defined_names) =
            parse_workbook(&mut Cursor::new(&workbook_bin), &rels, true)
                .expect("parse workbook.bin");

        assert_eq!(ctx.extern_sheet_index("Sheet1"), Some(0));
        assert_eq!(ctx.extern_sheet_index("Sheet2"), Some(1));
    }

    #[test]
    fn normalize_sheet_target_handles_absolute_and_prefixed_paths() {
        assert_eq!(
            normalize_sheet_target("worksheets/sheet1.bin"),
            "xl/worksheets/sheet1.bin"
        );
        assert_eq!(
            normalize_sheet_target(r"worksheets\sheet1.bin"),
            "xl/worksheets/sheet1.bin"
        );
        assert_eq!(
            normalize_sheet_target("/xl/worksheets/sheet1.bin"),
            "xl/worksheets/sheet1.bin"
        );
        assert_eq!(
            normalize_sheet_target("xl/worksheets/sheet1.bin"),
            "xl/worksheets/sheet1.bin"
        );
    }

    #[test]
    fn parse_workbook_populates_defined_names() {
        let mut workbook_bin = Vec::new();

        // Sheet1 (rId1).
        let mut sheet1 = Vec::new();
        sheet1.extend_from_slice(&0u32.to_le_bytes()); // flags/state
        sheet1.extend_from_slice(&1u32.to_le_bytes()); // sheet id
        write_utf16_string(&mut sheet1, "rId1");
        write_utf16_string(&mut sheet1, "Sheet1");
        write_record(&mut workbook_bin, biff12::SHEET, &sheet1);

        // Sheet2 (rId2).
        let mut sheet2 = Vec::new();
        sheet2.extend_from_slice(&0u32.to_le_bytes()); // flags/state
        sheet2.extend_from_slice(&2u32.to_le_bytes()); // sheet id
        write_utf16_string(&mut sheet2, "rId2");
        write_utf16_string(&mut sheet2, "Sheet2");
        write_record(&mut workbook_bin, biff12::SHEET, &sheet2);

        write_record(&mut workbook_bin, biff12::SHEETS_END, &[]);

        // Defined names (record id 0x0018 / NAME).
        // Index 1: workbook-scope "MyName".
        let mut name1 = Vec::new();
        name1.extend_from_slice(&0u16.to_le_bytes()); // flags
        name1.extend_from_slice(&0xFFFFFFFFu32.to_le_bytes()); // workbook scope sentinel
        write_utf16_string(&mut name1, "MyName");
        write_record(&mut workbook_bin, 0x0018, &name1);

        // Index 2: sheet-scope "LocalName" on Sheet2 (0-based sheet index = 1).
        let mut name2 = Vec::new();
        name2.extend_from_slice(&0u16.to_le_bytes()); // flags
        name2.extend_from_slice(&1u32.to_le_bytes()); // sheet index
        write_utf16_string(&mut name2, "LocalName");
        write_record(&mut workbook_bin, 0x0018, &name2);

        let rels: HashMap<String, String> = HashMap::from([
            ("rId1".to_string(), "worksheets/sheet1.bin".to_string()),
            ("rId2".to_string(), "worksheets/sheet2.bin".to_string()),
        ]);

        let (_sheets, ctx, _props, _defined_names) =
            parse_workbook(&mut Cursor::new(&workbook_bin), &rels, true)
                .expect("parse workbook.bin");

        assert_eq!(ctx.name_index("MyName", None), Some(1));
        assert_eq!(ctx.name_index("LocalName", Some("Sheet2")), Some(2));
        assert_eq!(ctx.name_index("LocalName", Some("Sheet1")), None);
    }

    #[test]
    fn parse_workbook_populates_defined_names_brtname() {
        let mut workbook_bin = Vec::new();

        // Sheet1 (rId1).
        let mut sheet1 = Vec::new();
        sheet1.extend_from_slice(&0u32.to_le_bytes()); // flags/state
        sheet1.extend_from_slice(&1u32.to_le_bytes()); // sheet id
        write_utf16_string(&mut sheet1, "rId1");
        write_utf16_string(&mut sheet1, "Sheet1");
        write_record(&mut workbook_bin, biff12::SHEET, &sheet1);

        // Sheet2 (rId2).
        let mut sheet2 = Vec::new();
        sheet2.extend_from_slice(&0u32.to_le_bytes()); // flags/state
        sheet2.extend_from_slice(&2u32.to_le_bytes()); // sheet id
        write_utf16_string(&mut sheet2, "rId2");
        write_utf16_string(&mut sheet2, "Sheet2");
        write_record(&mut workbook_bin, biff12::SHEET, &sheet2);

        write_record(&mut workbook_bin, biff12::SHEETS_END, &[]);

        // Defined names using BIFF12 `BrtName` (record id `0x0027`).
        // Index 1: workbook-scope "MyName" with refersTo = 42.
        let mut name1 = Vec::new();
        name1.extend_from_slice(&0u32.to_le_bytes()); // flags
        name1.extend_from_slice(&0xFFFF_FFFFu32.to_le_bytes()); // workbook scope sentinel
        name1.push(0); // reserved
        write_utf16_string(&mut name1, "MyName");
        let rgce1: Vec<u8> = vec![0x1E, 42u16.to_le_bytes()[0], 42u16.to_le_bytes()[1]]; // PtgInt(42)
        name1.extend_from_slice(&(rgce1.len() as u32).to_le_bytes());
        name1.extend_from_slice(&rgce1);
        write_record(&mut workbook_bin, biff12::NAME, &name1);

        // Index 2: sheet-scope "LocalName" on Sheet2 with refersTo = 7.
        let mut name2 = Vec::new();
        name2.extend_from_slice(&0u32.to_le_bytes()); // flags
        name2.extend_from_slice(&1u32.to_le_bytes()); // 0-based sheet index (Sheet2)
        name2.push(0); // reserved
        write_utf16_string(&mut name2, "LocalName");
        let rgce2: Vec<u8> = vec![0x1E, 7u16.to_le_bytes()[0], 7u16.to_le_bytes()[1]]; // PtgInt(7)
        name2.extend_from_slice(&(rgce2.len() as u32).to_le_bytes());
        name2.extend_from_slice(&rgce2);
        write_record(&mut workbook_bin, biff12::NAME, &name2);

        let rels: HashMap<String, String> = HashMap::from([
            ("rId1".to_string(), "worksheets/sheet1.bin".to_string()),
            ("rId2".to_string(), "worksheets/sheet2.bin".to_string()),
        ]);

        let (_sheets, ctx, _props, defined_names) =
            parse_workbook(&mut Cursor::new(&workbook_bin), &rels, true)
                .expect("parse workbook.bin");

        assert_eq!(ctx.name_index("MyName", None), Some(1));
        assert_eq!(ctx.name_index("LocalName", Some("Sheet2")), Some(2));

        assert_eq!(defined_names.len(), 2);
        assert_eq!(defined_names[0].name, "MyName");
        assert_eq!(defined_names[0].scope_sheet, None);
        assert_eq!(
            defined_names[0]
                .formula
                .as_ref()
                .and_then(|f| f.text.as_deref()),
            Some("42")
        );

        assert_eq!(defined_names[1].name, "LocalName");
        assert_eq!(defined_names[1].scope_sheet, Some(1));
        assert_eq!(
            defined_names[1]
                .formula
                .as_ref()
                .and_then(|f| f.text.as_deref()),
            Some("7")
        );
    }

    #[test]
    fn parse_workbook_decodes_defined_name_with_relative_ptg_using_base_cell() {
        let mut workbook_bin = Vec::new();

        // Sheet1 (rId1).
        let mut sheet1 = Vec::new();
        sheet1.extend_from_slice(&0u32.to_le_bytes()); // flags/state
        sheet1.extend_from_slice(&1u32.to_le_bytes()); // sheet id
        write_utf16_string(&mut sheet1, "rId1");
        write_utf16_string(&mut sheet1, "Sheet1");
        write_record(&mut workbook_bin, biff12::SHEET, &sheet1);

        write_record(&mut workbook_bin, biff12::SHEETS_END, &[]);

        // Workbook-scoped defined name `RelName` with a relative reference token (PtgRefN).
        //
        // `PtgRefN` needs a base cell to decode; we use `A1` as a best-effort base for defined
        // name formulas.
        let mut name = Vec::new();
        name.extend_from_slice(&0u32.to_le_bytes()); // flags
        name.extend_from_slice(&0xFFFF_FFFFu32.to_le_bytes()); // workbook scope sentinel
        name.push(0); // reserved
        write_utf16_string(&mut name, "RelName");
        let rgce: Vec<u8> = vec![
            0x2C, // PtgRefN (ref class)
            0, 0, 0, 0, // row_off (i32)
            0, 0, // col_off (i16)
        ];
        name.extend_from_slice(&(rgce.len() as u32).to_le_bytes());
        name.extend_from_slice(&rgce);
        write_record(&mut workbook_bin, biff12::NAME, &name);

        let rels: HashMap<String, String> =
            HashMap::from([("rId1".to_string(), "worksheets/sheet1.bin".to_string())]);

        let (_sheets, _ctx, _props, defined_names) =
            parse_workbook(&mut Cursor::new(&workbook_bin), &rels, true)
                .expect("parse workbook.bin");

        assert_eq!(defined_names.len(), 1);
        assert_eq!(defined_names[0].name, "RelName");
        assert_eq!(
            defined_names[0]
                .formula
                .as_ref()
                .and_then(|f| f.text.as_deref()),
            Some("A1"),
        );
    }

    #[test]
    fn parse_workbook_defined_name_indices_skip_unparsed_records() {
        let mut workbook_bin = Vec::new();

        // Sheet1 (rId1) is needed so `parse_workbook` can resolve workbook relationships.
        let mut sheet1 = Vec::new();
        sheet1.extend_from_slice(&0u32.to_le_bytes()); // flags/state
        sheet1.extend_from_slice(&1u32.to_le_bytes()); // sheet id
        write_utf16_string(&mut sheet1, "rId1");
        write_utf16_string(&mut sheet1, "Sheet1");
        write_record(&mut workbook_bin, biff12::SHEET, &sheet1);

        write_record(&mut workbook_bin, biff12::SHEETS_END, &[]);

        // Defined name record that is too short to parse. This should still consume an index so
        // subsequent PtgName tokens remain aligned with Excel's name ids.
        write_record(&mut workbook_bin, 0x0018, &[0xFF]);

        // Next defined name record should be assigned index=2.
        let mut name2 = Vec::new();
        name2.extend_from_slice(&0u16.to_le_bytes()); // flags
        name2.extend_from_slice(&0xFFFFFFFFu32.to_le_bytes()); // workbook scope sentinel
        write_utf16_string(&mut name2, "GoodName");
        write_record(&mut workbook_bin, 0x0018, &name2);

        let rels: HashMap<String, String> =
            HashMap::from([("rId1".to_string(), "worksheets/sheet1.bin".to_string())]);

        let (_sheets, ctx, _props, defined_names) =
            parse_workbook(&mut Cursor::new(&workbook_bin), &rels, true)
                .expect("parse workbook.bin");

        // Only one name parsed, but it should still get the correct (skipped) index.
        assert_eq!(defined_names.len(), 1);
        assert_eq!(defined_names[0].name, "GoodName");
        assert_eq!(defined_names[0].index, 2);
        assert_eq!(ctx.name_index("GoodName", None), Some(2));
    }

    #[test]
    fn parse_workbook_populates_external_workbook_extern_sheet_targets() {
        // workbook.bin containing an external SupBook (with sheet names) + ExternSheet mapping.
        let mut workbook_bin = Vec::new();

        // End of sheets list (we keep scanning for context records).
        write_record(&mut workbook_bin, biff12::SHEETS_END, &[]);

        // External SupBook: ctab=2, raw_name is a path, followed by sheet name list.
        let mut supbook = Vec::new();
        supbook.extend_from_slice(&2u16.to_le_bytes()); // ctab
        write_utf16_string(&mut supbook, r"C:\tmp\Book2.xlsb");
        write_utf16_string(&mut supbook, "SheetA");
        write_utf16_string(&mut supbook, "SheetB");
        write_record(&mut workbook_bin, 0x00AE, &supbook);

        // ExternSheet entry referencing the external SupBook's sheet range.
        let mut extern_sheet = Vec::new();
        extern_sheet.extend_from_slice(&1u16.to_le_bytes()); // cxti
        extern_sheet.extend_from_slice(&0u16.to_le_bytes()); // supbook index
        extern_sheet.extend_from_slice(&0u16.to_le_bytes()); // sheet first
        extern_sheet.extend_from_slice(&1u16.to_le_bytes()); // sheet last
        write_record(&mut workbook_bin, 0x0017, &extern_sheet);

        let rels: HashMap<String, String> = HashMap::new();
        let (_sheets, ctx, _props, _defined_names) =
            parse_workbook(&mut Cursor::new(&workbook_bin), &rels, true)
                .expect("parse workbook.bin");

        let (workbook, first, last) = ctx.extern_sheet_target(0).expect("extern sheet target");
        assert_eq!(workbook, Some("Book2.xlsb"));
        assert_eq!(first, "SheetA");
        assert_eq!(last, "SheetB");
    }

    #[test]
    fn parses_shared_string_phonetic_tail_as_opaque_bytes() {
        // SI with phonetic flag set: [flags][xlWideString][phonetic bytes...]
        let mut payload = Vec::new();
        payload.push(0x02); // phonetic flag
        write_utf16_string(&mut payload, "Hi");
        payload.extend_from_slice(&[0xDE, 0xAD, 0xBE, 0xEF]); // opaque phonetic tail

        let si = parse_shared_string_item(&payload).expect("parse SI");
        assert_eq!(si.plain_text(), "Hi");
        assert_eq!(si.rich_text.text, "Hi");
        assert_eq!(si.rich_text.runs.len(), 0);
        assert_eq!(si.phonetic, Some(vec![0xDE, 0xAD, 0xBE, 0xEF]));
        assert!(si.raw_si.is_some());
    }

    #[test]
    fn parse_workbook_reads_date1904_and_sheet_visibility() {
        let mut workbook_bin = Vec::new();

        // BrtWbProp (date1904=true).
        let mut wb_prop = Vec::new();
        wb_prop.extend_from_slice(&0x0000_0001u32.to_le_bytes());
        write_record(&mut workbook_bin, biff12::WB_PROP, &wb_prop);

        // Hidden Sheet1 (rId1).
        let mut sheet = Vec::new();
        sheet.extend_from_slice(&1u32.to_le_bytes()); // hidden state flags
        sheet.extend_from_slice(&1u32.to_le_bytes()); // sheet id
        write_utf16_string(&mut sheet, "rId1");
        write_utf16_string(&mut sheet, "Sheet1");
        write_record(&mut workbook_bin, biff12::SHEET, &sheet);

        let rels: HashMap<String, String> =
            HashMap::from([("rId1".to_string(), "worksheets/sheet1.bin".to_string())]);

        let (sheets, _ctx, props, _defined_names) =
            parse_workbook(&mut Cursor::new(&workbook_bin), &rels, true)
                .expect("parse workbook.bin");

        assert!(props.date_system_1904);
        assert_eq!(sheets.len(), 1);
        assert_eq!(sheets[0].visibility, SheetVisibility::Hidden);
    }

    #[test]
    fn parses_shared_string_rich_runs_with_surrogate_pairs() {
        // Rich SI with a surrogate pair (😀) to validate UTF-16 -> char index mapping.
        let mut payload = Vec::new();
        payload.push(0x01); // rich text flag
        write_utf16_string(&mut payload, "Hi 😀Bold");

        // cRun = 3
        payload.extend_from_slice(&3u32.to_le_bytes());

        // Runs: [ich (u32 UTF-16 offset)][ifnt (u16)][reserved (u16)]
        // "Hi " starts at 0
        payload.extend_from_slice(&0u32.to_le_bytes());
        payload.extend_from_slice(&0u16.to_le_bytes());
        payload.extend_from_slice(&0u16.to_le_bytes());
        // "😀" starts after "Hi " (3 UTF-16 units)
        payload.extend_from_slice(&3u32.to_le_bytes());
        payload.extend_from_slice(&1u16.to_le_bytes());
        payload.extend_from_slice(&0u16.to_le_bytes());
        // "Bold" starts after emoji (5 UTF-16 units)
        payload.extend_from_slice(&5u32.to_le_bytes());
        payload.extend_from_slice(&2u16.to_le_bytes());
        payload.extend_from_slice(&0u16.to_le_bytes());

        let si = parse_shared_string_item(&payload).expect("parse SI");
        assert_eq!(si.plain_text(), "Hi 😀Bold");
        assert_eq!(si.rich_text.runs.len(), 3);
        assert_eq!(si.rich_text.slice_run_text(&si.rich_text.runs[0]), "Hi ");
        assert_eq!(si.rich_text.slice_run_text(&si.rich_text.runs[1]), "😀");
        assert_eq!(si.rich_text.slice_run_text(&si.rich_text.runs[2]), "Bold");

        assert_eq!(si.run_formats.len(), 3);
        assert_eq!(si.run_formats[0], vec![0, 0, 0, 0]);
        assert_eq!(si.run_formats[1], vec![1, 0, 0, 0]);
        assert_eq!(si.run_formats[2], vec![2, 0, 0, 0]);
    }

    #[test]
    fn parse_sheet_stream_preserves_shared_string_phonetic_bytes() {
        let ctx = WorkbookContext::default();
        let shared_strings = vec!["Base".to_string()];
        let phonetic_bytes = vec![0xDE, 0xAD, 0xBE, 0xEF];

        let shared_strings_table = vec![SharedString {
            rich_text: RichText::new("Base".to_string()),
            run_formats: Vec::new(),
            phonetic: Some(phonetic_bytes.clone()),
            raw_si: Some(Vec::new()),
        }];

        // Worksheet containing a single BrtCellIsst (`biff12::STRING`) cell referencing `isst=0`.
        let mut sheet_bin = Vec::new();
        write_record(&mut sheet_bin, biff12::SHEETDATA, &[]);

        // Row 0.
        let mut row = Vec::new();
        row.extend_from_slice(&0u32.to_le_bytes());
        write_record(&mut sheet_bin, biff12::ROW, &row);

        // Col 0: shared string index 0.
        let mut cell = Vec::new();
        cell.extend_from_slice(&0u32.to_le_bytes()); // col
        cell.extend_from_slice(&0u32.to_le_bytes()); // style
        cell.extend_from_slice(&0u32.to_le_bytes()); // isst
        write_record(&mut sheet_bin, biff12::STRING, &cell);

        write_record(&mut sheet_bin, biff12::SHEETDATA_END, &[]);

        let mut cells = Vec::new();
        parse_sheet_stream(
            &mut Cursor::new(&sheet_bin),
            &shared_strings,
            Some(&shared_strings_table),
            &ctx,
            false,
            false,
            |cell| {
                cells.push(cell);
                ControlFlow::Continue(())
            },
        )
        .expect("parse sheet");

        assert_eq!(cells.len(), 1);
        let cell = &cells[0];
        assert_eq!(cell.value, CellValue::Text("Base".to_string()));
        assert_eq!(
            cell.preserved_string
                .as_ref()
                .and_then(|s| s.phonetic.as_deref()),
            Some(phonetic_bytes.as_slice())
        );
    }

    #[test]
    fn parse_sheet_stream_decodes_each_formula_cell_once() {
        reset_formula_decode_attempts();

        let ctx = WorkbookContext::default();
        let shared_strings: Vec<String> = Vec::new();

        let mut sheet_bin = Vec::new();
        write_record(&mut sheet_bin, biff12::SHEETDATA, &[]);

        // Row 0.
        let mut row = Vec::new();
        row.extend_from_slice(&0u32.to_le_bytes());
        write_record(&mut sheet_bin, biff12::ROW, &row);

        // Col 0: =1+2 (cached value 3.0)
        let encoded = encode_rgce_with_context("=1+2", &ctx, CellCoord::new(0, 0)).expect("encode");
        let mut fmla_num = Vec::new();
        fmla_num.extend_from_slice(&0u32.to_le_bytes()); // col
        fmla_num.extend_from_slice(&0u32.to_le_bytes()); // style
        fmla_num.extend_from_slice(&3.0f64.to_le_bytes()); // cached value
        fmla_num.extend_from_slice(&0u16.to_le_bytes()); // grbitFmla
        fmla_num.extend_from_slice(&(encoded.rgce.len() as u32).to_le_bytes());
        fmla_num.extend_from_slice(&encoded.rgce);
        fmla_num.extend_from_slice(&encoded.rgcb);
        write_record(&mut sheet_bin, biff12::FORMULA_FLOAT, &fmla_num);

        // Col 1: =SUM({1,2}) (cached value 3.0, includes rgcb)
        let encoded =
            encode_rgce_with_context("=SUM({1,2})", &ctx, CellCoord::new(0, 1)).expect("encode");
        let mut fmla_num = Vec::new();
        fmla_num.extend_from_slice(&1u32.to_le_bytes()); // col
        fmla_num.extend_from_slice(&0u32.to_le_bytes()); // style
        fmla_num.extend_from_slice(&3.0f64.to_le_bytes()); // cached value
        fmla_num.extend_from_slice(&0u16.to_le_bytes()); // grbitFmla
        fmla_num.extend_from_slice(&(encoded.rgce.len() as u32).to_le_bytes());
        fmla_num.extend_from_slice(&encoded.rgce);
        fmla_num.extend_from_slice(&encoded.rgcb);
        write_record(&mut sheet_bin, biff12::FORMULA_FLOAT, &fmla_num);

        // Col 2: =TRUE (cached value 1). Encode directly since the formula parser treats TRUE as
        // an identifier rather than a boolean literal.
        let rgce = vec![
            0x1D, // PtgBool
            0x01, // TRUE
        ];
        let mut fmla_bool = Vec::new();
        fmla_bool.extend_from_slice(&2u32.to_le_bytes()); // col
        fmla_bool.extend_from_slice(&0u32.to_le_bytes()); // style
        fmla_bool.push(1); // cached value
        fmla_bool.extend_from_slice(&0u16.to_le_bytes()); // grbitFmla
        fmla_bool.extend_from_slice(&(rgce.len() as u32).to_le_bytes());
        fmla_bool.extend_from_slice(&rgce);
        write_record(&mut sheet_bin, biff12::FORMULA_BOOL, &fmla_bool);

        write_record(&mut sheet_bin, biff12::SHEETDATA_END, &[]);

        let mut cells = Vec::new();
        parse_sheet_stream(
            &mut Cursor::new(&sheet_bin),
            &shared_strings,
            None,
            &ctx,
            false,
            true,
            |cell| {
                cells.push(cell);
                ControlFlow::Continue(())
            },
        )
        .expect("parse sheet");

        assert_eq!(cells.len(), 3);
        assert_eq!(
            cells[0].formula.as_ref().and_then(|f| f.text.as_deref()),
            Some("1+2")
        );
        assert_eq!(
            cells[1].formula.as_ref().and_then(|f| f.text.as_deref()),
            Some("SUM({1,2})")
        );
        assert_eq!(
            cells[2].formula.as_ref().and_then(|f| f.text.as_deref()),
            Some("TRUE")
        );

        // The worksheet parser should only attempt rgce decoding once per formula cell.
        assert_eq!(formula_decode_attempts(), 3);
    }

    #[test]
    fn parse_ptg_exp_candidates_handles_multiple_coordinate_payload_layouts() {
        // `PtgExp` payloads appear in the wild in multiple layouts; we keep all plausible
        // interpretations so shared-formula materialization can match an actual `BrtShrFmla`
        // anchor.

        // BIFF8-style: row u16, col u16 (4-byte payload).
        let rgce_u16_u16 = vec![0x01, 0x34, 0x12, 0xBC, 0x0A]; // row=0x1234, col=0x0ABC
        assert_eq!(
            parse_ptg_exp_candidates(&rgce_u16_u16),
            Some(vec![(0x1234, 0x0ABC)])
        );

        // BIFF12-ish: row u32, col u16 (6-byte payload). Choose a row whose high u16 portion is a
        // plausible column index so we can observe both candidates.
        let row_u32: u32 = 0x0002_0010; // little-endian bytes: 10 00 02 00
        let col_u16: u16 = 5;
        let rgce_u32_u16 = {
            let mut v = Vec::new();
            v.push(0x01);
            v.extend_from_slice(&row_u32.to_le_bytes());
            v.extend_from_slice(&col_u16.to_le_bytes());
            v
        };
        // Candidates are ordered by how many bytes they consume (newest formats first).
        assert_eq!(
            parse_ptg_exp_candidates(&rgce_u32_u16),
            Some(vec![(row_u32, col_u16 as u32), (0x0010, 0x0002)])
        );

        // BIFF12-ish: row u32, col u32 (8-byte payload). The u32/u32 and u32/u16 interpretations
        // yield the same coordinate for valid column ranges, but we still want to preserve both
        // candidates for robustness.
        let col_u32: u32 = 7;
        let rgce_u32_u32 = {
            let mut v = Vec::new();
            v.push(0x01);
            v.extend_from_slice(&row_u32.to_le_bytes());
            v.extend_from_slice(&col_u32.to_le_bytes());
            v
        };
        assert_eq!(
            parse_ptg_exp_candidates(&rgce_u32_u32),
            Some(vec![
                (row_u32, col_u32),
                (row_u32, col_u32),
                (0x0010, 0x0002)
            ])
        );

        // Some producers include trailing bytes after the coordinates; we should still return the
        // same candidate list.
        let rgce_u32_u32_trailing = {
            let mut v = rgce_u32_u32.clone();
            v.extend_from_slice(&[0xAA, 0xBB, 0xCC]);
            v
        };
        assert_eq!(
            parse_ptg_exp_candidates(&rgce_u32_u32_trailing),
            Some(vec![
                (row_u32, col_u32),
                (row_u32, col_u32),
                (0x0010, 0x0002)
            ])
        );

        // Trailing bytes after a u32/u16 payload should not prevent parsing. Use non-zero trailing
        // bytes so the u32/u32 interpretation is rejected by the col bounds check.
        let rgce_u32_u16_trailing = {
            let mut v = rgce_u32_u16.clone();
            v.extend_from_slice(&0x1234u16.to_le_bytes());
            v
        };
        assert_eq!(
            parse_ptg_exp_candidates(&rgce_u32_u16_trailing),
            Some(vec![(row_u32, col_u16 as u32), (0x0010, 0x0002)])
        );
    }
}

pub(crate) fn parse_shared_strings<R: Read>(
    shared_strings_bin: &mut R,
) -> Result<Vec<SharedString>, Error> {
    let mut reader = Biff12Reader::new(shared_strings_bin);
    let mut buf = Vec::new();
    let mut strings = Vec::new();
    while let Some(rec) = reader.read_record(&mut buf)? {
        match rec.id {
            biff12::SI => {
                strings.push(parse_shared_string_item(rec.data)?);
            }
            biff12::SST_END => break,
            _ => {}
        }
    }
    Ok(strings)
}

fn parse_shared_string_item(data: &[u8]) -> Result<SharedString, Error> {
    let mut rr = RecordReader::new(data);
    let flags = rr.read_u8()?;
    let base_text = rr.read_utf16_string()?;

    // MS-XLSB `SI` flags:
    // - bit 0: rich text runs
    // - bit 1: phonetic (ruby) data
    // Higher bits are reserved but should be preserved.
    let has_rich = flags & 0x01 != 0;
    let has_phonetic = flags & 0x02 != 0;

    // Plain strings can be represented losslessly by their decoded text.
    if !has_rich && !has_phonetic && flags == 0 {
        return Ok(SharedString {
            rich_text: RichText::new(base_text),
            run_formats: Vec::new(),
            phonetic: None,
            raw_si: None,
        });
    }

    let raw_si = Some(data.to_vec());

    let (rich_text, run_formats, offset_after_rich) = if has_rich {
        let start_offset = rr.offset;
        match parse_rich_runs_best_effort(data, start_offset, &base_text, has_phonetic) {
            Some((rt, formats, new_offset)) => (rt, formats, new_offset),
            None => (RichText::new(base_text.clone()), Vec::new(), start_offset),
        }
    } else {
        (RichText::new(base_text.clone()), Vec::new(), rr.offset)
    };

    rr.offset = offset_after_rich;
    let remaining = rr.data.len().saturating_sub(rr.offset);
    let phonetic = if has_phonetic && remaining > 0 {
        Some(rr.read_slice(remaining)?.to_vec())
    } else {
        None
    };

    Ok(SharedString {
        rich_text,
        run_formats,
        phonetic,
        raw_si,
    })
}

/// Best-effort rich text run parsing.
///
/// XLSB rich runs reference fonts/styles; we treat the formatting bytes as opaque.
///
/// Returns `(rich_text, run_formats, new_offset)` on success, otherwise `None` and the caller
/// should fall back to treating the string as plain text while still preserving `raw_si`.
fn parse_rich_runs_best_effort(
    data: &[u8],
    offset: usize,
    text: &str,
    has_phonetic: bool,
) -> Option<(RichText, Vec<Vec<u8>>, usize)> {
    // Candidate layouts:
    // - MS-XLSB uses a u32 run count followed by `StrRun` entries.
    //   A `StrRun` is 8 bytes: [ich: u32][ifnt: u16][reserved: u16].
    // - Some producers may use legacy BIFF8-style runs (u16 count, 4-byte runs). We support
    //   this as a fallback to avoid hard failures on weird files.
    const LAYOUTS: &[(usize, usize, fn(&[u8]) -> usize)] = &[
        // (count_size, run_size, read_count_fn)
        (4, 8, |b| {
            u32::from_le_bytes([b[0], b[1], b[2], b[3]]) as usize
        }),
        (2, 4, |b| u16::from_le_bytes([b[0], b[1]]) as usize),
    ];

    for (count_size, run_size, read_count) in LAYOUTS {
        let mut off = offset;
        if data.len().saturating_sub(off) < *count_size {
            continue;
        }
        let count = read_count(&data[off..off + count_size]);
        off += count_size;

        // Sanity: avoid absurd allocations if the count is garbage.
        if count > 1_000_000 {
            continue;
        }

        let remaining = data.len().saturating_sub(off);
        let needed = count.checked_mul(*run_size)?;
        if needed > remaining {
            continue;
        }

        let run_bytes = &data[off..off + needed];
        off += needed;

        // If there is no phonetic data, we expect to consume the whole record payload.
        if !has_phonetic && off != data.len() {
            // Some files may include padding or unknown tail bytes; accept them by not requiring
            // full consumption, but prefer layouts that fully consume the record.
            // We'll still allow this layout if other layouts don't match.
        }

        let (starts_utf16, formats) = match *run_size {
            8 => parse_runs_8(run_bytes, count),
            4 => parse_runs_4(run_bytes, count),
            _ => continue,
        }?;

        let (rich_text, run_formats) = build_rich_text_from_runs(text, &starts_utf16, formats)?;
        return Some((rich_text, run_formats, off));
    }

    None
}

fn parse_runs_8(run_bytes: &[u8], count: usize) -> Option<(Vec<usize>, Vec<Vec<u8>>)> {
    if run_bytes.len() != count.checked_mul(8)? {
        return None;
    }
    let mut starts = Vec::new();
    let _ = starts.try_reserve_exact(count);
    let mut formats = Vec::new();
    let _ = formats.try_reserve_exact(count);
    for chunk in run_bytes.chunks_exact(8) {
        let ich = u32::from_le_bytes([chunk[0], chunk[1], chunk[2], chunk[3]]) as usize;
        starts.push(ich);
        formats.push(chunk[4..8].to_vec());
    }
    Some((starts, formats))
}

fn parse_runs_4(run_bytes: &[u8], count: usize) -> Option<(Vec<usize>, Vec<Vec<u8>>)> {
    if run_bytes.len() != count.checked_mul(4)? {
        return None;
    }
    let mut starts = Vec::new();
    let _ = starts.try_reserve_exact(count);
    let mut formats = Vec::new();
    let _ = formats.try_reserve_exact(count);
    for chunk in run_bytes.chunks_exact(4) {
        let ich = u16::from_le_bytes([chunk[0], chunk[1]]) as usize;
        starts.push(ich);
        formats.push(chunk[2..4].to_vec());
    }
    Some((starts, formats))
}

fn build_rich_text_from_runs(
    text: &str,
    starts_utf16: &[usize],
    mut formats: Vec<Vec<u8>>,
) -> Option<(RichText, Vec<Vec<u8>>)> {
    if starts_utf16.is_empty() {
        // Rich flag set but no runs. Treat as plain.
        return Some((RichText::new(text.to_string()), Vec::new()));
    }

    // Convert UTF-16 code unit indices into Rust `char` indices (Unicode scalar values).
    let mut starts: Vec<usize> = starts_utf16
        .iter()
        .filter_map(|&u16_idx| utf16_offset_to_char_index(text, u16_idx))
        .collect();

    // If we couldn't map all run boundaries, fall back to treating them as char indices.
    if starts.len() != starts_utf16.len() {
        starts = starts_utf16.to_vec();
    }

    // Pair starts with formats and sort by start index.
    let mut paired: Vec<(usize, Vec<u8>)> = starts.into_iter().zip(formats.drain(..)).collect();
    paired.sort_by_key(|(s, _)| *s);

    // Drop any runs that start beyond the end of the string.
    let char_len = text.chars().count();
    paired.retain(|(s, _)| *s <= char_len);
    if paired.is_empty() {
        return Some((RichText::new(text.to_string()), Vec::new()));
    }

    // Ensure we have a run starting at 0. If missing, insert a default run with empty formatting.
    if paired[0].0 != 0 {
        paired.insert(0, (0, Vec::new()));
    }

    let mut segments = Vec::new();
    let _ = segments.try_reserve_exact(paired.len());
    let mut out_formats = Vec::new();
    let _ = out_formats.try_reserve_exact(paired.len());

    for i in 0..paired.len() {
        let start = paired[i].0;
        let end = paired.get(i + 1).map(|(s, _)| *s).unwrap_or(char_len);
        let seg = slice_by_char_range(text, start, end).to_string();
        segments.push((seg, RichTextRunStyle::default()));
        out_formats.push(paired[i].1.clone());
    }

    Some((RichText::from_segments(segments), out_formats))
}

fn utf16_offset_to_char_index(text: &str, utf16_offset: usize) -> Option<usize> {
    if utf16_offset == 0 {
        return Some(0);
    }

    let mut u16_cursor = 0usize;
    let mut char_cursor = 0usize;

    for ch in text.chars() {
        let ch_u16 = ch.len_utf16();
        if u16_cursor + ch_u16 > utf16_offset {
            // Inside a UTF-16 codepoint (surrogate pair boundary mismatch).
            return None;
        }
        u16_cursor += ch_u16;
        char_cursor += 1;
        if u16_cursor == utf16_offset {
            return Some(char_cursor);
        }
    }

    if u16_cursor == utf16_offset {
        Some(char_cursor)
    } else {
        None
    }
}

fn slice_by_char_range(text: &str, start: usize, end: usize) -> &str {
    if start == end {
        return "";
    }

    let mut start_byte = None;
    let mut end_byte = None;

    for (i, (byte_idx, _ch)) in text.char_indices().enumerate() {
        if i == start {
            start_byte = Some(byte_idx);
        }
        if i == end {
            end_byte = Some(byte_idx);
            break;
        }
    }

    let start_byte = start_byte.unwrap_or_else(|| text.len());
    let end_byte = end_byte.unwrap_or_else(|| text.len());

    &text[start_byte..end_byte]
}

pub(crate) fn parse_sheet<R: Read>(
    sheet_bin: &mut R,
    shared_strings: &[String],
    shared_strings_table: Option<&[SharedString]>,
    ctx: &WorkbookContext,
    preserve_parsed_parts: bool,
    decode_formulas: bool,
) -> Result<SheetData, Error> {
    let mut cells = Vec::new();
    let dimension = parse_sheet_stream(
        sheet_bin,
        shared_strings,
        shared_strings_table,
        ctx,
        preserve_parsed_parts,
        decode_formulas,
        |cell| {
            cells.push(cell);
            ControlFlow::Continue(())
        },
    )?;
    Ok(SheetData { dimension, cells })
}

pub(crate) fn parse_sheet_stream<R: Read, F: FnMut(Cell) -> ControlFlow<(), ()>>(
    sheet_bin: &mut R,
    shared_strings: &[String],
    shared_strings_table: Option<&[SharedString]>,
    ctx: &WorkbookContext,
    preserve_parsed_parts: bool,
    decode_formulas: bool,
    mut on_cell: F,
) -> Result<Option<Dimension>, Error> {
    let mut reader = Biff12Reader::new(sheet_bin);
    let mut buf = Vec::new();
    let mut dimension: Option<Dimension> = None;

    let mut in_sheet_data = false;
    let mut current_row: Option<u32> = None;
    let mut shared_formulas: HashMap<(u32, u32), SharedFormulaDef> = HashMap::new();

    'records: while let Some(rec) = reader.read_record(&mut buf)? {
        match rec.id {
            biff12::DIMENSION => {
                let mut rr = RecordReader::new(rec.data);
                let r1 = rr.read_u32()?;
                let r2 = rr.read_u32()?;
                let c1 = rr.read_u32()?;
                let c2 = rr.read_u32()?;
                dimension = Some(Dimension {
                    start_row: r1,
                    start_col: c1,
                    height: r2.saturating_sub(r1) + 1,
                    width: c2.saturating_sub(c1) + 1,
                });
            }
            biff12::SHEETDATA => {
                in_sheet_data = true;
            }
            biff12::SHEETDATA_END => {
                break;
            }
            biff12::SHR_FMLA if in_sheet_data => {
                if let Some(def) = SharedFormulaDef::parse(rec.data) {
                    // Key by the base cell (top-left of the range). `PtgExp`
                    // tokens inside the range typically reference this anchor.
                    shared_formulas.insert((def.base_row, def.base_col), def);
                }
            }
            biff12::ROW if in_sheet_data => {
                let mut rr = RecordReader::new(rec.data);
                current_row = Some(rr.read_u32()?);
            }
            biff12::BLANK
            | biff12::NUM
            | biff12::BOOLERR
            | biff12::BOOL
            | biff12::FLOAT
            | biff12::CELL_ST
            | biff12::STRING
            | biff12::FORMULA_STRING
            | biff12::FORMULA_FLOAT
            | biff12::FORMULA_BOOL
            | biff12::FORMULA_BOOLERR
                if in_sheet_data =>
            {
                let row = current_row.unwrap_or(0);
                let mut rr = RecordReader::new(rec.data);
                let col = rr.read_u32()?;
                let style = rr.read_u32()?;
                let base = crate::rgce::CellCoord::new(row, col);

                let (value, formula, preserved_string) = match rec.id {
                    biff12::BLANK => (CellValue::Blank, None, None),
                    biff12::NUM => (CellValue::Number(rr.read_rk_number()?), None, None),
                    biff12::BOOLERR => (CellValue::Error(rr.read_u8()?), None, None),
                    biff12::BOOL => (CellValue::Bool(rr.read_u8()? != 0), None, None),
                    biff12::FLOAT => (CellValue::Number(rr.read_f64()?), None, None),
                    biff12::CELL_ST => {
                        // BrtCellSt inline strings appear in the wild in at least two layouts:
                        // - simple wide string: [cch: u32][utf16 chars...]
                        // - rich/phonetic wide string: [cch: u32][flags: u8][utf16 chars...][extras...]
                        //
                        // When patch-writing we currently emit the simple layout, so the reader
                        // must accept both.
                        let start_offset = rr.offset;
                        // NOTE: We preserve rich/phonetic extra blocks even when
                        // `preserve_parsed_parts=false` so downstream consumers (e.g. model export)
                        // can access metadata like phonetic guides without requiring raw part
                        // preservation.
                        // Some streams contain malformed/mixed layout strings (e.g. a missing
                        // rich/phonetic block when the corresponding flag bit is set, or trailing
                        // junk bytes after a simple inline string). Be tolerant and choose between
                        // the simple vs flagged UTF-16 start offsets based on which slice looks
                        // more like UTF-16LE text.
                        let cch = {
                            let bytes: [u8; 4] = rr
                                .data
                                .get(start_offset..start_offset + 4)
                                .ok_or(Error::UnexpectedEof)?
                                .try_into()
                                .unwrap();
                            u32::from_le_bytes(bytes) as usize
                        };
                        let utf16_len = cch.checked_mul(2).ok_or(Error::UnexpectedEof)?;
                        let simple_utf16_start =
                            start_offset.checked_add(4).ok_or(Error::UnexpectedEof)?;
                        let simple_utf16_end = simple_utf16_start
                            .checked_add(utf16_len)
                            .ok_or(Error::UnexpectedEof)?;
                        let flagged_utf16_start = simple_utf16_start
                            .checked_add(1)
                            .ok_or(Error::UnexpectedEof)?;
                        let flagged_utf16_end = flagged_utf16_start
                            .checked_add(utf16_len)
                            .ok_or(Error::UnexpectedEof)?;

                        let has_simple_bytes = simple_utf16_end <= rr.data.len();
                        let has_flagged_bytes = flagged_utf16_end <= rr.data.len();

                        let score_utf16_candidate = |raw: &[u8]| -> i32 {
                            // Score a candidate UTF-16LE byte slice by attempting to decode it.
                            // This is used to disambiguate the "simple" (`[cch][utf16...]`) vs
                            // "flagged" (`[cch][flags:u8][utf16...]`) BrtCellSt layouts when the
                            // record length is ambiguous (e.g. trailing bytes).
                            //
                            // Keep this bounded: strings can be large, and we only need enough
                            // signal to pick the correct offset.
                            const MAX_DECODED_CHARS: usize = 16;
                            let iter = raw
                                .chunks_exact(2)
                                .map(|chunk| u16::from_le_bytes([chunk[0], chunk[1]]));
                            let mut score = 0i32;
                            for decoded in std::char::decode_utf16(iter).take(MAX_DECODED_CHARS) {
                                match decoded {
                                    Ok(ch) => {
                                        let cp = ch as u32;
                                        // NUL and control characters are uncommon in visible cell
                                        // text and are often a sign of misalignment.
                                        if ch == '\0' {
                                            score -= 5;
                                        } else if cp >= 0x1_0000 {
                                            // Characters outside the BMP require valid surrogate
                                            // pairs in UTF-16. Misaligned parsing is unlikely to
                                            // accidentally form valid pairs, so prefer candidates
                                            // that contain supplementary-plane chars.
                                            score += 3;
                                        } else if ch.is_control()
                                            && ch != '\t'
                                            && ch != '\n'
                                            && ch != '\r'
                                        {
                                            score -= 2;
                                        } else if (0xE000..=0xF8FF).contains(&cp) {
                                            // Private-use characters are rare in normal cell text.
                                            score -= 2;
                                        } else {
                                            score += 1;
                                        }
                                    }
                                    Err(_) => {
                                        // Invalid surrogate pairing (common when misaligned).
                                        score -= 5;
                                    }
                                }
                            }
                            score
                        };

                        let choose_flagged = match (has_simple_bytes, has_flagged_bytes) {
                            (true, false) => false,
                            (false, true) => true,
                            (false, false) => return Err(Error::UnexpectedEof),
                            (true, true) => {
                                // Heuristic: if the would-be flags byte contains bits outside the
                                // observed BrtCellSt wide-string flag set (rich/phonetic +
                                // reserved 0x80), treat the record as using the simple layout.
                                //
                                // This avoids misinterpreting the first UTF-16 code unit byte as
                                // a flags byte for simple-layout strings that include trailing
                                // bytes (seen in the wild), which can shift the decode by 1 byte
                                // and corrupt the decoded text (especially for non-ASCII strings).
                                if rr.data[simple_utf16_start] & !0x83 != 0 {
                                    false
                                } else {
                                    let simple_score = score_utf16_candidate(
                                        &rr.data[simple_utf16_start..simple_utf16_end],
                                    );
                                    let flagged_score = score_utf16_candidate(
                                        &rr.data[flagged_utf16_start..flagged_utf16_end],
                                    );
                                    // Tie-break in favor of the flagged layout (more common in the
                                    // wild and required for rich/phonetic preservation).
                                    flagged_score >= simple_score
                                }
                            }
                        };

                        let parsed = if choose_flagged {
                            // Flagged layout: try the full `XLWideString` parse first so rich /
                            // phonetic extras are preserved. If the extras are missing (malformed
                            // flags) fall back to a minimal parse that reads only `[cch][flags][utf16]`.
                            rr.offset = start_offset;
                            match read_xl_wide_string_with_flags(&mut rr, FlagsWidth::U8, true) {
                                Ok((_flags, parsed)) => parsed,
                                Err(Error::UnexpectedEof) => {
                                    rr.offset = start_offset;
                                    let cch = rr.read_u32()? as usize;
                                    let _flags = rr.read_u8()?;
                                    ParsedXlsbString {
                                        text: rr.read_utf16_chars(cch)?,
                                        rich: None,
                                        phonetic: None,
                                    }
                                }
                                Err(e) => return Err(e),
                            }
                        } else {
                            rr.offset = start_offset;
                            ParsedXlsbString {
                                text: rr.read_utf16_string()?,
                                rich: None,
                                phonetic: None,
                            }
                        };

                        if parsed.rich.is_some() || parsed.phonetic.is_some() {
                            (CellValue::Text(parsed.text.clone()), None, Some(parsed))
                        } else {
                            (CellValue::Text(parsed.text), None, None)
                        }
                    }
                    biff12::STRING => {
                        let idx = rr.read_u32()? as usize;
                        let text = shared_strings
                            .get(idx)
                            .cloned()
                            .or_else(|| {
                                shared_strings_table
                                    .and_then(|table| table.get(idx))
                                    .map(|s| s.plain_text().to_string())
                            })
                            .unwrap_or_default();

                        let preserved = shared_strings_table
                            .and_then(|table| table.get(idx))
                            .and_then(|s| {
                                // Match the behavior for inline string records: preserve phonetic
                                // guides (furigana) so downstream consumers (e.g. `formula-io`
                                // model conversion) can access them even when
                                // `preserve_parsed_parts=false` (raw-part preservation off).
                                //
                                // Rich text runs can be much larger; only preserve them when
                                // `preserve_parsed_parts` is enabled.
                                let has_rich = preserve_parsed_parts && !s.rich_text.is_plain();
                                let has_phonetic = s.phonetic.is_some();
                                if has_rich || has_phonetic {
                                    let rich = if has_rich {
                                        // Convert our parsed shared-string run boundaries into an
                                        // XLWideString-style `StrRun` byte array:
                                        //   [ich: u32 (UTF-16 code units)][ifnt: u16][reserved: u16]
                                        //
                                        // This is best-effort: we preserve the original `ifnt`
                                        // bytes (and any other run-format bytes) opaquely, and
                                        // recompute `ich` from the decoded string.
                                        let mut utf16_offsets: Vec<u32> = Vec::new();
                                        let _ = utf16_offsets
                                            .try_reserve_exact(s.rich_text.char_len() + 1);
                                        let mut u16_cursor: u32 = 0;
                                        utf16_offsets.push(0);
                                        for ch in s.rich_text.text.chars() {
                                            u16_cursor =
                                                u16_cursor.saturating_add(ch.len_utf16() as u32);
                                            utf16_offsets.push(u16_cursor);
                                        }

                                        let mut runs_bytes: Vec<u8> = Vec::new();
                                        let _ = runs_bytes
                                            .try_reserve_exact(s.rich_text.runs.len() * 8);
                                        for (i, run) in s.rich_text.runs.iter().enumerate() {
                                            let ich = utf16_offsets
                                                .get(run.start)
                                                .copied()
                                                .unwrap_or(run.start as u32);
                                            runs_bytes.extend_from_slice(&ich.to_le_bytes());

                                            let fmt = s.run_formats.get(i).map(|v| v.as_slice());
                                            match fmt.map(|v| v.len()).unwrap_or(0) {
                                                0 => runs_bytes.extend_from_slice(&[0u8; 4]),
                                                2 => {
                                                    runs_bytes.extend_from_slice(fmt.unwrap());
                                                    runs_bytes.extend_from_slice(&[0u8; 2]);
                                                }
                                                4 => runs_bytes.extend_from_slice(fmt.unwrap()),
                                                n => {
                                                    let fmt = fmt.unwrap();
                                                    let take = n.min(4);
                                                    runs_bytes.extend_from_slice(&fmt[..take]);
                                                    if take < 4 {
                                                        runs_bytes.extend_from_slice(
                                                            &[0u8; 4][..4 - take],
                                                        );
                                                    }
                                                }
                                            }
                                        }
                                        Some(crate::strings::OpaqueRichText { runs: runs_bytes })
                                    } else {
                                        None
                                    };

                                    Some(ParsedXlsbString {
                                        text: text.clone(),
                                        rich,
                                        phonetic: s.phonetic.clone(),
                                    })
                                } else {
                                    None
                                }
                            });
                        (CellValue::Text(text), None, preserved)
                    }
                    biff12::FORMULA_STRING => {
                        // BrtFmlaString stores the cached string value as an XLWideString-like
                        // payload (cch + flags + utf16 + optional rich/phonetic blocks),
                        // followed by the formula token stream ([cce][rgce...]).
                        //
                        // In practice, some writers appear to set the "rich/phonetic" bits in
                        // the flags field without supplying the corresponding extra blocks
                        // (reusing those bits for unrelated formula flags). Use a heuristic:
                        // prefer parsing rich/phonetic payloads when they keep the trailing
                        // `[cce][rgce]` fields in-bounds; otherwise, fall back to treating the
                        // cached value as a plain UTF-16 string.
                        let start_offset = rr.offset;
                        let mut parsed = None;
                        match read_xl_wide_string_with_flags(&mut rr, FlagsWidth::U16, true) {
                            Ok((flags, candidate)) => {
                                let cce_offset = rr.offset;
                                let mut accept = false;
                                if let Some(raw) = rr.data.get(cce_offset..cce_offset + 4) {
                                    let cce = u32::from_le_bytes([raw[0], raw[1], raw[2], raw[3]])
                                        as usize;
                                    let rgce_offset = cce_offset + 4;
                                    if let Some(rgce_end) = rgce_offset.checked_add(cce) {
                                        if rgce_end <= rr.data.len() {
                                            accept = true;
                                        }
                                    }
                                }
                                if accept {
                                    parsed = Some((flags, candidate));
                                } else {
                                    rr.offset = start_offset;
                                }
                            }
                            Err(Error::UnexpectedEof) => {
                                rr.offset = start_offset;
                            }
                            Err(e) => return Err(e),
                        }

                        let (flags, v, preserved) = if let Some((flags, parsed)) = parsed {
                            if parsed.rich.is_some() || parsed.phonetic.is_some() {
                                (flags, parsed.text.clone(), Some(parsed))
                            } else {
                                (flags, parsed.text, None)
                            }
                        } else {
                            let cch = rr.read_u32()? as usize;
                            let flags = rr.read_u16()?;
                            (flags, rr.read_utf16_chars(cch)?, None)
                        };
                        let cce = rr.read_u32()? as usize;
                        let mut rgce = rr.read_slice(cce)?.to_vec();
                        let extra = rr.data[rr.offset..].to_vec();
                        let mut rgcb_for_decode: &[u8] = &extra;
                        if let Some(materialized) =
                            materialize_shared_formula(&rgce, row, col, &shared_formulas, ctx)
                        {
                            rgce = materialized.rgce;
                            if rgcb_for_decode.is_empty() {
                                rgcb_for_decode = materialized.rgcb;
                            }
                        }

                        let (text, warnings) = if decode_formulas {
                            let decoded = decode_formula_text(&rgce, rgcb_for_decode, ctx, base);
                            (decoded.text, decoded.warnings)
                        } else {
                            (None, Vec::new())
                        };
                        (
                            CellValue::Text(v),
                            Some(Formula {
                                rgce,
                                text,
                                flags,
                                extra,
                                warnings,
                            }),
                            preserved,
                        )
                    }
                    biff12::FORMULA_FLOAT => {
                        // BrtFmlaNum: [value: f64][flags: u16][cce: u32][rgce bytes...]
                        let v = rr.read_f64()?;
                        let flags = rr.read_u16()?;
                        let cce = rr.read_u32()? as usize;
                        let mut rgce = rr.read_slice(cce)?.to_vec();
                        let extra = rr.data[rr.offset..].to_vec();
                        let mut rgcb_for_decode: &[u8] = &extra;
                        if let Some(materialized) =
                            materialize_shared_formula(&rgce, row, col, &shared_formulas, ctx)
                        {
                            rgce = materialized.rgce;
                            if rgcb_for_decode.is_empty() {
                                rgcb_for_decode = materialized.rgcb;
                            }
                        }
                        let (text, warnings) = if decode_formulas {
                            let decoded = decode_formula_text(&rgce, rgcb_for_decode, ctx, base);
                            (decoded.text, decoded.warnings)
                        } else {
                            (None, Vec::new())
                        };
                        (
                            CellValue::Number(v),
                            Some(Formula {
                                rgce,
                                text,
                                flags,
                                extra,
                                warnings,
                            }),
                            None,
                        )
                    }
                    biff12::FORMULA_BOOL => {
                        // BrtFmlaBool: [value: u8][flags: u16][cce: u32][rgce bytes...]
                        let v = rr.read_u8()? != 0;
                        let flags = rr.read_u16()?;
                        let cce = rr.read_u32()? as usize;
                        let mut rgce = rr.read_slice(cce)?.to_vec();
                        let extra = rr.data[rr.offset..].to_vec();
                        let mut rgcb_for_decode: &[u8] = &extra;
                        if let Some(materialized) =
                            materialize_shared_formula(&rgce, row, col, &shared_formulas, ctx)
                        {
                            rgce = materialized.rgce;
                            if rgcb_for_decode.is_empty() {
                                rgcb_for_decode = materialized.rgcb;
                            }
                        }
                        let (text, warnings) = if decode_formulas {
                            let decoded = decode_formula_text(&rgce, rgcb_for_decode, ctx, base);
                            (decoded.text, decoded.warnings)
                        } else {
                            (None, Vec::new())
                        };
                        (
                            CellValue::Bool(v),
                            Some(Formula {
                                rgce,
                                text,
                                flags,
                                extra,
                                warnings,
                            }),
                            None,
                        )
                    }
                    biff12::FORMULA_BOOLERR => {
                        // BrtFmlaError: [value: u8][flags: u16][cce: u32][rgce bytes...]
                        let v = rr.read_u8()?;
                        let flags = rr.read_u16()?;
                        let cce = rr.read_u32()? as usize;
                        let mut rgce = rr.read_slice(cce)?.to_vec();
                        let extra = rr.data[rr.offset..].to_vec();
                        let mut rgcb_for_decode: &[u8] = &extra;
                        if let Some(materialized) =
                            materialize_shared_formula(&rgce, row, col, &shared_formulas, ctx)
                        {
                            rgce = materialized.rgce;
                            if rgcb_for_decode.is_empty() {
                                rgcb_for_decode = materialized.rgcb;
                            }
                        }
                        let (text, warnings) = if decode_formulas {
                            let decoded = decode_formula_text(&rgce, rgcb_for_decode, ctx, base);
                            (decoded.text, decoded.warnings)
                        } else {
                            (None, Vec::new())
                        };
                        (
                            CellValue::Error(v),
                            Some(Formula {
                                rgce,
                                text,
                                flags,
                                extra,
                                warnings,
                            }),
                            None,
                        )
                    }
                    _ => unreachable!(),
                };

                let cell = Cell {
                    row,
                    col,
                    style,
                    value,
                    formula,
                    preserved_string,
                };
                if let ControlFlow::Break(()) = on_cell(cell) {
                    break 'records;
                }
            }
            _ => {}
        }
    }

    Ok(dimension)
}

#[derive(Debug, Clone)]
struct SharedFormulaDef {
    base_row: u32,
    base_col: u32,
    range_r1: u32,
    range_r2: u32,
    range_c1: u32,
    range_c2: u32,
    rgce: Vec<u8>,
    rgcb: Vec<u8>,
}

impl SharedFormulaDef {
    fn parse(data: &[u8]) -> Option<Self> {
        if data.len() < 20 {
            return None;
        }

        let range_r1 = u32::from_le_bytes(data.get(0..4)?.try_into().ok()?);
        let range_r2 = u32::from_le_bytes(data.get(4..8)?.try_into().ok()?);
        let range_c1 = u32::from_le_bytes(data.get(8..12)?.try_into().ok()?);
        let range_c2 = u32::from_le_bytes(data.get(12..16)?.try_into().ok()?);

        // Basic sanity checks to avoid misclassifying unrelated records.
        if range_r1 > range_r2 || range_c1 > range_c2 {
            return None;
        }

        let tail = &data[16..];
        let (rgce, rgce_end) = parse_rgce_tail(tail)?;
        let rgcb = tail.get(rgce_end..)?.to_vec();

        Some(Self {
            base_row: range_r1,
            base_col: range_c1,
            range_r1,
            range_r2,
            range_c1,
            range_c2,
            rgce,
            rgcb,
        })
    }

    fn contains_cell(&self, row: u32, col: u32) -> bool {
        row >= self.range_r1 && row <= self.range_r2 && col >= self.range_c1 && col <= self.range_c2
    }
}

fn parse_rgce_tail(tail: &[u8]) -> Option<(Vec<u8>, usize)> {
    // BrtShrFmla contains the rgce length and bytes, but the header layout can
    // vary slightly between producers. We accept a few common shapes:
    // - [cce: u32][rgce...]
    // - [flags: u16][cce: u32][rgce...]
    // - [flags: u32][cce: u32][rgce...]

    for &(prefix, cce_offset) in &[(0usize, 0usize), (2, 2), (4, 4)] {
        if tail.len() < prefix + 4 {
            continue;
        }
        let cce =
            u32::from_le_bytes(tail.get(cce_offset..cce_offset + 4)?.try_into().ok()?) as usize;
        let rgce_start = prefix + 4;
        if tail.len() < rgce_start + cce {
            continue;
        }
        let rgce = tail.get(rgce_start..rgce_start + cce)?.to_vec();
        return Some((rgce, rgce_start + cce));
    }

    None
}

fn materialize_shared_formula<'a>(
    rgce: &[u8],
    row: u32,
    col: u32,
    shared_formulas: &'a HashMap<(u32, u32), SharedFormulaDef>,
    ctx: &WorkbookContext,
) -> Option<MaterializedSharedFormula<'a>> {
    let candidates = parse_ptg_exp_candidates(rgce)?;

    for (base_row, base_col) in candidates {
        let Some(def) = shared_formulas
            .get(&(base_row, base_col))
            .filter(|def| def.contains_cell(row, col))
        else {
            continue;
        };

        // Produce a cell-specific rgce so callers don't need shared-formula context.
        let rgce = materialize_rgce(&def.rgce, base_row, base_col, row, col, ctx)?;
        return Some(MaterializedSharedFormula {
            rgce,
            rgcb: &def.rgcb,
        });
    }

    None
}

struct MaterializedSharedFormula<'a> {
    rgce: Vec<u8>,
    rgcb: &'a [u8],
}

fn parse_ptg_exp_candidates(rgce: &[u8]) -> Option<Vec<(u32, u32)>> {
    // PtgExp is used by shared formulas / array formulas to refer back to the
    // "master" formula. In practice it's usually the entire rgce for a cell.
    if rgce.first().copied()? != 0x01 {
        return None;
    }
    let payload = &rgce[1..];
    if payload.len() < 4 {
        return None;
    }

    const MAX_ROW: u32 = 1_048_575;
    const MAX_COL: u32 = 16_383;

    // Some writers appear to include trailing bytes after `PtgExp` coordinates. To stay robust,
    // collect all plausible interpretations and let the shared-formula lookup decide which one
    // matches an actual `BrtShrFmla` anchor.
    let mut candidates: Vec<(u32, u32, usize)> = Vec::new();

    // BIFF12-ish: row u32, col u32.
    if payload.len() >= 8 {
        let row = u32::from_le_bytes(payload.get(0..4)?.try_into().ok()?);
        let col = u32::from_le_bytes(payload.get(4..8)?.try_into().ok()?);
        if row <= MAX_ROW && col <= MAX_COL {
            candidates.push((row, col, 8));
        }
    }

    // BIFF12-ish: row u32, col u16.
    if payload.len() >= 6 {
        let row = u32::from_le_bytes(payload.get(0..4)?.try_into().ok()?);
        let col = u16::from_le_bytes(payload.get(4..6)?.try_into().ok()?) as u32;
        if row <= MAX_ROW && col <= MAX_COL {
            candidates.push((row, col, 6));
        }
    }

    // BIFF8-style: row u16, col u16.
    let row = u16::from_le_bytes(payload.get(0..2)?.try_into().ok()?) as u32;
    let col = u16::from_le_bytes(payload.get(2..4)?.try_into().ok()?) as u32;
    if row <= MAX_ROW && col <= MAX_COL {
        candidates.push((row, col, 4));
    }

    if candidates.is_empty() {
        return None;
    }

    // Prefer candidates that consume more bytes (newer formats), but always keep all viable ones.
    candidates.sort_by_key(|(_, _, n)| *n);
    let out: Vec<(u32, u32)> = candidates
        .into_iter()
        .rev()
        .map(|(r, c, _)| (r, c))
        .collect();
    Some(out)
}

fn materialize_rgce(
    base: &[u8],
    base_row: u32,
    base_col: u32,
    row: u32,
    col: u32,
    ctx: &WorkbookContext,
) -> Option<Vec<u8>> {
    const MAX_ROW: i64 = 1_048_575;
    const MAX_COL: i64 = 16_383;

    let delta_row = row as i64 - base_row as i64;
    let delta_col = col as i64 - base_col as i64;

    let mut out = Vec::new();
    let _ = out.try_reserve_exact(base.len());
    let mut i = 0usize;
    while i < base.len() {
        let ptg = *base.get(i)?;
        i += 1;

        match ptg {
            // Fixed-width / no-payload tokens we already support elsewhere.
            0x03..=0x16 | 0x2F => out.push(ptg),
            0x17 => {
                // PtgStr: [cch: u16][utf16 chars...]
                if i + 2 > base.len() {
                    return None;
                }
                let cch = u16::from_le_bytes([base[i], base[i + 1]]) as usize;
                let bytes = cch.checked_mul(2)?;
                if i + 2 + bytes > base.len() {
                    return None;
                }
                out.push(ptg);
                out.extend_from_slice(&base[i..i + 2 + bytes]);
                i += 2 + bytes;
            }
            0x18 | 0x38 | 0x58 => {
                // PtgExtend / PtgExtendV / PtgExtendA.
                //
                // Used by structured references (table formulas) and other newer operand tokens.
                // Structured references do not embed relative row/col offsets, so we can copy the
                // token verbatim during shared-formula materialization.
                let etpg = *base.get(i)?;
                i += 1;
                out.push(ptg);
                out.push(etpg);

                match etpg {
                    // PtgList (structured reference / table ref): payload layout is best-effort.
                    0x19 => {
                        let remaining = base.get(i..)?;
                        let payload_len =
                            crate::rgce::ptg_list_payload_len_best_effort(remaining, Some(ctx))?;
                        let payload = remaining.get(..payload_len)?;
                        out.extend_from_slice(payload);
                        i += payload_len;
                    }
                    _ => return None,
                }
            }
            0x1C | 0x1D => {
                // PtgErr / PtgBool: 1 byte.
                out.push(ptg);
                out.push(*base.get(i)?);
                i += 1;
            }
            0x1E => {
                // PtgInt: 2 bytes.
                out.push(ptg);
                out.extend_from_slice(base.get(i..i + 2)?);
                i += 2;
            }
            0x1F => {
                // PtgNum: 8 bytes.
                out.push(ptg);
                out.extend_from_slice(base.get(i..i + 8)?);
                i += 8;
            }
            0x20 | 0x40 | 0x60 => {
                // PtgArray: [unused: 7 bytes] + array data stored in the trailing rgcb stream.
                //
                // Arrays don't embed relative references in the token stream, so we can copy the
                // token verbatim while leaving rgcb unchanged.
                if i + 7 > base.len() {
                    return None;
                }
                out.push(ptg);
                out.extend_from_slice(&base[i..i + 7]);
                i += 7;
            }
            0x21 | 0x41 | 0x61 => {
                // PtgFunc: [iftab: u16]
                out.push(ptg);
                out.extend_from_slice(base.get(i..i + 2)?);
                i += 2;
            }
            0x22 | 0x42 | 0x62 => {
                // PtgFuncVar: [argc: u8][iftab: u16]
                out.push(ptg);
                out.extend_from_slice(base.get(i..i + 3)?);
                i += 3;
            }
            0x23 | 0x43 | 0x63 => {
                // PtgName: [nameId: u32][reserved: u16]
                out.push(ptg);
                out.extend_from_slice(base.get(i..i + 6)?);
                i += 6;
            }
            0x39 | 0x59 | 0x79 => {
                // PtgNameX: [ixti: u16][nameIndex: u16]
                out.push(ptg);
                out.extend_from_slice(base.get(i..i + 4)?);
                i += 4;
            }
            0x3A | 0x5A | 0x7A => {
                // PtgRef3d: [ixti: u16][row: u32][col+flags: u16]
                //
                // Like `PtgRef`, the row/col fields are absolute coordinates with relative flags
                // in the high bits. Shared formulas can contain 3D references, so we need to
                // shift relative refs when materializing across the shared range.
                let ixti = u16::from_le_bytes(base.get(i..i + 2)?.try_into().ok()?);
                let row_raw = u32::from_le_bytes(base.get(i + 2..i + 6)?.try_into().ok()?) as i64;
                let col_raw_u16 = u16::from_le_bytes(base.get(i + 6..i + 8)?.try_into().ok()?);
                let col_raw = (col_raw_u16 & 0x3FFF) as i64;
                let row_rel = (col_raw_u16 & 0x4000) != 0;
                let col_rel = (col_raw_u16 & 0x8000) != 0;

                let new_row = if row_rel {
                    row_raw + delta_row
                } else {
                    row_raw
                };
                let new_col = if col_rel {
                    col_raw + delta_col
                } else {
                    col_raw
                };

                if new_row < 0 || new_row > MAX_ROW || new_col < 0 || new_col > MAX_COL {
                    out.push(ptg.saturating_add(0x02)); // PtgRef3d* -> PtgRefErr3d*
                    out.extend_from_slice(base.get(i..i + 8)?);
                    i += 8;
                    continue;
                }

                out.push(ptg);
                out.extend_from_slice(&ixti.to_le_bytes());
                out.extend_from_slice(&(new_row as u32).to_le_bytes());
                let new_col_u16 = pack_col_flags(new_col as u32, row_rel, col_rel)?;
                out.extend_from_slice(&new_col_u16.to_le_bytes());
                i += 8;
            }
            0x3B | 0x5B | 0x7B => {
                // PtgArea3d: [ixti: u16][r1: u32][r2: u32][c1+flags: u16][c2+flags: u16]
                //
                // Like `PtgArea`, area endpoints have independent relative flags. Materialize by
                // shifting any relative endpoints by the shared-formula delta.
                let ixti = u16::from_le_bytes(base.get(i..i + 2)?.try_into().ok()?);
                let r1_raw = u32::from_le_bytes(base.get(i + 2..i + 6)?.try_into().ok()?) as i64;
                let r2_raw = u32::from_le_bytes(base.get(i + 6..i + 10)?.try_into().ok()?) as i64;
                let c1_u16 = u16::from_le_bytes(base.get(i + 10..i + 12)?.try_into().ok()?);
                let c2_u16 = u16::from_le_bytes(base.get(i + 12..i + 14)?.try_into().ok()?);

                let c1_raw = (c1_u16 & 0x3FFF) as i64;
                let c2_raw = (c2_u16 & 0x3FFF) as i64;
                let r1_rel = (c1_u16 & 0x4000) != 0;
                let c1_rel = (c1_u16 & 0x8000) != 0;
                let r2_rel = (c2_u16 & 0x4000) != 0;
                let c2_rel = (c2_u16 & 0x8000) != 0;

                let new_r1 = if r1_rel { r1_raw + delta_row } else { r1_raw };
                let new_c1 = if c1_rel { c1_raw + delta_col } else { c1_raw };
                let new_r2 = if r2_rel { r2_raw + delta_row } else { r2_raw };
                let new_c2 = if c2_rel { c2_raw + delta_col } else { c2_raw };

                if new_r1 < 0
                    || new_r1 > MAX_ROW
                    || new_r2 < 0
                    || new_r2 > MAX_ROW
                    || new_c1 < 0
                    || new_c1 > MAX_COL
                    || new_c2 < 0
                    || new_c2 > MAX_COL
                {
                    out.push(ptg.saturating_add(0x02)); // PtgArea3d* -> PtgAreaErr3d*
                    out.extend_from_slice(base.get(i..i + 14)?);
                    i += 14;
                    continue;
                }

                out.push(ptg);
                out.extend_from_slice(&ixti.to_le_bytes());
                out.extend_from_slice(&(new_r1 as u32).to_le_bytes());
                out.extend_from_slice(&(new_r2 as u32).to_le_bytes());
                let new_c1_u16 = pack_col_flags(new_c1 as u32, r1_rel, c1_rel)?;
                let new_c2_u16 = pack_col_flags(new_c2 as u32, r2_rel, c2_rel)?;
                out.extend_from_slice(&new_c1_u16.to_le_bytes());
                out.extend_from_slice(&new_c2_u16.to_le_bytes());
                i += 14;
            }
            0x26 | 0x46 | 0x66 | 0x27 | 0x47 | 0x67 | 0x28 | 0x48 | 0x68 | 0x29 | 0x49 | 0x69
            | 0x2E | 0x4E | 0x6E => {
                // PtgMem*: [cce: u16][rgce subexpression bytes...]
                //
                // In BIFF12, the `cce` field is followed immediately by `cce` bytes containing a
                // nested rgce subexpression. These tokens are usually ignored for printing, but
                // the nested bytes still need to be materialized (shift relative refs) so the
                // dependency graph stays consistent for shared formulas.
                if i + 2 > base.len() {
                    return None;
                }
                let cce = u16::from_le_bytes([base[i], base[i + 1]]) as usize;
                out.push(ptg);
                out.extend_from_slice(&base[i..i + 2]);
                i += 2;

                if i + cce > base.len() {
                    return None;
                }
                let nested = base.get(i..i + cce)?;
                let nested_out = materialize_rgce(nested, base_row, base_col, row, col, ctx)?;
                // Materialization should preserve encoded size; bail out defensively otherwise.
                if nested_out.len() != cce {
                    return None;
                }
                out.extend_from_slice(&nested_out);
                i += cce;
            }
            0x19 => {
                // PtgAttr: [grbit: u8][wAttr: u16]
                //
                // When materializing shared formulas we generally keep `PtgAttr` tokens as-is,
                // but we still need to copy the payload (and any attribute-specific tail bytes)
                // so the output rgce stays aligned.
                if i + 3 > base.len() {
                    return None;
                }
                out.push(ptg);
                let grbit = base[i];
                let w_attr = u16::from_le_bytes([base[i + 1], base[i + 2]]);
                out.extend_from_slice(&base[i..i + 3]);
                i += 3;

                const T_ATTR_CHOOSE: u8 = 0x04;
                if grbit & T_ATTR_CHOOSE != 0 {
                    let needed = (w_attr as usize).checked_mul(2)?;
                    if i + needed > base.len() {
                        return None;
                    }
                    out.extend_from_slice(&base[i..i + needed]);
                    i += needed;
                }
            }
            0x24 | 0x44 | 0x64 => {
                // PtgRef: [row: u32][col+flags: u16]
                let row_raw = u32::from_le_bytes(base.get(i..i + 4)?.try_into().ok()?) as i64;
                let col_raw_u16 = u16::from_le_bytes(base.get(i + 4..i + 6)?.try_into().ok()?);
                let col_raw = (col_raw_u16 & 0x3FFF) as i64;
                let row_rel = (col_raw_u16 & 0x4000) != 0;
                let col_rel = (col_raw_u16 & 0x8000) != 0;

                let new_row = if row_rel {
                    row_raw + delta_row
                } else {
                    row_raw
                };
                let new_col = if col_rel {
                    col_raw + delta_col
                } else {
                    col_raw
                };

                if new_row < 0 || new_row > MAX_ROW || new_col < 0 || new_col > MAX_COL {
                    // Excel represents invalid references using `PtgRefErr`. When shared formulas
                    // are filled across a range near the sheet edges, relative references can
                    // overflow the valid row/col bounds. Materialize those as `#REF!` tokens
                    // instead of aborting materialization entirely.
                    out.push(ptg.saturating_add(0x06)); // PtgRef* -> PtgRefErr* (class-preserving)
                    out.extend_from_slice(base.get(i..i + 6)?);
                    i += 6;
                    continue;
                }

                out.push(ptg);
                out.extend_from_slice(&(new_row as u32).to_le_bytes());
                let new_col_u16 = pack_col_flags(new_col as u32, row_rel, col_rel)?;
                out.extend_from_slice(&new_col_u16.to_le_bytes());
                i += 6;
            }
            0x25 | 0x45 | 0x65 => {
                // PtgArea: [r1: u32][r2: u32][c1+flags: u16][c2+flags: u16]
                let r1_raw = u32::from_le_bytes(base.get(i..i + 4)?.try_into().ok()?) as i64;
                let r2_raw = u32::from_le_bytes(base.get(i + 4..i + 8)?.try_into().ok()?) as i64;
                let c1_u16 = u16::from_le_bytes(base.get(i + 8..i + 10)?.try_into().ok()?);
                let c2_u16 = u16::from_le_bytes(base.get(i + 10..i + 12)?.try_into().ok()?);

                let c1_raw = (c1_u16 & 0x3FFF) as i64;
                let c2_raw = (c2_u16 & 0x3FFF) as i64;
                let r1_rel = (c1_u16 & 0x4000) != 0;
                let c1_rel = (c1_u16 & 0x8000) != 0;
                let r2_rel = (c2_u16 & 0x4000) != 0;
                let c2_rel = (c2_u16 & 0x8000) != 0;

                let new_r1 = if r1_rel { r1_raw + delta_row } else { r1_raw };
                let new_c1 = if c1_rel { c1_raw + delta_col } else { c1_raw };
                let new_r2 = if r2_rel { r2_raw + delta_row } else { r2_raw };
                let new_c2 = if c2_rel { c2_raw + delta_col } else { c2_raw };

                if new_r1 < 0
                    || new_r1 > MAX_ROW
                    || new_r2 < 0
                    || new_r2 > MAX_ROW
                    || new_c1 < 0
                    || new_c1 > MAX_COL
                    || new_c2 < 0
                    || new_c2 > MAX_COL
                {
                    // Emit an error-range token when the adjusted area exceeds sheet bounds.
                    out.push(ptg.saturating_add(0x06)); // PtgArea* -> PtgAreaErr* (class-preserving)
                    out.extend_from_slice(base.get(i..i + 12)?);
                    i += 12;
                    continue;
                }

                out.push(ptg);
                out.extend_from_slice(&(new_r1 as u32).to_le_bytes());
                out.extend_from_slice(&(new_r2 as u32).to_le_bytes());
                let new_c1_u16 = pack_col_flags(new_c1 as u32, r1_rel, c1_rel)?;
                let new_c2_u16 = pack_col_flags(new_c2 as u32, r2_rel, c2_rel)?;
                out.extend_from_slice(&new_c1_u16.to_le_bytes());
                out.extend_from_slice(&new_c2_u16.to_le_bytes());
                i += 12;
            }
            0x2C | 0x4C | 0x6C => {
                // PtgRefN: relative row/col offsets (commonly used in shared formulas).
                // Layout is best-effort:
                // - [row_off: i32][col_off: i16]
                let row_off = i32::from_le_bytes(base.get(i..i + 4)?.try_into().ok()?) as i64;
                let col_off = i16::from_le_bytes(base.get(i + 4..i + 6)?.try_into().ok()?) as i64;
                let abs_row = row as i64 + row_off;
                let abs_col = col as i64 + col_off;
                if abs_row < 0 || abs_row > MAX_ROW || abs_col < 0 || abs_col > MAX_COL {
                    out.push(ptg.saturating_sub(0x02)); // PtgRefN* -> PtgRefErr* (class-preserving)
                    out.extend_from_slice(base.get(i..i + 6)?);
                    i += 6;
                    continue;
                }
                // Convert to a normal PtgRef token.
                out.push(ptg - 0x08);
                out.extend_from_slice(&(abs_row as u32).to_le_bytes());
                let col_u16 = pack_col_flags(abs_col as u32, true, true)?;
                out.extend_from_slice(&col_u16.to_le_bytes());
                i += 6;
            }
            0x2D | 0x4D | 0x6D => {
                // PtgAreaN: relative row/col offsets (commonly used in shared formulas).
                // Layout is best-effort:
                // - [r1_off: i32][r2_off: i32][c1_off: i16][c2_off: i16]
                let r1_off = i32::from_le_bytes(base.get(i..i + 4)?.try_into().ok()?) as i64;
                let r2_off = i32::from_le_bytes(base.get(i + 4..i + 8)?.try_into().ok()?) as i64;
                let c1_off = i16::from_le_bytes(base.get(i + 8..i + 10)?.try_into().ok()?) as i64;
                let c2_off = i16::from_le_bytes(base.get(i + 10..i + 12)?.try_into().ok()?) as i64;

                let abs_r1 = row as i64 + r1_off;
                let abs_r2 = row as i64 + r2_off;
                let abs_c1 = col as i64 + c1_off;
                let abs_c2 = col as i64 + c2_off;

                if abs_r1 < 0
                    || abs_r1 > MAX_ROW
                    || abs_r2 < 0
                    || abs_r2 > MAX_ROW
                    || abs_c1 < 0
                    || abs_c1 > MAX_COL
                    || abs_c2 < 0
                    || abs_c2 > MAX_COL
                {
                    out.push(ptg.saturating_sub(0x02)); // PtgAreaN* -> PtgAreaErr* (class-preserving)
                    out.extend_from_slice(base.get(i..i + 12)?);
                    i += 12;
                    continue;
                }

                out.push(ptg - 0x08);
                out.extend_from_slice(&(abs_r1 as u32).to_le_bytes());
                out.extend_from_slice(&(abs_r2 as u32).to_le_bytes());
                let c1_u16 = pack_col_flags(abs_c1 as u32, true, true)?;
                let c2_u16 = pack_col_flags(abs_c2 as u32, true, true)?;
                out.extend_from_slice(&c1_u16.to_le_bytes());
                out.extend_from_slice(&c2_u16.to_le_bytes());
                i += 12;
            }
            0x2A | 0x4A | 0x6A => {
                // PtgRefErr: [row: u32][col+flags: u16]
                out.push(ptg);
                out.extend_from_slice(base.get(i..i + 6)?);
                i += 6;
            }
            0x2B | 0x4B | 0x6B => {
                // PtgAreaErr: [r1: u32][r2: u32][c1+flags: u16][c2+flags: u16]
                out.push(ptg);
                out.extend_from_slice(base.get(i..i + 12)?);
                i += 12;
            }
            0x3C | 0x5C | 0x7C => {
                // PtgRefErr3d: [ixti: u16][row: u32][col+flags: u16]
                out.push(ptg);
                out.extend_from_slice(base.get(i..i + 8)?);
                i += 8;
            }
            0x3D | 0x5D | 0x7D => {
                // PtgAreaErr3d: [ixti: u16][r1: u32][r2: u32][c1+flags: u16][c2+flags: u16]
                out.push(ptg);
                out.extend_from_slice(base.get(i..i + 14)?);
                i += 14;
            }
            _ => return None,
        }
    }

    Some(out)
}

fn pack_col_flags(col: u32, row_rel: bool, col_rel: bool) -> Option<u16> {
    if col > 0x3FFF {
        return None;
    }
    let mut v = col as u16;
    if row_rel {
        v |= 0x4000;
    }
    if col_rel {
        v |= 0x8000;
    }
    Some(v)
}
fn normalize_sheet_target(target: &str) -> String {
    // Relationship targets are typically relative to `xl/` (the directory containing
    // `xl/workbook.bin`), but some producers emit absolute targets (leading `/`) or include an
    // `xl/` prefix despite the target being relative.
    //
    // Be tolerant and normalize:
    // - backslashes to `/`
    // - strip leading `/`
    // - avoid double-prefixing `xl/` when it is already present.
    let target = target.trim_start_matches(|c| c == '/' || c == '\\');
    let target = if target.contains('\\') {
        std::borrow::Cow::Owned(target.replace('\\', "/"))
    } else {
        std::borrow::Cow::Borrowed(target)
    };
    let target = target.as_ref();
    if target
        .get(.."xl/".len())
        .is_some_and(|prefix| prefix.eq_ignore_ascii_case("xl/"))
    {
        target.to_string()
    } else {
        format!("xl/{target}")
    }
}

fn is_defined_name_record(id: u32) -> bool {
    matches!(id, biff12::NAME | 0x0018)
}

fn is_supbook_record(id: u32) -> bool {
    matches!(
        id,
        // BIFF8 `SupBook`. In XLSB this appears as a BIFF12 record id.
        0x00AE
            // Common BIFF12 candidates observed in the wild (keep parsing robust across writers).
            | 0x0162
            | 0x0161
    )
}

fn is_end_supbook_record(id: u32) -> bool {
    matches!(id, 0x0163 | 0x00AF)
}

fn is_extern_sheet_record(id: u32) -> bool {
    matches!(
        id,
        // BIFF8 `ExternSheet`
        0x0017
            // Common BIFF12 candidate.
            | 0x0167
            // MS-XLSB `BrtExternSheet` record id used by Excel (and Calamine).
            | 0x016A
    )
}

fn is_extern_name_record(id: u32) -> bool {
    matches!(
        id,
        // BIFF8 `ExternName`
        0x0023
            // Common BIFF12 candidate.
            | 0x0168
    )
}

fn parse_supbook(data: &[u8]) -> Option<(SupBook, Vec<String>)> {
    // Try a few plausible layouts:
    // - u16 ctab + utf16string (BIFF8-like)
    // - u32 ctab + utf16string (BIFF12-like)
    {
        let mut rr = RecordReader::new(data);
        if let Ok(ctab) = rr.read_u16() {
            if let Ok(raw_name) = rr.read_utf16_string() {
                let kind = classify_supbook_name(&raw_name);
                let sheet_names = read_supbook_sheet_names(&mut rr, ctab as usize);
                return Some((SupBook { raw_name, kind }, sheet_names));
            }
        }
    }

    {
        let mut rr = RecordReader::new(data);
        if let Ok(ctab) = rr.read_u32() {
            if let Ok(raw_name) = rr.read_utf16_string() {
                let kind = classify_supbook_name(&raw_name);
                let sheet_names = read_supbook_sheet_names(&mut rr, ctab as usize);
                return Some((SupBook { raw_name, kind }, sheet_names));
            }
        }
    }

    None
}

fn read_supbook_sheet_names(rr: &mut RecordReader<'_>, ctab: usize) -> Vec<String> {
    // SupBook `ctab` is the number of sheet names in the record.
    //
    // We parse best-effort: if anything looks malformed, fall back to no sheet list so callers
    // can still decode formulas deterministically using placeholders.
    const MAX_SUPBOOK_SHEETS: usize = 16_384;
    if ctab == 0 || ctab > MAX_SUPBOOK_SHEETS {
        return Vec::new();
    }

    let start_offset = rr.offset;
    let mut out = Vec::new();
    let _ = out.try_reserve_exact(ctab);
    for _ in 0..ctab {
        match rr.read_utf16_string() {
            Ok(s) => out.push(s),
            Err(_) => {
                rr.offset = start_offset;
                return Vec::new();
            }
        }
    }
    out
}

fn resolve_supbook_sheet_name(sheet_list: Option<&Vec<String>>, sheet_index: u32) -> String {
    sheet_list
        .and_then(|sheets| sheets.get(sheet_index as usize))
        .cloned()
        .unwrap_or_else(|| {
            // ExternSheet sheet indices are 0-based. Use the traditional 1-based "Sheet{n}" naming
            // so the fallback is stable and looks like a plausible Excel sheet name.
            let n = sheet_index.checked_add(1).unwrap_or(sheet_index);
            format!("Sheet{n}")
        })
}

const WORKBOOK_PATH_SUFFIXES: [&str; 11] = [
    ".xls", ".xlt", ".xla", ".xlsx", ".xlsm", ".xltx", ".xltm", ".xlsb", ".xlam", ".xll",
    // Legacy Excel 2-4 format extensions are uncommon but still appear in old workbooks.
    ".xlw",
];

fn looks_like_workbook_path(value: &str) -> bool {
    if value.contains(['/', '\\']) {
        return true;
    }
    for suffix in WORKBOOK_PATH_SUFFIXES {
        if value
            .get(value.len().saturating_sub(suffix.len())..)
            .is_some_and(|tail| tail.eq_ignore_ascii_case(suffix))
        {
            return true;
        }
    }
    false
}

fn classify_supbook_name(raw_name: &str) -> SupBookKind {
    if raw_name.is_empty() {
        return SupBookKind::Internal;
    }
    if raw_name == "\u{0001}" {
        return SupBookKind::AddIn;
    }

    // Heuristic: if the string looks like a file path or workbook filename, treat it as an
    // external workbook reference. Otherwise treat it as an internal SupBook (some producers
    // store the first sheet name here).
    if looks_like_workbook_path(raw_name) {
        SupBookKind::ExternalWorkbook
    } else {
        SupBookKind::Internal
    }
}

fn supbook_is_plausible(supbook: &SupBook) -> bool {
    if supbook.raw_name.is_empty() || supbook.raw_name == "\u{0001}" {
        return true;
    }

    looks_like_workbook_path(&supbook.raw_name)
}

fn parse_extern_sheet(data: &[u8]) -> Option<Vec<ExternSheet>> {
    // BIFF8 layout: u16 cxti + cxti * (u16, u16, u16)
    if data.len() >= 2 {
        let cxti = u16::from_le_bytes([data[0], data[1]]) as usize;
        if data.len() == 2 + cxti * 6 {
            let mut out = Vec::new();
            let _ = out.try_reserve_exact(cxti);
            let mut offset = 2;
            for _ in 0..cxti {
                let supbook = u16::from_le_bytes([data[offset], data[offset + 1]]);
                let first = u16::from_le_bytes([data[offset + 2], data[offset + 3]]) as u32;
                let last = u16::from_le_bytes([data[offset + 4], data[offset + 5]]) as u32;
                out.push(ExternSheet {
                    supbook_index: supbook,
                    sheet_first: first,
                    sheet_last: last,
                });
                offset += 6;
            }
            return Some(out);
        }
    }

    // BIFF12-like layout: u32 cxti + entries (either 6 or 12 bytes each).
    if data.len() >= 4 {
        let cxti = u32::from_le_bytes([data[0], data[1], data[2], data[3]]) as usize;
        if data.len() == 4 + cxti * 6 {
            let mut out = Vec::new();
            let _ = out.try_reserve_exact(cxti);
            let mut offset = 4;
            for _ in 0..cxti {
                let supbook = u16::from_le_bytes([data[offset], data[offset + 1]]);
                let first = u16::from_le_bytes([data[offset + 2], data[offset + 3]]) as u32;
                let last = u16::from_le_bytes([data[offset + 4], data[offset + 5]]) as u32;
                out.push(ExternSheet {
                    supbook_index: supbook,
                    sheet_first: first,
                    sheet_last: last,
                });
                offset += 6;
            }
            return Some(out);
        }

        if data.len() == 4 + cxti * 12 {
            let mut out = Vec::new();
            let _ = out.try_reserve_exact(cxti);
            let mut offset = 4;
            for _ in 0..cxti {
                let supbook = u32::from_le_bytes([
                    data[offset],
                    data[offset + 1],
                    data[offset + 2],
                    data[offset + 3],
                ]) as u16;
                let first = u32::from_le_bytes([
                    data[offset + 4],
                    data[offset + 5],
                    data[offset + 6],
                    data[offset + 7],
                ]);
                let last = u32::from_le_bytes([
                    data[offset + 8],
                    data[offset + 9],
                    data[offset + 10],
                    data[offset + 11],
                ]);
                out.push(ExternSheet {
                    supbook_index: supbook,
                    sheet_first: first,
                    sheet_last: last,
                });
                offset += 12;
            }
            return Some(out);
        }
    }

    None
}

fn parse_extern_name(data: &[u8]) -> Option<ExternName> {
    // Try a few plausible layouts. We only need the name string and (optionally) sheet scope.
    let mut rr = RecordReader::new(data);
    let flags = rr.read_u16().ok()?;

    // Layout A: flags: u16, scope: u16, name: xlWideString
    if let Ok(scope) = rr.read_u16() {
        if let Ok(name) = rr.read_utf16_string() {
            return Some(ExternName {
                name,
                is_function: flags & 0x0002 != 0,
                scope_sheet: scope_to_option(scope as u32),
            });
        }
    }

    // Layout B: flags: u16, scope: u32, name: xlWideString
    let mut rr = RecordReader::new(data);
    let flags = rr.read_u16().ok()?;
    if let Ok(scope) = rr.read_u32() {
        if let Ok(name) = rr.read_utf16_string() {
            return Some(ExternName {
                name,
                is_function: flags & 0x0002 != 0,
                scope_sheet: scope_to_option(scope),
            });
        }
    }

    None
}

fn scope_to_option(scope: u32) -> Option<u32> {
    // Many BIFF structures represent "no sheet scope" as either:
    // - `0xFFFF` (u16 sentinel)
    // - `0xFFFFFFFF` (u32 / i32=-1 sentinel)
    //
    // Some BIFF12 producers also use other negative `i32` values (e.g. `-2`), so treat any
    // negative `i32` as workbook scope.
    if scope == 0xFFFF || scope == 0xFFFFFFFF {
        return None;
    }
    if (scope as i32) < 0 {
        return None;
    }
    Some(scope)
}

fn parse_defined_name_record(data: &[u8]) -> Option<DefinedName> {
    // MS-XLSB `BrtName` (record id `0x0027`) layout as produced by Excel and observed in other
    // readers (e.g. Calamine):
    //   [flags: u32][itab: u32/i32][reserved: u8][name: xlWideString][cce: u32][rgce bytes...][rgcb bytes...]
    //
    // Notes:
    // - `itab` is the sheet scope: a negative `i32` indicates workbook scope, otherwise it's a
    //   0-based sheet index.
    // - Excel's BIFF `Name` flags use bit 0 as "hidden"; we preserve that here.
    if let Some(parsed) = (|| {
        let mut rr = RecordReader::new(data);
        let flags = rr.read_u32().ok()?;
        let scope_raw = rr.read_u32().ok()?;
        rr.read_u8().ok()?; // reserved / unused

        let name = rr.read_utf16_string().ok()?;
        let rgce_len = rr.read_u32().ok()? as usize;
        let rgce = rr.read_slice(rgce_len).ok()?.to_vec();
        let extra = rr.data[rr.offset..].to_vec();

        Some(DefinedName {
            index: 0, // patched by caller
            name,
            scope_sheet: scope_to_option(scope_raw),
            hidden: (flags & 0x0001) != 0,
            formula: Some(Formula {
                rgce,
                text: None,
                flags: 0,
                extra,
                warnings: Vec::new(),
            }),
            comment: None,
        })
    })() {
        return Some(parsed);
    }

    // Legacy / alternate layouts: best-effort parse of just the name + scope.
    // We intentionally do not attempt to parse the `refersTo` formula for these.
    let legacy = |name: String, scope_raw: u32, hidden: bool| DefinedName {
        index: 0, // patched by caller
        name,
        scope_sheet: scope_to_option(scope_raw),
        hidden,
        formula: None,
        comment: None,
    };

    // Layout A: [flags: u16][scope: u32][name: xlWideString]
    {
        let mut rr = RecordReader::new(data);
        let flags = rr.read_u16().ok()?;
        if let Ok(scope) = rr.read_u32() {
            if let Ok(name) = rr.read_utf16_string() {
                return Some(legacy(name, scope, (flags & 0x0001) != 0));
            }
        }
    }

    // Layout B: [flags: u16][scope: u16][name: xlWideString]
    {
        let mut rr = RecordReader::new(data);
        let flags = rr.read_u16().ok()?;
        if let Ok(scope) = rr.read_u16() {
            if let Ok(name) = rr.read_utf16_string() {
                return Some(legacy(name, scope as u32, (flags & 0x0001) != 0));
            }
        }
    }

    None
}
