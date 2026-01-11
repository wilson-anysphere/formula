use std::collections::HashMap;
use std::io::{self, BufReader, Read};

use crate::biff12_varint;
use thiserror::Error;

// Record IDs (BIFF12 / MS-XLSB). Values taken from pyxlsb (public domain-ish) and MS-XLSB.
#[allow(dead_code)]
pub(crate) mod biff12 {
    pub const SHEETS_END: u32 = 0x0190;
    pub const SHEET: u32 = 0x019C;

    pub const WORKSHEET: u32 = 0x0181;
    pub const WORKSHEET_END: u32 = 0x0182;
    pub const SHEETDATA: u32 = 0x0191;
    pub const SHEETDATA_END: u32 = 0x0192;
    pub const DIMENSION: u32 = 0x0194;

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
    //
    // NOTE: Record IDs are decoded by `Biff12Reader::read_record()` in the same
    // way as the rest of this file (continuation bits are preserved), hence the
    // literal value here follows that convention.
    pub const SHR_FMLA: u32 = 0x0010;

    pub const SST: u32 = 0x019F;
    pub const SST_END: u32 = 0x01A0;
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
}

#[derive(Debug, Clone, PartialEq, Eq)]
pub struct SheetMeta {
    pub name: String,
    pub part_path: String,
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
        buf.resize(len, 0);
        self.inner.read_exact(buf)?;
        Ok(Some(Biff12Record { id, data: buf }))
    }

    fn read_id(&mut self) -> Result<Option<u32>, Error> {
        Ok(biff12_varint::read_record_id(&mut self.inner)?)
    }

    fn read_len(&mut self) -> Result<Option<u32>, Error> {
        Ok(biff12_varint::read_record_len(&mut self.inner)?)
    }
}

struct RecordReader<'a> {
    data: &'a [u8],
    offset: usize,
}

impl<'a> RecordReader<'a> {
    fn new(data: &'a [u8]) -> Self {
        Self { data, offset: 0 }
    }

    fn skip(&mut self, n: usize) -> Result<(), Error> {
        self.offset = self
            .offset
            .checked_add(n)
            .filter(|&o| o <= self.data.len())
            .ok_or(Error::UnexpectedEof)?;
        Ok(())
    }

    fn read_u8(&mut self) -> Result<u8, Error> {
        let b = *self.data.get(self.offset).ok_or(Error::UnexpectedEof)?;
        self.offset += 1;
        Ok(b)
    }

    fn read_u16(&mut self) -> Result<u16, Error> {
        let bytes: [u8; 2] = self
            .data
            .get(self.offset..self.offset + 2)
            .ok_or(Error::UnexpectedEof)?
            .try_into()
            .unwrap();
        self.offset += 2;
        Ok(u16::from_le_bytes(bytes))
    }

    fn read_u32(&mut self) -> Result<u32, Error> {
        let bytes: [u8; 4] = self
            .data
            .get(self.offset..self.offset + 4)
            .ok_or(Error::UnexpectedEof)?
            .try_into()
            .unwrap();
        self.offset += 4;
        Ok(u32::from_le_bytes(bytes))
    }

    fn read_f64(&mut self) -> Result<f64, Error> {
        let bytes: [u8; 8] = self
            .data
            .get(self.offset..self.offset + 8)
            .ok_or(Error::UnexpectedEof)?
            .try_into()
            .unwrap();
        self.offset += 8;
        Ok(f64::from_le_bytes(bytes))
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

    fn read_utf16_chars(&mut self, len_chars: usize) -> Result<String, Error> {
        let byte_len = len_chars.checked_mul(2).ok_or(Error::UnexpectedEof)?;
        let raw = self
            .data
            .get(self.offset..self.offset + byte_len)
            .ok_or(Error::UnexpectedEof)?;
        self.offset += byte_len;

        let mut units = Vec::with_capacity(len_chars);
        for chunk in raw.chunks_exact(2) {
            units.push(u16::from_le_bytes([chunk[0], chunk[1]]));
        }
        Ok(String::from_utf16_lossy(&units))
    }

    fn read_slice(&mut self, len: usize) -> Result<&'a [u8], Error> {
        let raw = self
            .data
            .get(self.offset..self.offset + len)
            .ok_or(Error::UnexpectedEof)?;
        self.offset += len;
        Ok(raw)
    }
}

pub(crate) fn parse_workbook_sheets<R: Read>(
    workbook_bin: &mut R,
    rels: &HashMap<String, String>,
) -> Result<Vec<SheetMeta>, Error> {
    let mut reader = Biff12Reader::new(workbook_bin);
    let mut buf = Vec::new();
    let mut sheets = Vec::new();
    while let Some(rec) = reader.read_record(&mut buf)? {
        match rec.id {
            biff12::SHEET => {
                let mut rr = RecordReader::new(rec.data);
                rr.skip(4)?; // unknown flags / state
                let _sheet_id = rr.read_u32()?;
                let rel_id = rr.read_utf16_string()?;
                let name = rr.read_utf16_string()?;
                let Some(target) = rels.get(&rel_id) else {
                    return Err(Error::MissingSheetRelationship(rel_id));
                };
                let part_path = normalize_sheet_target(target);
                sheets.push(SheetMeta { name, part_path });
            }
            biff12::SHEETS_END => break,
            _ => {}
        }
    }
    Ok(sheets)
}

pub(crate) fn parse_shared_strings<R: Read>(
    shared_strings_bin: &mut R,
) -> Result<Vec<String>, Error> {
    let mut reader = Biff12Reader::new(shared_strings_bin);
    let mut buf = Vec::new();
    let mut strings = Vec::new();
    while let Some(rec) = reader.read_record(&mut buf)? {
        match rec.id {
            biff12::SI => {
                let mut rr = RecordReader::new(rec.data);
                // Flags byte (rich text / phonetic) – not handled yet.
                rr.skip(1)?;
                strings.push(rr.read_utf16_string()?);
            }
            biff12::SST_END => break,
            _ => {}
        }
    }
    Ok(strings)
}

pub(crate) fn parse_sheet<R: Read>(
    sheet_bin: &mut R,
    shared_strings: &[String],
) -> Result<SheetData, Error> {
    let mut cells = Vec::new();
    let dimension = parse_sheet_stream(sheet_bin, shared_strings, |cell| cells.push(cell))?;
    Ok(SheetData { dimension, cells })
}

pub(crate) fn parse_sheet_stream<R: Read, F: FnMut(Cell)>(
    sheet_bin: &mut R,
    shared_strings: &[String],
    mut on_cell: F,
) -> Result<Option<Dimension>, Error> {
    let mut reader = Biff12Reader::new(sheet_bin);
    let mut buf = Vec::new();
    let mut dimension: Option<Dimension> = None;

    let mut in_sheet_data = false;
    let mut current_row: Option<u32> = None;
    let mut shared_formulas: HashMap<(u32, u32), SharedFormulaDef> = HashMap::new();

    while let Some(rec) = reader.read_record(&mut buf)? {
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

                let (value, formula) = match rec.id {
                    biff12::BLANK => (CellValue::Blank, None),
                    biff12::NUM => (CellValue::Number(rr.read_rk_number()?), None),
                    biff12::BOOLERR => (CellValue::Error(rr.read_u8()?), None),
                    biff12::BOOL => (CellValue::Bool(rr.read_u8()? != 0), None),
                    biff12::FLOAT => (CellValue::Number(rr.read_f64()?), None),
                    biff12::CELL_ST => (CellValue::Text(rr.read_utf16_string()?), None),
                    biff12::STRING => {
                        let idx = rr.read_u32()? as usize;
                        let s = shared_strings.get(idx).cloned().unwrap_or_default();
                        (CellValue::Text(s), None)
                    }
                    biff12::FORMULA_STRING => {
                        // BrtFmlaString: [cch: u32][flags: u16][utf16 chars][cce: u32][rgce bytes...]
                        let cch = rr.read_u32()? as usize;
                        let flags = rr.read_u16()?;
                        let v = rr.read_utf16_chars(cch)?;
                        let cce = rr.read_u32()? as usize;
                        let mut rgce = rr.read_slice(cce)?.to_vec();
                        if let Some(materialized) =
                            materialize_shared_formula(&rgce, row, col, &shared_formulas)
                        {
                            rgce = materialized;
                        }
                        let text = crate::rgce::decode_rgce(&rgce).ok();
                        let extra = rr.data[rr.offset..].to_vec();
                        (
                            CellValue::Text(v),
                            Some(Formula {
                                rgce,
                                text,
                                flags,
                                extra,
                            }),
                        )
                    }
                    biff12::FORMULA_FLOAT => {
                        // BrtFmlaNum: [value: f64][flags: u16][cce: u32][rgce bytes...]
                        let v = rr.read_f64()?;
                        let flags = rr.read_u16()?;
                        let cce = rr.read_u32()? as usize;
                        let mut rgce = rr.read_slice(cce)?.to_vec();
                        if let Some(materialized) =
                            materialize_shared_formula(&rgce, row, col, &shared_formulas)
                        {
                            rgce = materialized;
                        }
                        let text = crate::rgce::decode_rgce(&rgce).ok();
                        let extra = rr.data[rr.offset..].to_vec();
                        (
                            CellValue::Number(v),
                            Some(Formula {
                                rgce,
                                text,
                                flags,
                                extra,
                            }),
                        )
                    }
                    biff12::FORMULA_BOOL => {
                        // BrtFmlaBool: [value: u8][flags: u16][cce: u32][rgce bytes...]
                        let v = rr.read_u8()? != 0;
                        let flags = rr.read_u16()?;
                        let cce = rr.read_u32()? as usize;
                        let mut rgce = rr.read_slice(cce)?.to_vec();
                        if let Some(materialized) =
                            materialize_shared_formula(&rgce, row, col, &shared_formulas)
                        {
                            rgce = materialized;
                        }
                        let text = crate::rgce::decode_rgce(&rgce).ok();
                        let extra = rr.data[rr.offset..].to_vec();
                        (
                            CellValue::Bool(v),
                            Some(Formula {
                                rgce,
                                text,
                                flags,
                                extra,
                            }),
                        )
                    }
                    biff12::FORMULA_BOOLERR => {
                        // BrtFmlaError: [value: u8][flags: u16][cce: u32][rgce bytes...]
                        let v = rr.read_u8()?;
                        let flags = rr.read_u16()?;
                        let cce = rr.read_u32()? as usize;
                        let mut rgce = rr.read_slice(cce)?.to_vec();
                        if let Some(materialized) =
                            materialize_shared_formula(&rgce, row, col, &shared_formulas)
                        {
                            rgce = materialized;
                        }
                        let text = crate::rgce::decode_rgce(&rgce).ok();
                        let extra = rr.data[rr.offset..].to_vec();
                        (
                            CellValue::Error(v),
                            Some(Formula {
                                rgce,
                                text,
                                flags,
                                extra,
                            }),
                        )
                    }
                    _ => unreachable!(),
                };

                on_cell(Cell {
                    row,
                    col,
                    style,
                    value,
                    formula,
                });
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

        let (rgce, _) = parse_rgce_tail(&data[16..])?;

        Some(Self {
            base_row: range_r1,
            base_col: range_c1,
            range_r1,
            range_r2,
            range_c1,
            range_c2,
            rgce,
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
        let cce = u32::from_le_bytes(tail.get(cce_offset..cce_offset + 4)?.try_into().ok()?) as usize;
        let rgce_start = prefix + 4;
        if tail.len() < rgce_start + cce {
            continue;
        }
        let rgce = tail.get(rgce_start..rgce_start + cce)?.to_vec();
        return Some((rgce, rgce_start + cce));
    }

    None
}

fn materialize_shared_formula(
    rgce: &[u8],
    row: u32,
    col: u32,
    shared_formulas: &HashMap<(u32, u32), SharedFormulaDef>,
) -> Option<Vec<u8>> {
    let (base_row, base_col) = parse_ptg_exp(rgce)?;

    let base_rgce = shared_formulas
        .get(&(base_row, base_col))
        .filter(|def| def.contains_cell(row, col))
        .map(|def| def.rgce.as_slice())?;

    // Produce a cell-specific rgce so callers don't need shared-formula context.
    materialize_rgce(base_rgce, base_row, base_col, row, col)
}

fn parse_ptg_exp(rgce: &[u8]) -> Option<(u32, u32)> {
    // PtgExp is used by shared formulas / array formulas to refer back to the
    // "master" formula. In practice it's usually the entire rgce for a cell.
    if rgce.first().copied()? != 0x01 {
        return None;
    }
    let payload = &rgce[1..];
    match payload.len() {
        // BIFF8-style: row u16, col u16.
        4 => {
            let row = u16::from_le_bytes(payload.get(0..2)?.try_into().ok()?) as u32;
            let col = u16::from_le_bytes(payload.get(2..4)?.try_into().ok()?) as u32;
            Some((row, col))
        }
        // BIFF12-ish: row u32, col u16.
        6 => {
            let row = u32::from_le_bytes(payload.get(0..4)?.try_into().ok()?);
            let col = u16::from_le_bytes(payload.get(4..6)?.try_into().ok()?) as u32;
            Some((row, col))
        }
        // BIFF12-ish: row u32, col u32.
        8 => {
            let row = u32::from_le_bytes(payload.get(0..4)?.try_into().ok()?);
            let col = u32::from_le_bytes(payload.get(4..8)?.try_into().ok()?);
            Some((row, col))
        }
        _ => None,
    }
}

fn materialize_rgce(base: &[u8], base_row: u32, base_col: u32, row: u32, col: u32) -> Option<Vec<u8>> {
    const MAX_ROW: i64 = 1_048_575;
    const MAX_COL: i64 = 16_383;

    let delta_row = row as i64 - base_row as i64;
    let delta_col = col as i64 - base_col as i64;

    let mut out = Vec::with_capacity(base.len());
    let mut i = 0usize;
    while i < base.len() {
        let ptg = *base.get(i)?;
        i += 1;

        match ptg {
            // Fixed-width / no-payload tokens we already support elsewhere.
            0x03..=0x16 => out.push(ptg),
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
            0x24 | 0x44 | 0x64 => {
                // PtgRef: [row: u32][col+flags: u16]
                let row_raw = u32::from_le_bytes(base.get(i..i + 4)?.try_into().ok()?) as i64;
                let col_raw_u16 = u16::from_le_bytes(base.get(i + 4..i + 6)?.try_into().ok()?);
                let col_raw = (col_raw_u16 & 0x3FFF) as i64;
                let row_rel = (col_raw_u16 & 0x4000) != 0;
                let col_rel = (col_raw_u16 & 0x8000) != 0;

                let new_row = if row_rel { row_raw + delta_row } else { row_raw };
                let new_col = if col_rel { col_raw + delta_col } else { col_raw };

                if new_row < 0 || new_row > MAX_ROW || new_col < 0 || new_col > MAX_COL {
                    return None;
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
                    return None;
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
                    return None;
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
                    return None;
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
    // Relationship targets are typically relative to `xl/`.
    let target = target.trim_start_matches('/');
    format!("xl/{}", target.replace('\\', "/"))
}
