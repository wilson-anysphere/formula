use std::collections::HashMap;
use std::io::{self, BufReader, Read};

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
    pub const STRING: u32 = 0x0007;
    pub const FORMULA_STRING: u32 = 0x0008;
    pub const FORMULA_FLOAT: u32 = 0x0009;
    pub const FORMULA_BOOL: u32 = 0x000A;
    pub const FORMULA_BOOLERR: u32 = 0x000B;

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
pub struct Cell {
    pub row: u32,
    pub col: u32,
    pub style: u32,
    pub value: CellValue,
    /// Raw bytes for the formula tokens / expression payload, if available.
    ///
    /// XLSB stores formulas in a binary encoding; conversion back into the textual A1 form is
    /// intentionally deferred (but these bytes can be preserved for round-tripping).
    pub formula: Option<Vec<u8>>,
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

    pub fn read_record<'a>(&mut self, buf: &'a mut Vec<u8>) -> Result<Option<Biff12Record<'a>>, Error> {
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
        let mut v: u32 = 0;
        for i in 0..4 {
            let Some(byte) = read_u8_opt(&mut self.inner)? else {
                return Ok(None);
            };
            v |= (byte as u32) << (8 * i);
            if byte & 0x80 == 0 {
                break;
            }
        }
        Ok(Some(v))
    }

    fn read_len(&mut self) -> Result<Option<u32>, Error> {
        let mut v: u32 = 0;
        for i in 0..4 {
            let Some(byte) = read_u8_opt(&mut self.inner)? else {
                return Ok(None);
            };
            v |= ((byte & 0x7F) as u32) << (7 * i);
            if byte & 0x80 == 0 {
                break;
            }
        }
        Ok(Some(v))
    }
}

fn read_u8_opt<R: Read>(mut r: R) -> Result<Option<u8>, Error> {
    let mut buf = [0u8; 1];
    match r.read(&mut buf)? {
        0 => Ok(None),
        _ => Ok(Some(buf[0])),
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

    fn remaining(&self) -> &'a [u8] {
        &self.data[self.offset..]
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
        let byte_len = len_chars
            .checked_mul(2)
            .ok_or(Error::UnexpectedEof)?;
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

pub(crate) fn parse_shared_strings<R: Read>(shared_strings_bin: &mut R) -> Result<Vec<String>, Error> {
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

pub(crate) fn parse_sheet<R: Read>(sheet_bin: &mut R, shared_strings: &[String]) -> Result<SheetData, Error> {
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
            biff12::ROW if in_sheet_data => {
                let mut rr = RecordReader::new(rec.data);
                current_row = Some(rr.read_u32()?);
            }
            biff12::BLANK
            | biff12::NUM
            | biff12::BOOLERR
            | biff12::BOOL
            | biff12::FLOAT
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
                    biff12::STRING => {
                        let idx = rr.read_u32()? as usize;
                        let s = shared_strings.get(idx).cloned().unwrap_or_default();
                        (CellValue::Text(s), None)
                    }
                    biff12::FORMULA_STRING => {
                        let v = rr.read_utf16_string()?;
                        (CellValue::Text(v), Some(rr.remaining().to_vec()))
                    }
                    biff12::FORMULA_FLOAT => {
                        let v = rr.read_f64()?;
                        (CellValue::Number(v), Some(rr.remaining().to_vec()))
                    }
                    biff12::FORMULA_BOOL => {
                        let v = rr.read_u8()? != 0;
                        (CellValue::Bool(v), Some(rr.remaining().to_vec()))
                    }
                    biff12::FORMULA_BOOLERR => {
                        let v = rr.read_u8()?;
                        (CellValue::Error(v), Some(rr.remaining().to_vec()))
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

fn normalize_sheet_target(target: &str) -> String {
    // Relationship targets are typically relative to `xl/`.
    let target = target.trim_start_matches('/');
    format!("xl/{}", target.replace('\\', "/"))
}
