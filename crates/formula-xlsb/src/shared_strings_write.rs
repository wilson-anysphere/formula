use std::collections::HashMap;
use std::io::{self, Cursor};

use crate::biff12_varint;
use crate::parser::{biff12, Error};

#[derive(Debug, Clone, Copy)]
struct RecordRange {
    start: usize,
    payload_start: usize,
    payload_end: usize,
    end: usize,
}

/// Shared strings (workbook shared strings part, typically `xl/sharedStrings.bin`) patcher that
/// preserves existing records byte-for-byte.
///
/// The shared string table is a BIFF12 record stream containing:
/// - `BrtSST` (`0x009F`) header with `[totalCount:u32][uniqueCount:u32]`
/// - `BrtSI` (`0x0013`) entries
/// - `BrtSSTEnd` (`0x00A0`)
///
/// This patcher supports interning *plain* strings (no rich text runs / phonetic data) by:
/// - reusing an existing `BrtSI` entry when the plain text matches
/// - otherwise appending a new `BrtSI` record with `flags=0`
///
/// All existing records (including unknown ones) are copied byte-for-byte, except for the first
/// 8 bytes of the `BrtSST` payload which are updated to reflect the new counts.
pub struct SharedStringsWriter {
    bytes: Vec<u8>,
    records: Vec<RecordRange>,
    sst_record_idx: usize,
    /// Total number of worksheet cells that reference the shared string table (`cstTotal` /
    /// `count` in XLSX).
    sst_total_count: u32,
    /// Total number of unique string items present in the table (`cstUnique` / `uniqueCount` in
    /// XLSX).
    sst_unique_count: u32,
    original_sst_total_count: u32,
    original_sst_unique_count: u32,
    insert_record_idx: usize,
    plain_to_index: HashMap<String, u32>,
    base_si_count: u32,
    appended_plain: Vec<String>,
}

impl SharedStringsWriter {
    pub fn new(bytes: Vec<u8>) -> Result<Self, Error> {
        let mut cursor = Cursor::new(&bytes);
        let mut records = Vec::new();

        let mut sst_record_idx: Option<usize> = None;
        let mut sst_end_record_idx: Option<usize> = None;
        let mut sst_total_count: u32 = 0;
        let mut sst_unique_count: u32 = 0;

        let mut seen_sst_end = false;
        let mut last_si_record_idx: Option<usize> = None;
        let mut plain_to_index: HashMap<String, u32> = HashMap::new();
        let mut si_index: u32 = 0;

        loop {
            let start = cursor.position() as usize;
            let Some(id) = biff12_varint::read_record_id(&mut cursor).map_err(map_io_error)? else {
                break;
            };
            let Some(len) = biff12_varint::read_record_len(&mut cursor).map_err(map_io_error)?
            else {
                return Err(Error::UnexpectedEof);
            };
            let len = len as usize;

            let header_end = cursor.position() as usize;
            let payload_start = header_end;
            let payload_end = payload_start
                .checked_add(len)
                .filter(|&end| end <= bytes.len())
                .ok_or(Error::UnexpectedEof)?;
            cursor.set_position(payload_end as u64);

            let record_idx = records.len();
            records.push(RecordRange {
                start,
                payload_start,
                payload_end,
                end: payload_end,
            });

            match id {
                biff12::SST => {
                    sst_record_idx = Some(record_idx);
                    let payload = bytes
                        .get(payload_start..payload_end)
                        .ok_or(Error::UnexpectedEof)?;
                    if payload.len() < 8 {
                        return Err(Error::UnexpectedEof);
                    }
                    // BrtSST: [cstTotal: u32][cstUnique: u32] (matches BIFF8 + XLSX `sst` attrs).
                    sst_total_count = u32::from_le_bytes(payload[0..4].try_into().unwrap());
                    sst_unique_count = u32::from_le_bytes(payload[4..8].try_into().unwrap());
                }
                biff12::SI if !seen_sst_end => {
                    if let Some(text) = parse_plain_si_text(&bytes[payload_start..payload_end]) {
                        plain_to_index.entry(text).or_insert(si_index);
                    }
                    last_si_record_idx = Some(record_idx);
                    si_index = si_index.saturating_add(1);
                }
                biff12::SST_END => {
                    seen_sst_end = true;
                    if sst_end_record_idx.is_none() {
                        sst_end_record_idx = Some(record_idx);
                    }
                }
                _ => {}
            }
        }

        let Some(sst_record_idx) = sst_record_idx else {
            return Err(Error::Io(io::Error::new(
                io::ErrorKind::InvalidData,
                "invalid sharedStrings.bin: missing BrtSST record",
            )));
        };
        let Some(sst_end_record_idx) = sst_end_record_idx else {
            return Err(Error::Io(io::Error::new(
                io::ErrorKind::InvalidData,
                "invalid sharedStrings.bin: missing BrtSSTEnd record",
            )));
        };

        let insert_after = last_si_record_idx.unwrap_or(sst_record_idx);
        let insert_record_idx = (insert_after + 1).min(sst_end_record_idx);

        Ok(Self {
            bytes,
            records,
            sst_record_idx,
            sst_total_count,
            sst_unique_count,
            original_sst_total_count: sst_total_count,
            original_sst_unique_count: sst_unique_count,
            insert_record_idx,
            plain_to_index,
            base_si_count: si_index,
            appended_plain: Vec::new(),
        })
    }

    /// Intern a plain shared string and return its `isst` index.
    pub fn intern_plain(&mut self, s: &str) -> Result<u32, Error> {
        if let Some(&idx) = self.plain_to_index.get(s) {
            return Ok(idx);
        }

        let idx = self
            .base_si_count
            .checked_add(self.appended_plain.len() as u32)
            .ok_or(Error::UnexpectedEof)?;
        self.plain_to_index.insert(s.to_string(), idx);
        self.appended_plain.push(s.to_string());

        // New `BrtSI` record -> the table has one more unique string item.
        let expected_unique_count = self
            .base_si_count
            .checked_add(self.appended_plain.len() as u32)
            .ok_or(Error::UnexpectedEof)?;
        self.sst_unique_count = expected_unique_count;

        Ok(idx)
    }

    /// Adjust the `BrtSST` total reference count (`cstTotal`) by a signed delta.
    pub fn note_total_ref_delta(&mut self, delta: i64) -> Result<(), Error> {
        if delta == 0 {
            return Ok(());
        }

        let current = self.sst_total_count as i64;
        let updated = current
            .checked_add(delta)
            .ok_or_else(|| {
                Error::Io(io::Error::new(
                    io::ErrorKind::InvalidInput,
                    "sharedStrings totalCount overflow",
                ))
            })?;
        if updated < 0 || updated > u32::MAX as i64 {
            return Err(Error::Io(io::Error::new(
                io::ErrorKind::InvalidInput,
                "sharedStrings totalCount out of range",
            )));
        }
        self.sst_total_count = updated as u32;

        // Ensure `uniqueCount` stays consistent with the number of `BrtSI` records we will write.
        self.sst_unique_count = self
            .base_si_count
            .checked_add(self.appended_plain.len() as u32)
            .ok_or(Error::UnexpectedEof)?;
        Ok(())
    }

    pub fn into_bytes(self) -> Result<Vec<u8>, Error> {
        if self.appended_plain.is_empty()
            && self.sst_total_count == self.original_sst_total_count
            && self.sst_unique_count == self.original_sst_unique_count
        {
            return Ok(self.bytes);
        }

        let mut out = Vec::with_capacity(
            self.bytes
                .len()
                .saturating_add(self.appended_plain.len().saturating_mul(16)),
        );

        for (idx, rec) in self.records.iter().enumerate() {
            if idx == self.insert_record_idx {
                write_appended_si_records(&mut out, &self.appended_plain)?;
            }

            if idx == self.sst_record_idx {
                out.extend_from_slice(
                    self.bytes
                        .get(rec.start..rec.payload_start)
                        .ok_or(Error::UnexpectedEof)?,
                );
                let payload = self
                    .bytes
                    .get(rec.payload_start..rec.payload_end)
                    .ok_or(Error::UnexpectedEof)?;
                if payload.len() < 8 {
                    return Err(Error::UnexpectedEof);
                }
                let mut patched = payload.to_vec();
                patched[0..4].copy_from_slice(&self.sst_total_count.to_le_bytes());
                patched[4..8].copy_from_slice(&self.sst_unique_count.to_le_bytes());
                out.extend_from_slice(&patched);
            } else {
                out.extend_from_slice(
                    self.bytes
                        .get(rec.start..rec.end)
                        .ok_or(Error::UnexpectedEof)?,
                );
            }
        }

        if self.insert_record_idx == self.records.len() {
            write_appended_si_records(&mut out, &self.appended_plain)?;
        }

        Ok(out)
    }
}

fn parse_plain_si_text(payload: &[u8]) -> Option<String> {
    // BrtSI payload:
    //   [flags: u8][text: XLWideString]
    // Rich/phonetic data are only present if corresponding flag bits are set.
    let flags = *payload.first()?;
    if flags != 0 {
        return None;
    }

    if payload.len() < 1 + 4 {
        return None;
    }
    let cch = u32::from_le_bytes(payload[1..5].try_into().ok()?) as usize;
    let byte_len = cch.checked_mul(2)?;
    let expected_len = 1usize.checked_add(4)?.checked_add(byte_len)?;
    // Some writers include benign trailing bytes after the UTF-16 text even when `flags==0`.
    // Be tolerant and only require that the declared UTF-16 bytes are present.
    if payload.len() < expected_len {
        return None;
    }

    let raw = payload.get(5..expected_len)?;
    let mut units = Vec::with_capacity(cch);
    for chunk in raw.chunks_exact(2) {
        units.push(u16::from_le_bytes([chunk[0], chunk[1]]));
    }
    Some(String::from_utf16_lossy(&units))
}

fn write_appended_si_records(out: &mut Vec<u8>, strings: &[String]) -> Result<(), Error> {
    for s in strings {
        write_plain_si_record(out, s)?;
    }
    Ok(())
}

fn write_plain_si_record(out: &mut Vec<u8>, s: &str) -> Result<(), Error> {
    // BrtSI payload: [flags: u8][text: XLWideString]
    // XLWideString: [cch: u32][utf16 chars...]
    let cch = s.encode_utf16().count();
    let cch = u32::try_from(cch).map_err(|_| {
        Error::Io(io::Error::new(
            io::ErrorKind::InvalidInput,
            "string is too large for sharedStrings.bin",
        ))
    })?;
    let byte_len = cch.checked_mul(2).ok_or(Error::UnexpectedEof)?;
    let payload_len = 1u32
        .checked_add(4)
        .and_then(|v| v.checked_add(byte_len))
        .ok_or(Error::UnexpectedEof)?;

    biff12_varint::write_record_id(out, biff12::SI).map_err(map_io_error)?;
    biff12_varint::write_record_len(out, payload_len).map_err(map_io_error)?;
    out.push(0u8);
    out.extend_from_slice(&cch.to_le_bytes());
    for unit in s.encode_utf16() {
        out.extend_from_slice(&unit.to_le_bytes());
    }
    Ok(())
}

fn map_io_error(err: io::Error) -> Error {
    if err.kind() == io::ErrorKind::UnexpectedEof {
        Error::UnexpectedEof
    } else {
        Error::Io(err)
    }
}
