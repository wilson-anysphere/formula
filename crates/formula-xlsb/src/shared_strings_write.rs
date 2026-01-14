use std::collections::HashMap;
use std::io::{self, Cursor, Read, Write};

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
        let updated = current.checked_add(delta).ok_or_else(|| {
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

/// Streaming shared strings patcher for `xl/sharedStrings.bin`.
///
/// This is a lower-level alternative to [`SharedStringsWriter`] that does not require
/// materializing the entire part in memory. Existing records (including record ID/length varint
/// bytes) are copied byte-for-byte, except:
/// - the first 8 bytes of the `BrtSST` payload (`[totalCount:u32][uniqueCount:u32]`) are patched,
/// - new plain `BrtSI` records are inserted immediately after the last existing `BrtSI` record
///   (or after `BrtSST` when there are no entries), before the first `BrtSSTEnd`.
///
/// `base_si_count` is the number of existing `BrtSI` records present in the original table
/// (i.e. the expected `uniqueCount`). Callers can compute this from the parsed shared strings
/// table without scanning the binary part.
pub struct SharedStringsWriterStreaming;

#[derive(Debug)]
struct RawVarint {
    buf: [u8; 4],
    len: u8,
}

impl RawVarint {
    fn as_slice(&self) -> &[u8] {
        &self.buf[..self.len as usize]
    }
}

#[derive(Debug)]
struct RawRecordHeader {
    id: u32,
    id_raw: RawVarint,
    len: u32,
    len_raw: RawVarint,
}

impl SharedStringsWriterStreaming {
    /// Patch a shared string table record stream.
    ///
    /// Returns `Ok(true)` when the output differs from the input stream, and `Ok(false)` when
    /// the output is byte-identical.
    pub fn patch<R: Read, W: Write>(
        mut input: R,
        mut output: W,
        new_plain_strings: &[String],
        base_si_count: u32,
        total_ref_delta: i64,
    ) -> Result<bool, Error> {
        if new_plain_strings.is_empty() && total_ref_delta == 0 {
            // Fast path: byte-for-byte passthrough (also preserves non-canonical varint headers).
            io::copy(&mut input, &mut output).map_err(map_io_error)?;
            return Ok(false);
        }

        let updated_unique_count = base_si_count
            .checked_add(u32::try_from(new_plain_strings.len()).map_err(|_| Error::UnexpectedEof)?)
            .ok_or(Error::UnexpectedEof)?;

        let mut seen_sst = false;
        let mut seen_sst_end = false;
        // Buffer records that occur after the last observed BrtSI and before BrtSSTEnd.
        //
        // When we reach BrtSSTEnd we need to insert new BrtSI entries *after* the final existing
        // BrtSI record, but *before* any trailing records. We cannot know which BrtSI is the final
        // one until we hit BrtSSTEnd, so we opportunistically buffer records until another BrtSI
        // proves they're not in the suffix.
        let mut tail: Vec<u8> = Vec::new();

        while let Some(header) = read_record_header(&mut input)? {
            match header.id {
                biff12::SST => {
                    seen_sst = true;
                    write_raw_header(&mut output, &header)?;

                    let len = header.len as usize;
                    if len < 8 {
                        return Err(Error::UnexpectedEof);
                    }

                    // BrtSST payload prefix: [totalCount:u32][uniqueCount:u32]
                    let mut prefix = [0u8; 8];
                    input.read_exact(&mut prefix).map_err(map_io_error)?;

                    let original_total =
                        u32::from_le_bytes(prefix[0..4].try_into().expect("u32 bytes"));
                    let updated_total = apply_total_ref_delta(original_total, total_ref_delta)?;

                    prefix[0..4].copy_from_slice(&updated_total.to_le_bytes());
                    prefix[4..8].copy_from_slice(&updated_unique_count.to_le_bytes());

                    output.write_all(&prefix).map_err(map_io_error)?;
                    copy_exact(&mut input, &mut output, len.saturating_sub(prefix.len()))?;
                }
                biff12::SI if seen_sst && !seen_sst_end => {
                    // We just encountered another BrtSI, so any buffered records must belong
                    // before it (they're not part of the suffix after the final BrtSI).
                    if !tail.is_empty() {
                        output.write_all(&tail).map_err(map_io_error)?;
                        tail.clear();
                    }
                    write_raw_header(&mut output, &header)?;
                    copy_exact(&mut input, &mut output, header.len as usize)?;
                }
                biff12::SST_END if seen_sst && !seen_sst_end => {
                    // We reached the end of the table. Insert new BrtSI records after the last
                    // existing BrtSI and before any trailing records and BrtSSTEnd.
                    write_appended_si_records(&mut output, new_plain_strings)?;
                    if !tail.is_empty() {
                        output.write_all(&tail).map_err(map_io_error)?;
                        tail.clear();
                    }

                    seen_sst_end = true;
                    write_raw_header(&mut output, &header)?;
                    copy_exact(&mut input, &mut output, header.len as usize)?;
                }
                _ if seen_sst && !seen_sst_end => {
                    // Defer writing the record until we know whether it belongs after the final
                    // BrtSI. If we see another BrtSI later, we'll flush this buffer before that
                    // record. Otherwise, we'll flush it after inserting appended BrtSI entries.
                    tail.extend_from_slice(header.id_raw.as_slice());
                    tail.extend_from_slice(header.len_raw.as_slice());
                    copy_exact(&mut input, &mut tail, header.len as usize)?;
                }
                _ => {
                    // Outside the shared string table, we can copy records verbatim.
                    write_raw_header(&mut output, &header)?;
                    copy_exact(&mut input, &mut output, header.len as usize)?;
                }
            }
        }

        if !seen_sst {
            return Err(Error::Io(io::Error::new(
                io::ErrorKind::InvalidData,
                "invalid sharedStrings.bin: missing BrtSST record",
            )));
        }
        if !seen_sst_end {
            return Err(Error::Io(io::Error::new(
                io::ErrorKind::InvalidData,
                "invalid sharedStrings.bin: missing BrtSSTEnd record",
            )));
        }

        Ok(true)
    }
}

fn parse_plain_si_text(payload: &[u8]) -> Option<String> {
    let utf16_end = reusable_plain_si_utf16_end(payload)?;
    let _cch = u32::from_le_bytes(payload.get(1..5)?.try_into().ok()?) as usize;
    let raw = payload.get(5..utf16_end)?;

    // Avoid allocating an intermediate `Vec<u16>` for attacker-controlled strings; decode
    // UTF-16LE directly into a `String`.
    let mut out = String::with_capacity(raw.len());
    let iter = raw
        .chunks_exact(2)
        .map(|chunk| u16::from_le_bytes([chunk[0], chunk[1]]));
    for decoded in std::char::decode_utf16(iter) {
        match decoded {
            Ok(ch) => out.push(ch),
            Err(_) => out.push('\u{FFFD}'),
        }
    }
    Some(out)
}

/// Return the end offset (in bytes) of the UTF-16 text payload for a reusable "plain" `BrtSI`.
///
/// This treats strings as reusable when:
/// - `flags & 0x03 == 0` (even if reserved bits and/or benign trailing bytes exist), or
/// - only the rich/phonetic flag bits are set and the corresponding blocks are empty
///   (`cRun == 0` / `cb == 0`), with no trailing bytes.
pub(crate) fn reusable_plain_si_utf16_end(payload: &[u8]) -> Option<usize> {
    // BrtSI payload:
    //   [flags: u8][text: XLWideString]
    // Rich/phonetic data are only present if corresponding flag bits are set.
    let flags = *payload.first()?;

    if payload.len() < 1 + 4 {
        return None;
    }
    let cch = u32::from_le_bytes(payload.get(1..5)?.try_into().ok()?) as usize;
    let byte_len = cch.checked_mul(2)?;
    let utf16_end = 1usize.checked_add(4)?.checked_add(byte_len)?;
    if payload.len() < utf16_end {
        return None;
    }

    // Plain string: no rich/phonetic blocks. Be tolerant of reserved bits in `flags` and trailing
    // bytes after the UTF-16 payload.
    if flags & 0x03 == 0 {
        return Some(utf16_end);
    }

    // Some real-world XLSB writers set the rich/phonetic bits in BrtSI flags even when the
    // corresponding blocks are present-but-empty (cRun=0 / cb=0). These strings are effectively
    // plain and can be safely reused when interning by text.
    //
    // Be conservative: only treat the string as plain if the flags contain *only* the rich /
    // phonetic bits and the record payload is exactly the base text followed by the empty block
    // headers (no run/phonetic bytes and no trailing data).
    if flags & !0x03 != 0 {
        return None;
    }

    let mut offset = utf16_end;
    if flags & 0x01 != 0 {
        let c_run = u32::from_le_bytes(payload.get(offset..offset + 4)?.try_into().ok()?);
        if c_run != 0 {
            return None;
        }
        offset = offset.checked_add(4)?;
    }
    if flags & 0x02 != 0 {
        let cb = u32::from_le_bytes(payload.get(offset..offset + 4)?.try_into().ok()?);
        if cb != 0 {
            return None;
        }
        offset = offset.checked_add(4)?;
    }

    if offset != payload.len() {
        return None;
    }

    Some(utf16_end)
}

fn write_appended_si_records(out: &mut impl Write, strings: &[String]) -> Result<(), Error> {
    for s in strings {
        write_plain_si_record(out, s)?;
    }
    Ok(())
}

fn write_plain_si_record(out: &mut impl Write, s: &str) -> Result<(), Error> {
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
    out.write_all(&[0u8]).map_err(map_io_error)?;
    out.write_all(&cch.to_le_bytes()).map_err(map_io_error)?;
    for unit in s.encode_utf16() {
        out.write_all(&unit.to_le_bytes()).map_err(map_io_error)?;
    }
    Ok(())
}

fn apply_total_ref_delta(original: u32, delta: i64) -> Result<u32, Error> {
    if delta == 0 {
        return Ok(original);
    }

    let current = original as i64;
    let updated = current.checked_add(delta).ok_or_else(|| {
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
    Ok(updated as u32)
}

fn read_record_header<R: Read>(r: &mut R) -> Result<Option<RawRecordHeader>, Error> {
    let Some((id, id_raw)) = read_record_id_raw(r)? else {
        return Ok(None);
    };
    let (len, len_raw) = read_record_len_raw(r)?;
    Ok(Some(RawRecordHeader {
        id,
        id_raw,
        len,
        len_raw,
    }))
}

fn read_record_id_raw<R: Read>(r: &mut R) -> Result<Option<(u32, RawVarint)>, Error> {
    let mut v: u32 = 0;
    let mut raw = RawVarint {
        buf: [0u8; 4],
        len: 0,
    };

    for i in 0..4 {
        let mut buf = [0u8; 1];
        let n = r.read(&mut buf).map_err(map_io_error)?;
        if n == 0 {
            if i == 0 {
                return Ok(None);
            }
            return Err(Error::UnexpectedEof);
        }

        let byte = buf[0];
        raw.buf[i] = byte;
        raw.len = raw.len.saturating_add(1);
        v |= ((byte & 0x7F) as u32) << (7 * i);
        if byte & 0x80 == 0 {
            return Ok(Some((v, raw)));
        }
    }

    Err(Error::Io(io::Error::new(
        io::ErrorKind::InvalidData,
        "invalid BIFF12 record id (more than 4 bytes)",
    )))
}

fn read_record_len_raw<R: Read>(r: &mut R) -> Result<(u32, RawVarint), Error> {
    let mut v: u32 = 0;
    let mut raw = RawVarint {
        buf: [0u8; 4],
        len: 0,
    };

    for i in 0..4 {
        let mut buf = [0u8; 1];
        let n = r.read(&mut buf).map_err(map_io_error)?;
        if n == 0 {
            return Err(Error::UnexpectedEof);
        }

        let byte = buf[0];
        raw.buf[i] = byte;
        raw.len = raw.len.saturating_add(1);
        v |= ((byte & 0x7F) as u32) << (7 * i);
        if byte & 0x80 == 0 {
            return Ok((v, raw));
        }
    }

    Err(Error::Io(io::Error::new(
        io::ErrorKind::InvalidData,
        "invalid BIFF12 record length (more than 4 bytes)",
    )))
}

fn write_raw_header<W: Write>(w: &mut W, header: &RawRecordHeader) -> Result<(), Error> {
    w.write_all(header.id_raw.as_slice())
        .map_err(map_io_error)?;
    w.write_all(header.len_raw.as_slice())
        .map_err(map_io_error)?;
    Ok(())
}

fn copy_exact<R: Read, W: Write>(
    input: &mut R,
    output: &mut W,
    mut len: usize,
) -> Result<(), Error> {
    let mut buf = [0u8; 16 * 1024];
    while len > 0 {
        let chunk_len = buf.len().min(len);
        input
            .read_exact(&mut buf[..chunk_len])
            .map_err(map_io_error)?;
        output.write_all(&buf[..chunk_len]).map_err(map_io_error)?;
        len = len.saturating_sub(chunk_len);
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
