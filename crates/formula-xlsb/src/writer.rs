use std::io::{self, Write};

/// Low-level writer for BIFF12 record streams (used by XLSB parts).
///
/// Records are encoded as:
/// - record type: "varint" (Excel-specific; see `Biff12Reader::read_id`)
/// - record length: 7-bit varint (see `Biff12Reader::read_len`)
/// - record payload bytes
pub(crate) struct Biff12Writer<W: Write> {
    inner: W,
}

impl<W: Write> Biff12Writer<W> {
    pub(crate) fn new(inner: W) -> Self {
        Self { inner }
    }

    pub(crate) fn write_record(&mut self, id: u32, payload: &[u8]) -> io::Result<()> {
        self.write_record_header(id, payload.len() as u32)?;
        self.inner.write_all(payload)
    }

    pub(crate) fn write_record_header(&mut self, id: u32, len: u32) -> io::Result<()> {
        write_record_id(&mut self.inner, id)?;
        write_record_len(&mut self.inner, len)?;
        Ok(())
    }

    pub(crate) fn write_raw(&mut self, bytes: &[u8]) -> io::Result<()> {
        self.inner.write_all(bytes)
    }

    pub(crate) fn write_u16(&mut self, v: u16) -> io::Result<()> {
        self.inner.write_all(&v.to_le_bytes())
    }

    pub(crate) fn write_u32(&mut self, v: u32) -> io::Result<()> {
        self.inner.write_all(&v.to_le_bytes())
    }

    pub(crate) fn write_f64(&mut self, v: f64) -> io::Result<()> {
        self.inner.write_all(&v.to_le_bytes())
    }

    pub(crate) fn write_utf16_string(&mut self, s: &str) -> io::Result<()> {
        let len = s.encode_utf16().count();
        self.write_u32(len as u32)?;
        for unit in s.encode_utf16() {
            self.inner.write_all(&unit.to_le_bytes())?;
        }
        Ok(())
    }
}

fn write_record_id<W: Write>(mut w: W, id: u32) -> io::Result<()> {
    // Keep in sync with `Biff12Reader::read_id` (note: 8-bit shifting, not 7-bit).
    let bytes = id.to_le_bytes();

    // Determine the minimal byte width that can represent the value.
    let mut n = 4usize;
    while n > 1 && bytes[n - 1] == 0 {
        n -= 1;
    }

    // The reader stops once it sees a byte with the continuation bit unset. If the most
    // significant byte we need has the continuation bit set, append a zero terminator
    // (when we have space). This keeps the decoded numeric value unchanged.
    if bytes[n - 1] & 0x80 != 0 && n < 4 {
        n += 1;
    }

    for i in 0..n {
        let byte = bytes[i];
        if i < n - 1 && byte & 0x80 == 0 {
            return Err(io::Error::new(
                io::ErrorKind::InvalidInput,
                format!("record id 0x{id:08X} cannot be encoded as a BIFF12 varint id"),
            ));
        }
        w.write_all(&[byte])?;
    }
    Ok(())
}

fn write_record_len<W: Write>(mut w: W, len: u32) -> io::Result<()> {
    // Keep in sync with `Biff12Reader::read_len` (7-bit varint).
    if len >= (1 << 28) {
        return Err(io::Error::new(
            io::ErrorKind::InvalidInput,
            format!("record length {len} exceeds BIFF12 28-bit varint limit"),
        ));
    }

    let mut v = len;
    loop {
        let mut byte = (v & 0x7F) as u8;
        v >>= 7;
        if v != 0 {
            byte |= 0x80;
        }
        w.write_all(&[byte])?;
        if v == 0 {
            break;
        }
    }
    Ok(())
}

