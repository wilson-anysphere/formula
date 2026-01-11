use std::io::{self, Write};

use crate::biff12_varint;

/// Low-level writer for BIFF12 record streams (used by XLSB parts).
///
/// Records are encoded as:
/// - record type: Excel-specific varint (see [`crate::biff12_varint::write_record_id`])
/// - record length: 7-bit varint (see [`crate::biff12_varint::write_record_len`])
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
        biff12_varint::write_record_id(&mut self.inner, id)?;
        biff12_varint::write_record_len(&mut self.inner, len)?;
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
