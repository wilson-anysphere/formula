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

impl<'a> Biff12Writer<&'a mut Vec<u8>> {
    pub(crate) fn bytes_written(&self) -> usize {
        self.inner.len()
    }
}

impl<W: Write> Biff12Writer<W> {
    pub(crate) fn new(inner: W) -> Self {
        Self { inner }
    }

    pub(crate) fn write_record(&mut self, id: u32, payload: &[u8]) -> io::Result<()> {
        let len = u32::try_from(payload.len()).map_err(|_| {
            io::Error::new(
                io::ErrorKind::InvalidInput,
                "record payload length does not fit in u32",
            )
        })?;
        self.write_record_header(id, len)?;
        self.inner.write_all(payload)
    }

    pub(crate) fn write_record_header(&mut self, id: u32, len: u32) -> io::Result<()> {
        // Encode into a temporary buffer first so invalid ids/lengths don't leave the output
        // stream with a partially written header.
        let mut header = Vec::new();
        let _ = header.try_reserve_exact(8);
        biff12_varint::write_record_id(&mut header, id)?;
        biff12_varint::write_record_len(&mut header, len)?;
        self.inner.write_all(&header)
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
        let len = u32::try_from(len)
            .map_err(|_| io::Error::new(io::ErrorKind::InvalidInput, "string is too large"))?;
        self.write_u32(len)?;
        for unit in s.encode_utf16() {
            self.inner.write_all(&unit.to_le_bytes())?;
        }
        Ok(())
    }
}
