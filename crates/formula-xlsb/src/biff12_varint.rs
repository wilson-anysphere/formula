use std::io::{self, Read, Write};

const MAX_RECORD_ID_BYTES: usize = 4;
const MAX_RECORD_LEN_BYTES: usize = 4;
const MAX_RECORD_ID: u32 = 0x0FFF_FFFF;
const MAX_RECORD_LEN: u32 = 0x0FFF_FFFF;

fn unexpected_eof(context: &'static str) -> io::Error {
    io::Error::new(io::ErrorKind::UnexpectedEof, context)
}

/// Read a BIFF12 record ID from `r`.
///
/// Record IDs are encoded as a 7-bit varint (LEB128-like) using up to 4 bytes.
///
/// Returns `Ok(None)` when `r` is at EOF before reading any bytes.
pub fn read_record_id(r: &mut impl Read) -> io::Result<Option<u32>> {
    let mut v: u32 = 0;
    for i in 0..MAX_RECORD_ID_BYTES {
        let mut buf = [0u8; 1];
        match r.read(&mut buf)? {
            0 if i == 0 => return Ok(None),
            0 => {
                return Err(unexpected_eof(
                    "unexpected EOF while reading BIFF12 record id",
                ))
            }
            _ => {}
        }

        let byte = buf[0];
        v |= ((byte & 0x7F) as u32) << (7 * i);
        if byte & 0x80 == 0 {
            return Ok(Some(v));
        }
    }

    Err(io::Error::new(
        io::ErrorKind::InvalidData,
        "invalid BIFF12 record id (more than 4 bytes)",
    ))
}

/// Write a BIFF12 record ID to `w`.
///
/// Record IDs are encoded as a 7-bit varint (LEB128-like) using up to 4 bytes.
pub fn write_record_id(w: &mut impl Write, id: u32) -> io::Result<()> {
    if id > MAX_RECORD_ID {
        return Err(io::Error::new(
            io::ErrorKind::InvalidInput,
            "BIFF12 record id exceeds 28-bit varint encoding",
        ));
    }

    let mut id = id;
    loop {
        let mut byte = (id & 0x7F) as u8;
        id >>= 7;
        if id != 0 {
            byte |= 0x80;
        }
        w.write_all(&[byte])?;
        if id == 0 {
            return Ok(());
        }
    }
}

/// Read a BIFF12 record payload length from `r`.
///
/// Record lengths are encoded as a 7-bit varint (LEB128-like) using up to 4 bytes.
///
/// Returns `Ok(None)` when `r` is at EOF before reading any bytes.
pub fn read_record_len(r: &mut impl Read) -> io::Result<Option<u32>> {
    let mut v: u32 = 0;
    for i in 0..MAX_RECORD_LEN_BYTES {
        let mut buf = [0u8; 1];
        match r.read(&mut buf)? {
            0 if i == 0 => return Ok(None),
            0 => {
                return Err(unexpected_eof(
                    "unexpected EOF while reading BIFF12 record length",
                ))
            }
            _ => {}
        }

        let byte = buf[0];
        v |= ((byte & 0x7F) as u32) << (7 * i);
        if byte & 0x80 == 0 {
            return Ok(Some(v));
        }
    }

    Err(io::Error::new(
        io::ErrorKind::InvalidData,
        "invalid BIFF12 record length (more than 4 bytes)",
    ))
}

/// Write a BIFF12 record payload length to `w`.
///
/// Lengths are encoded as a 7-bit varint (LEB128-like) using up to 4 bytes.
pub fn write_record_len(w: &mut impl Write, mut len: u32) -> io::Result<()> {
    if len > MAX_RECORD_LEN {
        return Err(io::Error::new(
            io::ErrorKind::InvalidInput,
            "BIFF12 record length exceeds 28-bit varint encoding",
        ));
    }

    loop {
        let mut byte = (len & 0x7F) as u8;
        len >>= 7;
        if len != 0 {
            byte |= 0x80;
        }
        w.write_all(&[byte])?;
        if len == 0 {
            return Ok(());
        }
    }
}
