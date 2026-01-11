use std::io::{self, Read, Write};

const MAX_RECORD_ID_BYTES: usize = 4;
const MAX_RECORD_LEN_BYTES: usize = 4;
const MAX_RECORD_LEN: u32 = 0x0FFF_FFFF;

fn unexpected_eof(context: &'static str) -> io::Error {
    io::Error::new(io::ErrorKind::UnexpectedEof, context)
}

/// Read a BIFF12 record ID from `r`.
///
/// BIFF12 record IDs use a continuation scheme that is *not* standard LEB128:
/// bytes are interpreted as a little-endian integer, and the high bit of each byte
/// is also used as the continuation flag.
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
        v |= (byte as u32) << (8 * i);
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
/// This mirrors [`read_record_id`]. Not all `u32` values are representable in this
/// encoding; values that cannot be encoded losslessly return `InvalidInput`.
pub fn write_record_id(w: &mut impl Write, id: u32) -> io::Result<()> {
    let bytes = id.to_le_bytes();

    // Determine the number of bytes we must write to satisfy the continuation-bit rule.
    // We keep emitting bytes while the previously written byte has its continuation bit set.
    let mut n = 1usize;
    while n < MAX_RECORD_ID_BYTES && (bytes[n - 1] & 0x80) != 0 {
        n += 1;
    }

    // If we had to use all 4 bytes, the final byte must still have MSB=0 or the reader
    // would expect a 5th byte, which is not allowed by BIFF12.
    if n == MAX_RECORD_ID_BYTES && (bytes[3] & 0x80) != 0 {
        return Err(io::Error::new(
            io::ErrorKind::InvalidInput,
            "BIFF12 record id requires more than 4 bytes",
        ));
    }

    // Any remaining high bytes must be zero, otherwise we'd truncate the value.
    if bytes[n..].iter().any(|&b| b != 0) {
        return Err(io::Error::new(
            io::ErrorKind::InvalidInput,
            "BIFF12 record id cannot be encoded without truncation",
        ));
    }

    w.write_all(&bytes[..n])
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
