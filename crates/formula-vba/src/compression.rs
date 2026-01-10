use thiserror::Error;

#[derive(Debug, Error)]
pub enum CompressionError {
    #[error("compressed container is empty")]
    Empty,
    #[error("invalid compressed container signature {0:#04x}")]
    InvalidSignature(u8),
    #[error("truncated compressed chunk header")]
    TruncatedChunkHeader,
    #[error("invalid compressed chunk signature bits {0:#x}")]
    InvalidChunkSignature(u16),
    #[error("truncated compressed chunk data")]
    TruncatedChunkData,
    #[error("truncated copy token")]
    TruncatedCopyToken,
    #[error("copy token references data before start of buffer (offset={offset}, out_len={out_len})")]
    BadCopyOffset { offset: usize, out_len: usize },
}

/// Decompress an MS-OVBA "CompressedContainer" into its decompressed bytes.
///
/// This is used for:
/// - `VBA/dir` stream
/// - compressed source portion of module streams (`TextOffset..`)
///
/// Spec reference: MS-OVBA 2.4.1.
pub fn decompress_container(input: &[u8]) -> Result<Vec<u8>, CompressionError> {
    let (&sig, rest) = input.split_first().ok_or(CompressionError::Empty)?;
    if sig != 0x01 {
        return Err(CompressionError::InvalidSignature(sig));
    }

    let mut out = Vec::new();
    let mut offset = 0usize;
    while offset < rest.len() {
        if offset + 2 > rest.len() {
            return Err(CompressionError::TruncatedChunkHeader);
        }

        let header = u16::from_le_bytes([rest[offset], rest[offset + 1]]);
        offset += 2;

        // bits 12..14 must be 0b011.
        let signature_bits = (header & 0x7000) >> 12;
        if signature_bits != 0b011 {
            return Err(CompressionError::InvalidChunkSignature(signature_bits));
        }

        let compressed = (header & 0x8000) != 0;
        let chunk_size = (header & 0x0FFF) as usize + 3; // includes header
        let chunk_data_size = chunk_size
            .checked_sub(2)
            .expect("chunk_size >=2 by construction");

        if offset + chunk_data_size > rest.len() {
            return Err(CompressionError::TruncatedChunkData);
        }
        let chunk_data = &rest[offset..offset + chunk_data_size];
        offset += chunk_data_size;

        if compressed {
            out.extend_from_slice(&decompress_chunk(chunk_data)?);
        } else {
            out.extend_from_slice(chunk_data);
        }
    }

    Ok(out)
}

fn decompress_chunk(chunk: &[u8]) -> Result<Vec<u8>, CompressionError> {
    let mut out: Vec<u8> = Vec::with_capacity(4096);
    let mut idx = 0usize;

    while idx < chunk.len() && out.len() < 4096 {
        let flags = chunk[idx];
        idx += 1;
        for bit in 0..8 {
            if idx >= chunk.len() {
                break;
            }
            if out.len() >= 4096 {
                break;
            }

            if flags & (1 << bit) == 0 {
                // literal
                out.push(chunk[idx]);
                idx += 1;
                continue;
            }

            // copy token (2 bytes)
            if idx + 2 > chunk.len() {
                return Err(CompressionError::TruncatedCopyToken);
            }
            let token = u16::from_le_bytes([chunk[idx], chunk[idx + 1]]);
            idx += 2;

            let bit_count = copy_token_bit_count(out.len());
            let length_bit_count = 16 - bit_count;
            let length_mask: u16 = (1u16 << length_bit_count) - 1;

            let offset_raw = (token >> length_bit_count) as usize;
            let length_raw = (token & length_mask) as usize;

            let offset = offset_raw + 1;
            let length = length_raw + 3;

            if offset == 0 || offset > out.len() {
                return Err(CompressionError::BadCopyOffset {
                    offset,
                    out_len: out.len(),
                });
            }

            for _ in 0..length {
                if out.len() >= 4096 {
                    break;
                }
                let src_index = out.len() - offset;
                let byte = out[src_index];
                out.push(byte);
            }
        }
    }

    Ok(out)
}

/// Compute CopyTokenBitCount per MS-OVBA.
///
/// It is derived from the current decompressed size (within a 4KiB chunk).
fn copy_token_bit_count(current_decompressed_len: usize) -> u32 {
    // Per spec the value is based on (current - 1), with a minimum of 4.
    let n = current_decompressed_len.saturating_sub(1);
    let bits_needed = if n == 0 {
        0
    } else {
        (usize::BITS - n.leading_zeros()) as u32
    };
    bits_needed.clamp(4, 12)
}
