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

/// Compress a decompressed byte stream into an MS-OVBA "CompressedContainer".
///
/// The resulting container can be decompressed with [`decompress_container`].
///
/// Notes:
/// - Compression happens independently per 4KiB chunk (MS-OVBA requirement).
/// - If a compressed representation isn't smaller (or doesn't fit in 4KiB),
///   the chunk is emitted as an *uncompressed* chunk (still within a
///   `CompressedContainer`).
pub fn compress_container(input: &[u8]) -> Vec<u8> {
    let mut out = Vec::new();
    let _ = out.try_reserve(input.len().saturating_add(3));
    out.push(0x01);

    for chunk in input.chunks(4096) {
        if chunk.is_empty() {
            continue;
        }

        let compressed = compress_chunk(chunk);
        if let Some(compressed) = compressed {
            if compressed.len() < chunk.len() && compressed.len() <= 4096 {
                write_chunk_header(&mut out, true, compressed.len());
                out.extend_from_slice(&compressed);
                continue;
            }
        }

        write_chunk_header(&mut out, false, chunk.len());
        out.extend_from_slice(chunk);
    }

    out
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
        let hdr_end = offset
            .checked_add(2)
            .ok_or(CompressionError::TruncatedChunkHeader)?;
        let Some(hdr) = rest.get(offset..hdr_end) else {
            return Err(CompressionError::TruncatedChunkHeader);
        };

        let header = u16::from_le_bytes([hdr[0], hdr[1]]);
        offset += 2;

        // bits 12..14 must be 0b011.
        let signature_bits = (header & 0x7000) >> 12;
        if signature_bits != 0b011 {
            return Err(CompressionError::InvalidChunkSignature(signature_bits));
        }

        let compressed = (header & 0x8000) != 0;
        let chunk_size = (header & 0x0FFF) as usize + 3; // includes header
        // `chunk_size` includes the 2-byte header and is always >= 3.
        let chunk_data_size = chunk_size - 2;

        let Some(chunk_end) = offset.checked_add(chunk_data_size) else {
            return Err(CompressionError::TruncatedChunkData);
        };
        if chunk_end > rest.len() {
            return Err(CompressionError::TruncatedChunkData);
        }
        let chunk_data = rest
            .get(offset..chunk_end)
            .ok_or(CompressionError::TruncatedChunkData)?;
        offset = chunk_end;

        if compressed {
            out.extend_from_slice(&decompress_chunk(chunk_data)?);
        } else {
            out.extend_from_slice(chunk_data);
        }
    }

    Ok(out)
}

fn write_chunk_header(out: &mut Vec<u8>, compressed: bool, chunk_data_len: usize) {
    // CompressedChunkSize = chunk_size - 3, chunk_size = header(2) + chunk_data_len.
    // => size_field = chunk_data_len - 1.
    let size_field = (chunk_data_len - 1) as u16;
    let header = if compressed { 0xB000 } else { 0x3000 } | (size_field & 0x0FFF);
    out.extend_from_slice(&header.to_le_bytes());
}

fn compress_chunk(chunk: &[u8]) -> Option<Vec<u8>> {
    let mut out = Vec::new();
    let mut pos = 0usize;

    while pos < chunk.len() {
        let flags_pos = out.len();
        out.push(0); // placeholder
        let mut flags: u8 = 0;

        for bit in 0..8 {
            if pos >= chunk.len() {
                break;
            }

            let current_out_len = pos;
            let bit_count = copy_token_bit_count(current_out_len);
            let max_offset = (1usize << bit_count).min(current_out_len);
            let length_bit_count = 16 - bit_count;
            let max_len = ((1usize << length_bit_count) + 2)
                .min(chunk.len().saturating_sub(pos));

            let (best_offset, best_len) = find_best_match(chunk, pos, max_offset, max_len);
            if best_len >= 3 {
                let length_raw = (best_len - 3) as u16;
                let offset_raw = (best_offset - 1) as u16;
                let token = (offset_raw << length_bit_count) | length_raw;
                out.extend_from_slice(&token.to_le_bytes());
                flags |= 1 << bit;
                pos += best_len;
            } else {
                out.push(chunk[pos]);
                pos += 1;
            }

            if out.len() > 4096 {
                return None;
            }
        }

        out[flags_pos] = flags;
    }

    Some(out)
}

fn find_best_match(
    chunk: &[u8],
    pos: usize,
    max_offset: usize,
    max_len: usize,
) -> (usize, usize) {
    if max_offset == 0 || max_len < 3 {
        return (0, 0);
    }

    let mut best_offset = 0usize;
    let mut best_len = 0usize;

    // Naive search across the already-emitted prefix. The chunk is capped at 4KiB so this is
    // acceptable for our current use-cases (UI display / future rewrites).
    for offset in 1..=max_offset {
        let start = pos - offset;
        let mut len = 0usize;
        while len < max_len && chunk[start + len] == chunk[pos + len] {
            len += 1;
            if pos + len >= chunk.len() {
                break;
            }
        }

        if len > best_len {
            best_len = len;
            best_offset = offset;
            if best_len == max_len {
                break;
            }
        }
    }

    (best_offset, best_len)
}

fn decompress_chunk(chunk: &[u8]) -> Result<Vec<u8>, CompressionError> {
    let mut out: Vec<u8> = Vec::new();
    let _ = out.try_reserve_exact(4096);
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
            let end = idx
                .checked_add(2)
                .ok_or(CompressionError::TruncatedCopyToken)?;
            let bytes = chunk
                .get(idx..end)
                .ok_or(CompressionError::TruncatedCopyToken)?;
            let token = u16::from_le_bytes([bytes[0], bytes[1]]);
            idx = end;

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
        usize::BITS - n.leading_zeros()
    };
    bits_needed.clamp(4, 12)
}
