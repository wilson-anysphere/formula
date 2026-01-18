use formula_xlsx::offcrypto::{
    decrypt_aes_cbc_no_padding_in_place, AesCbcDecryptError, AES_BLOCK_SIZE,
};
use sha1::{Digest, Sha1};
use std::cmp::min;
use std::fmt;
use std::io::{Read, Seek, SeekFrom};
use zeroize::{Zeroize, Zeroizing};

const SEGMENT_PLAINTEXT_LEN: u64 = 0x1000;
const SIZE_PREFIX_LEN: u64 = 8;
const AES_BLOCK_LEN: usize = AES_BLOCK_SIZE;

/// Streaming decryptor for a **segmented** `EncryptedPackage` layout (compatibility).
///
/// Note: baseline MS-OFFCRYPTO/ECMA-376 **Standard/CryptoAPI AES** `EncryptedPackage` uses
/// **AES-ECB** (no IV). This reader implements a legacy segmented AES-CBC layout that may be
/// encountered in some producer/test-fixture combinations.
///
/// The underlying stream contains:
/// - `orig_size: u64le` (8 bytes)
/// - AES-CBC ciphertext split into 4096-byte segments, each encrypted independently with an IV
///   derived from the segment index.
///
/// This reader exposes the decrypted plaintext as a `Read + Seek` stream without fully buffering
/// the decrypted package.
pub struct StandardAesEncryptedPackageReader<R> {
    inner: R,
    stream_start: u64,
    key: Zeroizing<Vec<u8>>,
    salt: Vec<u8>,
    orig_size: u64,
    ciphertext_len: u64,
    pos: u64,

    cached_segment_index: Option<u64>,
    cached_plaintext: Vec<u8>,

    pending_error: Option<std::io::Error>,
}

impl<R: Read + Seek> StandardAesEncryptedPackageReader<R> {
    /// Create a new streaming decryptor over an `EncryptedPackage` stream.
    ///
    /// `inner` must be positioned at the beginning of the `EncryptedPackage` stream (the
    /// `orig_size` prefix).
    pub fn new(
        mut inner: R,
        key: impl Into<Vec<u8>>,
        salt: impl Into<Vec<u8>>,
    ) -> std::io::Result<Self> {
        let stream_start = inner.seek(SeekFrom::Current(0))?;

        let mut size_buf = [0u8; SIZE_PREFIX_LEN as usize];
        inner
            .read_exact(&mut size_buf)
            .map_err(|e| truncated("EncryptedPackage size prefix", e))?;

        let ciphertext_start = stream_start.checked_add(SIZE_PREFIX_LEN).ok_or_else(|| {
            std::io::Error::new(
                std::io::ErrorKind::InvalidInput,
                "EncryptedPackage stream start offset overflow",
            )
        })?;
        let end = inner.seek(SeekFrom::End(0))?;
        if end < ciphertext_start {
            return Err(std::io::Error::new(
                std::io::ErrorKind::InvalidData,
                "EncryptedPackage stream is truncated (EOF before ciphertext start)",
            ));
        }
        let ciphertext_len = end - ciphertext_start;

        let orig_size =
            crate::parse_encrypted_package_size_prefix_bytes(size_buf, Some(ciphertext_len));

        // Restore position to the ciphertext start so subsequent reads work as expected.
        inner.seek(SeekFrom::Start(ciphertext_start))?;

        // Ciphertext must be block-aligned for AES-CBC without padding removal.
        if ciphertext_len % (AES_BLOCK_LEN as u64) != 0 {
            return Err(std::io::Error::new(
                std::io::ErrorKind::InvalidData,
                "EncryptedPackage ciphertext length is not a multiple of 16",
            ));
        }

        // Guardrail: the `orig_size` prefix is attacker-controlled; reject inputs where the
        // declared plaintext length cannot fit in the available ciphertext bytes.
        //
        // The ciphertext length must be at least `ceil(orig_size / 16) * 16` bytes for AES-CBC.
        // (We treat this as a hard requirement rather than deferring errors to `read()`: callers
        // cannot recover a valid OOXML ZIP package when the final ciphertext segment is missing.)
        let expected_min_ct_len = expected_min_ciphertext_len(orig_size)?;
        if ciphertext_len < expected_min_ct_len {
            return Err(std::io::Error::new(
                std::io::ErrorKind::InvalidData,
                format!(
                    "EncryptedPackage orig_size {orig_size} is implausibly large for ciphertext length {ciphertext_len}",
                ),
            ));
        }

        Ok(Self {
            inner,
            stream_start,
            key: Zeroizing::new(key.into()),
            salt: salt.into(),
            orig_size,
            ciphertext_len,
            pos: 0,
            cached_segment_index: None,
            cached_plaintext: Vec::new(),
            pending_error: None,
        })
    }

    /// The original (decrypted) package size, from the `u64le` prefix.
    pub fn orig_size(&self) -> u64 {
        self.orig_size
    }

    fn segment_count(&self) -> u64 {
        // Use `div_ceil` to avoid overflow when `orig_size` is near `u64::MAX`.
        self.orig_size.div_ceil(SEGMENT_PLAINTEXT_LEN)
    }

    fn derive_iv(&self, segment_index: u64) -> std::io::Result<[u8; 16]> {
        if segment_index > u32::MAX as u64 {
            return Err(std::io::Error::new(
                std::io::ErrorKind::InvalidInput,
                format!("segment index {segment_index} exceeds u32::MAX"),
            ));
        }
        let mut hasher = Sha1::new();
        hasher.update(&self.salt);
        hasher.update(&(segment_index as u32).to_le_bytes());
        let digest = hasher.finalize();
        let mut iv = [0u8; 16];
        iv.copy_from_slice(&digest[..16]);
        Ok(iv)
    }

    fn ciphertext_offset(&self, segment_index: u64) -> std::io::Result<u64> {
        let seg_offset = segment_index
            .checked_mul(SEGMENT_PLAINTEXT_LEN)
            .ok_or_else(|| {
                std::io::Error::new(std::io::ErrorKind::InvalidInput, "segment offset overflow")
            })?;
        let start = self
            .stream_start
            .checked_add(SIZE_PREFIX_LEN)
            .and_then(|v| v.checked_add(seg_offset))
            .ok_or_else(|| {
                std::io::Error::new(
                    std::io::ErrorKind::InvalidInput,
                    "ciphertext offset overflow",
                )
            })?;
        Ok(start)
    }

    fn load_segment(&mut self, segment_index: u64) -> std::io::Result<()> {
        if self.cached_segment_index == Some(segment_index) {
            return Ok(());
        }

        let seg_count = self.segment_count();
        if seg_count == 0 {
            return Err(std::io::Error::new(
                std::io::ErrorKind::UnexpectedEof,
                "EncryptedPackage has no segments",
            ));
        }
        if segment_index >= seg_count {
            return Err(std::io::Error::new(
                std::io::ErrorKind::InvalidInput,
                format!("segment index {segment_index} out of range (segments={seg_count})"),
            ));
        }

        let cipher_off = self.ciphertext_offset(segment_index)?;
        let is_final = segment_index + 1 == seg_count;

        let cipher_len_u64: u64 = if !is_final {
            SEGMENT_PLAINTEXT_LEN
        } else {
            let seg_off = segment_index
                .checked_mul(SEGMENT_PLAINTEXT_LEN)
                .ok_or_else(|| {
                    std::io::Error::new(std::io::ErrorKind::InvalidInput, "segment offset overflow")
                })?;
            self.ciphertext_len.checked_sub(seg_off).ok_or_else(|| {
                std::io::Error::new(
                    std::io::ErrorKind::InvalidData,
                    "EncryptedPackage stream is truncated (EOF before final segment start)",
                )
            })?
        };
        if cipher_len_u64 % (AES_BLOCK_LEN as u64) != 0 {
            return Err(std::io::Error::new(
                std::io::ErrorKind::InvalidData,
                "final EncryptedPackage ciphertext segment length is not a multiple of 16",
            ));
        }
        let cipher_len: usize = usize::try_from(cipher_len_u64).map_err(|_| {
            std::io::Error::new(
                std::io::ErrorKind::InvalidData,
                "EncryptedPackage ciphertext segment does not fit into platform usize",
            )
        })?;

        self.inner.seek(SeekFrom::Start(cipher_off))?;

        let mut buf = vec![0u8; cipher_len];
        self.inner
            .read_exact(&mut buf)
            .map_err(|e| truncated("EncryptedPackage ciphertext segment", e))?;

        let iv = self.derive_iv(segment_index)?;
        decrypt_aes_cbc_in_place(self.key.as_slice(), &iv, &mut buf)?;

        // Ensure the segment contains enough plaintext bytes to cover the declared orig_size.
        let seg_start = segment_index
            .checked_mul(SEGMENT_PLAINTEXT_LEN)
            .ok_or_else(|| {
                std::io::Error::new(std::io::ErrorKind::InvalidInput, "segment start overflow")
            })?;
        let needed = min(
            SEGMENT_PLAINTEXT_LEN,
            self.orig_size.saturating_sub(seg_start),
        ) as usize;
        if buf.len() < needed {
            return Err(std::io::Error::new(
                std::io::ErrorKind::InvalidData,
                "EncryptedPackage stream is truncated (final segment too short)",
            ));
        }

        self.cached_segment_index = Some(segment_index);
        zeroize_vec_u8_full(&mut self.cached_plaintext);
        self.cached_plaintext = buf;
        Ok(())
    }
}

impl<R> Drop for StandardAesEncryptedPackageReader<R> {
    fn drop(&mut self) {
        zeroize_vec_u8_full(&mut self.key);
        zeroize_vec_u8_full(&mut self.cached_plaintext);
    }
}

impl<R> fmt::Debug for StandardAesEncryptedPackageReader<R> {
    fn fmt(&self, f: &mut fmt::Formatter<'_>) -> fmt::Result {
        f.debug_struct("StandardAesEncryptedPackageReader")
            .field("stream_start", &self.stream_start)
            .field("orig_size", &self.orig_size)
            .field("ciphertext_len", &self.ciphertext_len)
            .field("pos", &self.pos)
            .field("cached_segment_index", &self.cached_segment_index)
            .field("cached_plaintext_len", &self.cached_plaintext.len())
            .field("key_len", &self.key.len())
            .field("salt_len", &self.salt.len())
            .finish()
    }
}

impl<R: Read + Seek> Read for StandardAesEncryptedPackageReader<R> {
    fn read(&mut self, out: &mut [u8]) -> std::io::Result<usize> {
        if let Some(err) = self.pending_error.take() {
            return Err(err);
        }
        if out.is_empty() {
            return Ok(0);
        }
        if self.pos >= self.orig_size {
            return Ok(0);
        }

        let mut written = 0usize;
        while written < out.len() && self.pos < self.orig_size {
            let segment_index = self.pos / SEGMENT_PLAINTEXT_LEN;
            let segment_off = (self.pos % SEGMENT_PLAINTEXT_LEN) as usize;

            if let Err(err) = self.load_segment(segment_index) {
                if written > 0 {
                    // Preserve partial progress: return the bytes we have and surface the error on
                    // the next `read()` call (matching common `Read` adapter behavior).
                    self.pending_error = Some(err);
                    break;
                }
                return Err(err);
            }

            let seg_start = segment_index * SEGMENT_PLAINTEXT_LEN;
            let seg_plain_len = min(SEGMENT_PLAINTEXT_LEN, self.orig_size - seg_start) as usize;
            let available = seg_plain_len.checked_sub(segment_off).ok_or_else(|| {
                std::io::Error::new(
                    std::io::ErrorKind::InvalidInput,
                    "segment offset out of range",
                )
            })?;

            let to_copy = min(available, out.len() - written);
            let dst = out.get_mut(written..written + to_copy).ok_or_else(|| {
                std::io::Error::new(
                    std::io::ErrorKind::InvalidInput,
                    "output slice bounds are inconsistent with bytes to copy",
                )
            })?;
            let src = self
                .cached_plaintext
                .get(segment_off..segment_off + to_copy)
                .ok_or_else(|| {
                    std::io::Error::new(
                        std::io::ErrorKind::InvalidData,
                        "segment cache does not contain the requested plaintext range",
                    )
                })?;
            dst.copy_from_slice(src);
            self.pos += to_copy as u64;
            written += to_copy;
        }

        Ok(written)
    }
}

impl<R: Read + Seek> Seek for StandardAesEncryptedPackageReader<R> {
    fn seek(&mut self, pos: SeekFrom) -> std::io::Result<u64> {
        self.pending_error = None;

        let current = self.pos as i128;
        let end = self.orig_size as i128;
        let next: i128 = match pos {
            SeekFrom::Start(off) => off as i128,
            SeekFrom::End(off) => end.checked_add(off as i128).ok_or_else(|| {
                std::io::Error::new(std::io::ErrorKind::InvalidInput, "seek overflow")
            })?,
            SeekFrom::Current(off) => current.checked_add(off as i128).ok_or_else(|| {
                std::io::Error::new(std::io::ErrorKind::InvalidInput, "seek overflow")
            })?,
        };

        if next < 0 {
            return Err(std::io::Error::new(
                std::io::ErrorKind::InvalidInput,
                "invalid seek to a negative position",
            ));
        }
        let next_u64: u64 = next.try_into().map_err(|_| {
            std::io::Error::new(std::io::ErrorKind::InvalidInput, "seek position overflow")
        })?;

        self.pos = next_u64;
        Ok(self.pos)
    }
}

fn decrypt_aes_cbc_in_place(key: &[u8], iv: &[u8; 16], buf: &mut [u8]) -> std::io::Result<()> {
    decrypt_aes_cbc_no_padding_in_place(key, iv, buf).map_err(|err| {
        let (kind, msg) = match err {
            AesCbcDecryptError::UnsupportedKeyLength(_) => {
                (std::io::ErrorKind::InvalidInput, err.to_string())
            }
            AesCbcDecryptError::InvalidIvLength(_)
            | AesCbcDecryptError::InvalidCiphertextLength(_) => {
                (std::io::ErrorKind::InvalidData, err.to_string())
            }
        };
        std::io::Error::new(kind, msg)
    })
}

fn zeroize_vec_u8_full(buf: &mut Vec<u8>) {
    buf.zeroize();
    // SAFETY: We only write zeros to uninitialized memory; it is valid to treat the spare capacity
    // as a raw byte slice for the purpose of clearing it.
    unsafe {
        let spare = buf.spare_capacity_mut();
        let ptr = spare.as_mut_ptr() as *mut u8;
        let len = spare.len();
        std::slice::from_raw_parts_mut(ptr, len).zeroize();
    }
}

fn expected_min_ciphertext_len(orig_size: u64) -> std::io::Result<u64> {
    if orig_size == 0 {
        return Ok(0);
    }

    let block = AES_BLOCK_LEN as u64;
    let blocks = orig_size
        .checked_add(block - 1)
        .map(|v| v / block)
        .ok_or_else(|| {
            std::io::Error::new(
                std::io::ErrorKind::InvalidData,
                "EncryptedPackage orig_size is too large",
            )
        })?;
    blocks.checked_mul(block).ok_or_else(|| {
        std::io::Error::new(
            std::io::ErrorKind::InvalidData,
            "EncryptedPackage orig_size is too large",
        )
    })
}

fn truncated(context: &'static str, err: std::io::Error) -> std::io::Error {
    if err.kind() == std::io::ErrorKind::UnexpectedEof {
        std::io::Error::new(
            std::io::ErrorKind::InvalidData,
            format!("{context} is truncated"),
        )
    } else {
        err
    }
}

#[cfg(test)]
mod tests {
    use super::*;
    use std::io::Cursor;

    #[test]
    fn reader_debug_redacts_key_and_plaintext_bytes() {
        let key_bytes = b"super_secret_key".to_vec();
        let plaintext_bytes = b"super_secret_plaintext".to_vec();
        let key_debug = format!("{key_bytes:?}");
        let plaintext_debug = format!("{plaintext_bytes:?}");

        let reader: StandardAesEncryptedPackageReader<Cursor<Vec<u8>>> =
            StandardAesEncryptedPackageReader {
                inner: Cursor::new(Vec::new()),
                stream_start: 0,
                key: Zeroizing::new(key_bytes),
                salt: vec![0u8; 16],
                orig_size: 0,
                ciphertext_len: 0,
                pos: 0,
                cached_segment_index: None,
                cached_plaintext: plaintext_bytes,
                pending_error: None,
            };

        let dbg = format!("{reader:?}");
        assert!(
            !dbg.contains(&key_debug),
            "Debug output leaked key bytes: {dbg}"
        );
        assert!(
            !dbg.contains(&plaintext_debug),
            "Debug output leaked plaintext bytes: {dbg}"
        );
    }
}
