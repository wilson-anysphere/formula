//! MS-OFFCRYPTO Standard / CryptoAPI RC4 `EncryptedPackage` decryptor.
//!
//! This module implements the **Standard Encryption (CryptoAPI)** RC4 variant used by
//! password-to-open OOXML files stored inside an OLE/CFB container (`EncryptionInfo` +
//! `EncryptedPackage` streams).
//!
//! See `docs/offcrypto-standard-cryptoapi-rc4.md` for a from-scratch writeup of:
//! - `EncryptedPackage` stream framing (`u64le` plaintext size prefix + ciphertext),
//! - RC4 re-keying every **0x200 bytes** (not the legacy BIFF8 RC4 0x400-byte interval), and
//! - per-block key derivation (UTF-16LE password + 50,000 spin loop + `LE32(block)`).
//!
use std::fmt;
use std::io::{Read, Seek, SeekFrom};

pub use crate::offcrypto::cryptoapi::HashAlg;
use crate::offcrypto::cryptoapi::{CALG_MD5, CALG_SHA1};
use md5::Md5;
use sha1::{Digest as _, Sha1};
use thiserror::Error;
use zeroize::{Zeroize, Zeroizing};

const ENCRYPTED_PACKAGE_SIZE_PREFIX_LEN: usize = 8;
const RC4_BLOCK_SIZE: usize = 0x200;

/// Errors returned when validating/constructing an RC4 CryptoAPI `EncryptedPackage` decryptor.
#[derive(Debug, Error)]
pub enum Rc4CryptoApiEncryptedPackageError {
    #[error(
        "`EncryptedPackage` stream is too short: expected at least {ENCRYPTED_PACKAGE_SIZE_PREFIX_LEN} bytes, got {len}"
    )]
    StreamTooShort { len: u64 },

    #[error(
        "`EncryptedPackage` package_size {package_size} exceeds ciphertext length {ciphertext_len}"
    )]
    PackageSizeExceedsCiphertext {
        package_size: u64,
        ciphertext_len: u64,
    },

    #[error(
        "invalid `EncryptionHeader.keySize` {key_size_bits} bits: must be a multiple of 8 (0 is interpreted as 40-bit)"
    )]
    InvalidKeySizeBits { key_size_bits: u32 },

    #[error(
        "unsupported `EncryptionHeader.keySize` {key_size_bits} bits (supported: 0/40, 56, 128 for RC4 CryptoAPI)"
    )]
    UnsupportedKeySizeBits { key_size_bits: u32 },

    #[error(
        "`EncryptionHeader.keySize` {key_size_bits} bits implies key_len {key_len} bytes, but algIdHash {alg_id_hash:#010x} digest length is {hash_len} bytes"
    )]
    KeySizeTooLargeForHash {
        key_size_bits: u32,
        key_len: u32,
        alg_id_hash: u32,
        hash_len: usize,
    },

    #[error(
        "unsupported `EncryptionHeader.algIdHash` {alg_id_hash:#010x} (supported: CALG_SHA1=0x00008004, CALG_MD5=0x00008003)"
    )]
    UnsupportedHashAlgorithm { alg_id_hash: u32 },

    #[error("`EncryptedPackage` stream offset arithmetic overflow while {context}")]
    OffsetOverflow { context: &'static str },

    #[error("I/O error while reading `EncryptedPackage`: {0}")]
    Io(#[from] std::io::Error),
}

/// Streaming decryptor for Standard/CryptoAPI RC4 `EncryptedPackage` (MS-OFFCRYPTO).
///
/// This reader exposes the decrypted plaintext bytes of the encrypted package without buffering the
/// whole ZIP payload in memory.
///
/// ## Block restart behavior
///
/// The `EncryptedPackage` payload is split into 0x200-byte blocks. Each block is encrypted with a
/// fresh RC4 key derived as:
///
/// `key_material_b = Hash(H || LE32(b))[0..key_len]`
/// where:
/// - `H` is the base hash bytes (typically `Hfinal`; 20 bytes for SHA-1 or 16 bytes for MD5).
/// - `b` is the 0-based block index.
/// - `key_len = keySize/8` (40→5 bytes, 56→7 bytes, 128→16 bytes). MS-OFFCRYPTO specifies that for
///   RC4, `keySize == 0` MUST be interpreted as 40-bit.
/// - For 40-bit RC4 (`key_len == 5`), use the 5-byte key directly (do not pad to 16 bytes).
///
/// Some other Office encryption formats (notably legacy BIFF8 `FILEPASS` CryptoAPI RC4) treat 40-bit
/// keys as a 16-byte RC4 key blob with the high 88 bits set to zero. That changes the RC4 KSA
/// because RC4 depends on both key bytes *and* key length, and it is incorrect for
/// Standard/CryptoAPI `EncryptedPackage`.
///
/// Seeking is supported by re-deriving the block key and discarding `o = pos % 0x200` bytes of
/// RC4 keystream.
pub struct Rc4CryptoApiDecryptReader<R: Read + Seek> {
    inner: R,

    /// Absolute stream offset of the first encrypted byte (i.e. after the 8-byte package size
    /// prefix in `EncryptedPackage`).
    ciphertext_start: u64,
    /// Current inner position *relative* to `ciphertext_start`.
    inner_pos: u64,

    /// Total plaintext size (from the 8-byte `EncryptedPackage` length prefix).
    package_size: u64,
    /// Current plaintext offset.
    pos: u64,

    /// Base hash bytes used for per-block key derivation.
    h: Zeroizing<Vec<u8>>,
    /// RC4 key length in bytes (e.g. `keySize / 8` from EncryptionHeader).
    key_len: usize,
    hash_alg: HashAlg,
    /// Compatibility toggle for legacy 40-bit RC4 key blobs.
    ///
    /// MS-OFFCRYPTO Standard RC4 uses the raw 5-byte key material when `keySize == 0/40` (i.e.
    /// `key_len == 5`). However, some legacy CryptoAPI implementations treat a 40-bit RC4 key as a
    /// 16-byte key blob where the remaining 11 bytes are zero (which changes the RC4 KSA because it
    /// depends on key length).
    ///
    /// When constructed via [`Self::from_encrypted_package_stream`], we attempt to auto-detect this
    /// quirk by decrypting the first 4 ciphertext bytes and checking for a ZIP `PK..` signature.
    pad_40_bit_to_128: bool,

    rc4: Option<Rc4>,
    block_index: Option<u32>,
    /// Offset within the current block that `rc4` is aligned to.
    block_offset: usize,
}

impl<R: Read + Seek> fmt::Debug for Rc4CryptoApiDecryptReader<R> {
    fn fmt(&self, f: &mut fmt::Formatter<'_>) -> fmt::Result {
        // Redact any secret/key material: `h` is effectively a key for the RC4 keystream derivation.
        f.debug_struct("Rc4CryptoApiDecryptReader")
            .field("ciphertext_start", &self.ciphertext_start)
            .field("package_size", &self.package_size)
            .field("pos", &self.pos)
            .field("h_len", &self.h.len())
            .field("key_len", &self.key_len)
            .field("hash_alg", &self.hash_alg)
            .field("pad_40_bit_to_128", &self.pad_40_bit_to_128)
            .field("block_index", &self.block_index)
            .field("block_offset", &self.block_offset)
            .field("rc4_initialized", &self.rc4.is_some())
            .finish()
    }
}

impl<R: Read + Seek> Rc4CryptoApiDecryptReader<R> {
    /// Create a decrypting reader from the start of an `EncryptedPackage` stream.
    ///
    /// This helper validates the `EncryptedPackage` framing:
    /// - The stream is at least 8 bytes long (the `package_size` prefix).
    /// - `package_size` does not claim more bytes than the available ciphertext.
    ///
    /// It also validates CryptoAPI parameters from the corresponding Standard `EncryptionHeader`:
    /// - `alg_id_hash` must be SHA-1 (`CALG_SHA1`) or MD5 (`CALG_MD5`).
    /// - `key_size_bits` must be one of 40/56/128 bits and fit within the hash digest length.
    ///
    /// Note: Some producers pad the ciphertext to a block boundary, so ciphertext may be longer than
    /// `package_size`. We treat `package_size > ciphertext_len` as a malformed/truncated stream and
    /// fail fast (rather than attempting reads that would end in EOF).
    pub fn from_encrypted_package_stream(
        mut inner: R,
        h: Vec<u8>,
        key_size_bits: u32,
        alg_id_hash: u32,
    ) -> Result<Self, Rc4CryptoApiEncryptedPackageError> {
        let start_pos = inner.stream_position()?;
        let end_pos = inner.seek(SeekFrom::End(0))?;
        let stream_len = end_pos.checked_sub(start_pos).ok_or(
            Rc4CryptoApiEncryptedPackageError::OffsetOverflow {
                context: "computing stream length",
            },
        )?;

        if stream_len < ENCRYPTED_PACKAGE_SIZE_PREFIX_LEN as u64 {
            return Err(Rc4CryptoApiEncryptedPackageError::StreamTooShort { len: stream_len });
        }

        // Restore the stream to the start of the EncryptedPackage header.
        inner.seek(SeekFrom::Start(start_pos))?;

        let mut size_bytes = [0u8; ENCRYPTED_PACKAGE_SIZE_PREFIX_LEN];
        inner.read_exact(&mut size_bytes)?;

        let ciphertext_len = stream_len
            .checked_sub(ENCRYPTED_PACKAGE_SIZE_PREFIX_LEN as u64)
            .ok_or(Rc4CryptoApiEncryptedPackageError::OffsetOverflow {
                context: "computing ciphertext length",
            })?;

        // MS-OFFCRYPTO describes this prefix as a `u64le`, but some producers treat it as
        // `(u32 size, u32 reserved)` and may write a non-zero reserved high DWORD. Use the shared
        // size-prefix parser to apply a ciphertext-length plausibility check and fall back to the
        // low DWORD in that case.
        let package_size =
            crate::parse_encrypted_package_size_prefix_bytes(size_bytes, Some(ciphertext_len));

        if package_size > ciphertext_len {
            return Err(
                Rc4CryptoApiEncryptedPackageError::PackageSizeExceedsCiphertext {
                    package_size,
                    ciphertext_len,
                },
            );
        }

        let hash_alg = match alg_id_hash {
            CALG_SHA1 => HashAlg::Sha1,
            CALG_MD5 => HashAlg::Md5,
            _ => {
                return Err(
                    Rc4CryptoApiEncryptedPackageError::UnsupportedHashAlgorithm { alg_id_hash },
                )
            }
        };

        // MS-OFFCRYPTO specifies that `keySize=0` MUST be interpreted as 40-bit (legacy "strong"
        // encryption export restrictions).
        let key_size_bits = if key_size_bits == 0 {
            40
        } else {
            key_size_bits
        };
        if key_size_bits % 8 != 0 {
            return Err(Rc4CryptoApiEncryptedPackageError::InvalidKeySizeBits { key_size_bits });
        }
        if !matches!(key_size_bits, 40 | 56 | 128) {
            // If we decide to accept a broader set in the future, keep the "key_len <= hashLen"
            // guardrails below.
            return Err(Rc4CryptoApiEncryptedPackageError::UnsupportedKeySizeBits {
                key_size_bits,
            });
        }

        let key_len_u32 = key_size_bits / 8;
        let hash_len = hash_alg.hash_len();
        if key_len_u32 as usize > hash_len {
            return Err(Rc4CryptoApiEncryptedPackageError::KeySizeTooLargeForHash {
                key_size_bits,
                key_len: key_len_u32,
                alg_id_hash,
                hash_len,
            });
        }

        // Ensure `inner` is positioned at the ciphertext start for the streaming decryptor.
        let ciphertext_start = start_pos
            .checked_add(ENCRYPTED_PACKAGE_SIZE_PREFIX_LEN as u64)
            .ok_or(Rc4CryptoApiEncryptedPackageError::OffsetOverflow {
                context: "computing ciphertext start offset",
            })?;
        inner.seek(SeekFrom::Start(ciphertext_start))?;

        let key_len = key_len_u32 as usize;

        // Compatibility: Some producers zero-pad 40-bit key material to a 16-byte RC4 key blob.
        // Since `EncryptedPackage` plaintext should be a ZIP container, we can usually detect this
        // by checking whether the decrypted prefix looks like `PK..`.
        let mut pad_40_bit_to_128 = false;
        if key_len == 5 && package_size >= 4 {
            let mut ciphertext_prefix = [0u8; 4];
            inner.read_exact(&mut ciphertext_prefix)?;
            inner.seek(SeekFrom::Start(ciphertext_start))?;

            fn is_zip_sig(prefix: &[u8; 4]) -> bool {
                matches!(
                    prefix,
                    b"PK\x03\x04" | b"PK\x05\x06" | b"PK\x07\x08"
                )
            }

            let digest0 = Zeroizing::new(match hash_alg {
                HashAlg::Sha1 => {
                    let mut hasher = Sha1::new();
                    hasher.update(&h);
                    hasher.update(0u32.to_le_bytes());
                    hasher.finalize().to_vec()
                }
                HashAlg::Md5 => {
                    let mut hasher = Md5::new();
                    hasher.update(&h);
                    hasher.update(0u32.to_le_bytes());
                    hasher.finalize().to_vec()
                }
            });

            // Spec-correct: raw 5-byte key.
            let mut plain_unpadded = ciphertext_prefix;
            let mut rc4 = Rc4::new(&digest0[..5]);
            rc4.apply_keystream(&mut plain_unpadded);

            // Compatibility: 16-byte key blob `key_material || 0x00 * 11`.
            let mut padded_key = Zeroizing::new([0u8; 16]);
            padded_key[..5].copy_from_slice(&digest0[..5]);
            let mut plain_padded = ciphertext_prefix;
            let mut rc4 = Rc4::new(padded_key.as_slice());
            rc4.apply_keystream(&mut plain_padded);

            let unpadded_zip = is_zip_sig(&plain_unpadded);
            let padded_zip = is_zip_sig(&plain_padded);
            if padded_zip && !unpadded_zip {
                pad_40_bit_to_128 = true;
            }
        }

        let mut reader = Self::new_with_hash_alg(inner, package_size, h, key_len, hash_alg)?;
        reader.pad_40_bit_to_128 = pad_40_bit_to_128;
        Ok(reader)
    }

    /// Create a decrypting reader wrapping `inner`.
    ///
    /// `inner` must be positioned at the start of the ciphertext payload (i.e. just after reading
    /// the 8-byte `package_size` prefix from the `EncryptedPackage` stream).
    ///
    /// This constructor assumes `algIdHash = CALG_SHA1` for per-block key derivation. Prefer
    /// [`Self::from_encrypted_package_stream`] when parsing real Office files so that the `algIdHash`
    /// and `keySize` parameters are validated.
    pub fn new(inner: R, package_size: u64, h: Vec<u8>, key_len: usize) -> std::io::Result<Self> {
        Self::new_with_hash_alg(inner, package_size, h, key_len, HashAlg::Sha1)
    }

    pub fn new_with_hash_alg(
        mut inner: R,
        package_size: u64,
        h: Vec<u8>,
        key_len: usize,
        hash_alg: HashAlg,
    ) -> std::io::Result<Self> {
        if key_len == 0 {
            return Err(std::io::Error::new(
                std::io::ErrorKind::InvalidInput,
                "RC4 key_len must be non-zero",
            ));
        }

        // RC4 key bytes are derived by hashing and truncating, so key_len cannot exceed the hash
        // digest size.
        let hash_len = hash_alg.hash_len();
        if key_len > hash_len {
            return Err(std::io::Error::new(
                std::io::ErrorKind::InvalidInput,
                format!("RC4 key_len {key_len} exceeds hash digest size {hash_len}"),
            ));
        }

        let ciphertext_start = inner.stream_position()?;
        Ok(Self {
            inner,
            ciphertext_start,
            inner_pos: 0,
            package_size,
            pos: 0,
            h: Zeroizing::new(h),
            key_len,
            hash_alg,
            pad_40_bit_to_128: false,
            rc4: None,
            block_index: None,
            block_offset: 0,
        })
    }

    /// Return the plaintext length of the encrypted package.
    pub fn package_size(&self) -> u64 {
        self.package_size
    }

    /// Consume the wrapper and return the underlying reader.
    pub fn into_inner(self) -> R {
        self.inner
    }

    fn ensure_inner_position(&mut self) -> std::io::Result<()> {
        if self.inner_pos == self.pos {
            return Ok(());
        }
        let abs = self.ciphertext_start.checked_add(self.pos).ok_or_else(|| {
            std::io::Error::new(
                std::io::ErrorKind::InvalidInput,
                "ciphertext offset overflow while seeking",
            )
        })?;
        self.inner.seek(SeekFrom::Start(abs))?;
        self.inner_pos = self.pos;
        Ok(())
    }

    fn ensure_block(&mut self) -> std::io::Result<()> {
        let block_index_u64 = self.pos / RC4_BLOCK_SIZE as u64;
        let block_index = u32::try_from(block_index_u64).map_err(|_| {
            std::io::Error::new(
                std::io::ErrorKind::InvalidInput,
                "encrypted package position exceeds 32-bit block address space",
            )
        })?;
        let offset = (self.pos % RC4_BLOCK_SIZE as u64) as usize;

        if self.block_index == Some(block_index) && self.block_offset == offset {
            return Ok(());
        }

        // Derive per-block RC4 key: Hash(H || LE32(block_index)) truncated to key_len.
        //
        // This digest (and the derived key bytes) are sensitive; wrap in `Zeroizing` so the
        // intermediate is wiped even if we early-return.
        let digest = Zeroizing::new(match self.hash_alg {
            HashAlg::Sha1 => {
                let mut hasher = Sha1::new();
                hasher.update(self.h.as_slice());
                hasher.update(block_index.to_le_bytes());
                hasher.finalize().to_vec()
            }
            HashAlg::Md5 => {
                let mut hasher = Md5::new();
                hasher.update(self.h.as_slice());
                hasher.update(block_index.to_le_bytes());
                hasher.finalize().to_vec()
            }
        });

        let mut rc4 = if self.pad_40_bit_to_128 && self.key_len == 5 {
            // Compatibility: treat a 40-bit key as a 16-byte key blob with the high 88 bits zero.
            let mut padded_key = Zeroizing::new([0u8; 16]);
            padded_key[..5].copy_from_slice(&digest[..5]);
            Rc4::new(padded_key.as_slice())
        } else {
            // Spec-correct: raw digest truncation.
            Rc4::new(&digest[..self.key_len])
        };
        // Drop the `Zeroizing` wrapper early so derived key material is wiped as soon as we've
        // initialized the RC4 state.
        drop(digest);
        rc4.skip(offset);

        self.rc4 = Some(rc4);
        self.block_index = Some(block_index);
        self.block_offset = offset;
        Ok(())
    }

    fn invalidate_cipher_state(&mut self) {
        self.rc4 = None;
        self.block_index = None;
        self.block_offset = 0;
    }
}

impl<R: Read + Seek> Read for Rc4CryptoApiDecryptReader<R> {
    fn read(&mut self, buf: &mut [u8]) -> std::io::Result<usize> {
        if buf.is_empty() {
            return Ok(0);
        }
        if self.pos >= self.package_size {
            return Ok(0);
        }

        let remaining_u64 = self.package_size - self.pos;
        let mut remaining = usize::try_from(remaining_u64).unwrap_or(usize::MAX);
        remaining = remaining.min(buf.len());

        let mut written = 0usize;
        while remaining > 0 {
            self.ensure_block()?;
            self.ensure_inner_position()?;

            let in_block_offset = (self.pos % RC4_BLOCK_SIZE as u64) as usize;
            let block_remaining = RC4_BLOCK_SIZE - in_block_offset;
            let chunk_len = remaining.min(block_remaining);

            let out = &mut buf[written..written + chunk_len];
            let n = self.inner.read(out)?;
            if n == 0 {
                if written == 0 {
                    return Err(std::io::Error::new(
                        std::io::ErrorKind::UnexpectedEof,
                        "unexpected EOF while reading EncryptedPackage ciphertext",
                    ));
                }
                break;
            }
            self.inner_pos = self.inner_pos.checked_add(n as u64).ok_or_else(|| {
                std::io::Error::new(
                    std::io::ErrorKind::InvalidInput,
                    "inner_pos overflow while reading EncryptedPackage ciphertext",
                )
            })?;

            self.rc4
                .as_mut()
                .ok_or_else(|| {
                    std::io::Error::new(
                        std::io::ErrorKind::Other,
                        "internal error: RC4 state missing",
                    )
                })?
                .apply_keystream(&mut out[..n]);

            self.pos = self.pos.checked_add(n as u64).ok_or_else(|| {
                std::io::Error::new(
                    std::io::ErrorKind::InvalidInput,
                    "plaintext position overflow while reading EncryptedPackage",
                )
            })?;
            self.block_offset = self.block_offset.checked_add(n).ok_or_else(|| {
                std::io::Error::new(
                    std::io::ErrorKind::InvalidInput,
                    "block offset overflow while reading EncryptedPackage",
                )
            })?;

            written += n;
            remaining -= n;

            // If the read ended early (common for some Read impls), return what we have.
            if n < chunk_len {
                break;
            }

            // Move to next block when we've fully consumed this one.
            if self.block_offset >= RC4_BLOCK_SIZE {
                self.invalidate_cipher_state();
            }
        }

        Ok(written)
    }
}

impl<R: Read + Seek> Seek for Rc4CryptoApiDecryptReader<R> {
    fn seek(&mut self, pos: SeekFrom) -> std::io::Result<u64> {
        let base: i128 = match pos {
            SeekFrom::Start(n) => n as i128,
            SeekFrom::Current(off) => self.pos as i128 + off as i128,
            SeekFrom::End(off) => self.package_size as i128 + off as i128,
        };
        if base < 0 {
            return Err(std::io::Error::new(
                std::io::ErrorKind::InvalidInput,
                "invalid seek to a negative position",
            ));
        }
        // Clamp to the plaintext EOF boundary. This avoids seeking the underlying stream past the
        // meaningful ciphertext range while still satisfying "seek beyond EOF behaves like EOF".
        //
        // Use an i128 comparison to avoid `as u64` wraparound for very large positive offsets.
        let new_pos = if base > self.package_size as i128 {
            self.package_size
        } else {
            base as u64
        };

        self.pos = new_pos;
        self.invalidate_cipher_state();

        let abs = self.ciphertext_start.checked_add(self.pos).ok_or_else(|| {
            std::io::Error::new(
                std::io::ErrorKind::InvalidInput,
                "ciphertext offset overflow while seeking",
            )
        })?;
        self.inner.seek(SeekFrom::Start(abs))?;
        self.inner_pos = self.pos;

        Ok(self.pos)
    }
}

#[derive(Clone)]
struct Rc4 {
    s: [u8; 256],
    i: u8,
    j: u8,
}

impl Drop for Rc4 {
    fn drop(&mut self) {
        self.s.zeroize();
        self.i.zeroize();
        self.j.zeroize();
    }
}

impl std::fmt::Debug for Rc4 {
    fn fmt(&self, f: &mut std::fmt::Formatter<'_>) -> std::fmt::Result {
        // Avoid dumping the full internal permutation in debug output.
        f.debug_struct("Rc4")
            .field("i", &self.i)
            .field("j", &self.j)
            .finish()
    }
}

impl Rc4 {
    fn new(key: &[u8]) -> Self {
        assert!(!key.is_empty(), "RC4 key must be non-empty");

        let mut s = [0u8; 256];
        for (i, v) in s.iter_mut().enumerate() {
            *v = i as u8;
        }

        let mut j: u8 = 0;
        for i in 0..256u16 {
            let si = s[i as usize];
            j = j.wrapping_add(si).wrapping_add(key[i as usize % key.len()]);
            s.swap(i as usize, j as usize);
        }

        Self { s, i: 0, j: 0 }
    }

    fn next_byte(&mut self) -> u8 {
        self.i = self.i.wrapping_add(1);
        self.j = self.j.wrapping_add(self.s[self.i as usize]);
        self.s.swap(self.i as usize, self.j as usize);
        let idx = self.s[self.i as usize].wrapping_add(self.s[self.j as usize]);
        self.s[idx as usize]
    }

    fn apply_keystream(&mut self, data: &mut [u8]) {
        for b in data {
            *b ^= self.next_byte();
        }
    }

    fn skip(&mut self, n: usize) {
        for _ in 0..n {
            let _ = self.next_byte();
        }
    }
}

#[cfg(test)]
mod tests {
    use super::*;
    use std::io::Cursor;

    #[test]
    fn reader_debug_redacts_hash_bytes() {
        let h_bytes = b"super_secret_h".to_vec();
        let h_debug = format!("{h_bytes:?}");

        let reader: Rc4CryptoApiDecryptReader<Cursor<Vec<u8>>> = Rc4CryptoApiDecryptReader {
            inner: Cursor::new(Vec::new()),
            ciphertext_start: 0,
            inner_pos: 0,
            package_size: 0,
            pos: 0,
            h: Zeroizing::new(h_bytes),
            key_len: 5,
            hash_alg: HashAlg::Sha1,
            pad_40_bit_to_128: false,
            rc4: None,
            block_index: None,
            block_offset: 0,
        };

        let dbg = format!("{reader:?}");
        assert!(
            !dbg.contains(&h_debug),
            "Debug output leaked key material: {dbg}"
        );
    }

    fn encrypt_rc4_cryptoapi(
        plaintext: &[u8],
        h: &[u8],
        key_len: usize,
        hash_alg: HashAlg,
    ) -> Vec<u8> {
        assert!(key_len <= hash_alg.hash_len());

        let mut out = vec![0u8; plaintext.len()];
        let mut offset = 0usize;
        let mut block_index = 0u32;
        while offset < plaintext.len() {
            let digest = match hash_alg {
                HashAlg::Sha1 => {
                    let mut hasher = Sha1::new();
                    hasher.update(h);
                    hasher.update(block_index.to_le_bytes());
                    hasher.finalize().to_vec()
                }
                HashAlg::Md5 => {
                    let mut hasher = Md5::new();
                    hasher.update(h);
                    hasher.update(block_index.to_le_bytes());
                    hasher.finalize().to_vec()
                }
            };
            // Derive per-block RC4 key: Hash(H || LE32(block_index)) truncated to key_len.
            //
            // For 40-bit RC4 (`keySize == 0`/`40` → `key_len == 5`), the key is the first 5 bytes of
            // the digest. Do not pad to 16 bytes: RC4's key schedule depends on the key length.
            let key = &digest[..key_len];
            let mut rc4 = Rc4::new(key);

            let block_len = (plaintext.len() - offset).min(RC4_BLOCK_SIZE);
            out[offset..offset + block_len].copy_from_slice(&plaintext[offset..offset + block_len]);
            rc4.apply_keystream(&mut out[offset..offset + block_len]);

            offset += block_len;
            block_index += 1;
        }
        out
    }

    fn encrypt_rc4_cryptoapi_padded_40_bit_key(
        plaintext: &[u8],
        h: &[u8],
        key_len: usize,
        hash_alg: HashAlg,
    ) -> Vec<u8> {
        assert!(key_len <= hash_alg.hash_len());

        let mut out = vec![0u8; plaintext.len()];
        let mut offset = 0usize;
        let mut block_index = 0u32;
        while offset < plaintext.len() {
            let digest = match hash_alg {
                HashAlg::Sha1 => {
                    let mut hasher = Sha1::new();
                    hasher.update(h);
                    hasher.update(block_index.to_le_bytes());
                    hasher.finalize().to_vec()
                }
                HashAlg::Md5 => {
                    let mut hasher = Md5::new();
                    hasher.update(h);
                    hasher.update(block_index.to_le_bytes());
                    hasher.finalize().to_vec()
                }
            };

            let mut rc4 = if key_len == 5 {
                // Compatibility behavior: treat a 40-bit key as a 16-byte key blob
                // `key_material || 0x00 * 11`.
                let mut padded_key = [0u8; 16];
                padded_key[..5].copy_from_slice(&digest[..5]);
                Rc4::new(&padded_key)
            } else {
                Rc4::new(&digest[..key_len])
            };

            let block_len = (plaintext.len() - offset).min(RC4_BLOCK_SIZE);
            out[offset..offset + block_len].copy_from_slice(&plaintext[offset..offset + block_len]);
            rc4.apply_keystream(&mut out[offset..offset + block_len]);

            offset += block_len;
            block_index += 1;
        }
        out
    }

    #[test]
    fn sequential_reads_across_block_boundary() {
        let h = b"0123456789ABCDEFGHIJ".to_vec(); // 20 bytes
        for key_len in [5usize, 7, 16] {
            // Ensure plaintext crosses a 0x200 boundary.
            let mut plaintext = vec![0u8; RC4_BLOCK_SIZE + 64];
            for (i, b) in plaintext.iter_mut().enumerate() {
                *b = (i % 251) as u8;
            }
            let ciphertext = encrypt_rc4_cryptoapi(&plaintext, &h, key_len, HashAlg::Sha1);

            // Simulate EncryptedPackage stream layout: [u64 package_size] + ciphertext.
            let mut stream = Vec::new();
            stream.extend_from_slice(&(plaintext.len() as u64).to_le_bytes());
            stream.extend_from_slice(&ciphertext);

            let mut cursor = Cursor::new(stream);
            cursor.seek(SeekFrom::Start(8)).unwrap();

            let mut reader =
                Rc4CryptoApiDecryptReader::new(cursor, plaintext.len() as u64, h.clone(), key_len)
                    .unwrap();

            let mut out = vec![0u8; plaintext.len()];
            // Read in small chunks to force multiple calls and cross-block behavior.
            let out_len = out.len();
            let mut read = 0usize;
            while read < out_len {
                let end = read + 33.min(out_len - read);
                let n = reader.read(&mut out[read..end]).unwrap();
                assert!(n > 0, "unexpected EOF while reading (key_len={key_len})");
                read += n;
            }
            assert_eq!(out, plaintext, "round-trip mismatch (key_len={key_len})");
        }
    }

    #[test]
    fn encrypted_package_round_trip_with_56_bit_key_via_header_keysize() {
        let h = b"0123456789ABCDEFGHIJ".to_vec(); // 20 bytes
        let key_len = 7; // 56-bit

        // Ensure plaintext crosses a 0x200 boundary.
        let mut plaintext = vec![0u8; RC4_BLOCK_SIZE + 64];
        for (i, b) in plaintext.iter_mut().enumerate() {
            *b = (i % 251) as u8;
        }
        let ciphertext = encrypt_rc4_cryptoapi(&plaintext, &h, key_len, HashAlg::Sha1);

        // Simulate EncryptedPackage stream layout: [u64 package_size] + ciphertext.
        let mut stream = Vec::new();
        stream.extend_from_slice(&(plaintext.len() as u64).to_le_bytes());
        stream.extend_from_slice(&ciphertext);

        let cursor = Cursor::new(stream);
        let mut reader = Rc4CryptoApiDecryptReader::from_encrypted_package_stream(
            cursor,
            h.clone(),
            56, // keySize (bits)
            CALG_SHA1,
        )
        .unwrap();

        let mut out = vec![0u8; plaintext.len()];
        reader.read_exact(&mut out).unwrap();
        assert_eq!(out, plaintext);
    }

    #[test]
    fn encrypted_package_keysize_zero_is_interpreted_as_40_bit() {
        let h = b"0123456789ABCDEFGHIJ".to_vec(); // 20 bytes
        let key_len = 5; // 40-bit (5-byte RC4 key)

        // Ensure plaintext crosses a 0x200 boundary.
        let mut plaintext = vec![0u8; RC4_BLOCK_SIZE + 64];
        for (i, b) in plaintext.iter_mut().enumerate() {
            *b = (i % 251) as u8;
        }
        let ciphertext = encrypt_rc4_cryptoapi(&plaintext, &h, key_len, HashAlg::Sha1);

        // Simulate EncryptedPackage stream layout: [u64 package_size] + ciphertext.
        let mut stream = Vec::new();
        stream.extend_from_slice(&(plaintext.len() as u64).to_le_bytes());
        stream.extend_from_slice(&ciphertext);

        let cursor = Cursor::new(stream);
        let mut reader = Rc4CryptoApiDecryptReader::from_encrypted_package_stream(
            cursor,
            h.clone(),
            0, // keySize (bits) => 40-bit per MS-OFFCRYPTO
            CALG_SHA1,
        )
        .unwrap();

        let mut out = vec![0u8; plaintext.len()];
        reader.read_exact(&mut out).unwrap();
        assert_eq!(out, plaintext);
    }

    #[test]
    fn encrypted_package_stream_falls_back_to_low_dword_when_size_prefix_high_dword_is_reserved() {
        let h = b"0123456789ABCDEFGHIJ".to_vec(); // 20 bytes
        let key_len = 7; // 56-bit

        // Ensure plaintext crosses a 0x200 boundary.
        let mut plaintext = vec![0u8; RC4_BLOCK_SIZE + 64];
        for (i, b) in plaintext.iter_mut().enumerate() {
            *b = (i % 251) as u8;
        }
        let ciphertext = encrypt_rc4_cryptoapi(&plaintext, &h, key_len, HashAlg::Sha1);

        // Simulate EncryptedPackage stream layout but with a non-zero high DWORD in the size prefix.
        // Some producers treat this prefix as `(u32 size, u32 reserved)` and may write a non-zero
        // reserved high DWORD.
        let size_lo = u32::try_from(plaintext.len()).expect("plaintext length fits in u32");
        let mut stream = Vec::new();
        stream.extend_from_slice(&size_lo.to_le_bytes());
        stream.extend_from_slice(&1u32.to_le_bytes()); // reserved high DWORD
        stream.extend_from_slice(&ciphertext);

        let cursor = Cursor::new(stream);
        let mut reader = Rc4CryptoApiDecryptReader::from_encrypted_package_stream(
            cursor,
            h.clone(),
            56, // keySize (bits)
            CALG_SHA1,
        )
        .expect("create RC4 decrypt reader with reserved size prefix high DWORD");

        let mut out = vec![0u8; plaintext.len()];
        reader.read_exact(&mut out).unwrap();
        assert_eq!(out, plaintext);
    }

    #[test]
    fn encrypted_package_can_auto_detect_padded_40_bit_rc4_key_blob() {
        // Some producers treat a 40-bit RC4 key as a 16-byte key blob where the high 88 bits are
        // zero. `from_encrypted_package_stream` can auto-detect this quirk for EncryptedPackage
        // payloads by probing for a ZIP `PK..` signature.
        let h = b"0123456789ABCDEFGHIJ".to_vec(); // 20 bytes
        let key_len = 5; // 40-bit

        let mut plaintext = b"PK\x03\x04".to_vec();
        plaintext.extend_from_slice(&[0u8; 1024]);

        let ciphertext =
            encrypt_rc4_cryptoapi_padded_40_bit_key(&plaintext, &h, key_len, HashAlg::Sha1);

        let mut stream = Vec::new();
        stream.extend_from_slice(&(plaintext.len() as u64).to_le_bytes());
        stream.extend_from_slice(&ciphertext);

        let cursor = Cursor::new(stream);
        let mut reader = Rc4CryptoApiDecryptReader::from_encrypted_package_stream(
            cursor,
            h.clone(),
            0, // keySize (bits) => 40-bit per MS-OFFCRYPTO
            CALG_SHA1,
        )
        .expect("create RC4 decrypt reader for padded 40-bit key blob");

        let mut out = vec![0u8; plaintext.len()];
        reader.read_exact(&mut out).unwrap();
        assert_eq!(out, plaintext);
    }

    #[test]
    fn seek_into_middle_of_block_and_read() {
        let h = b"0123456789ABCDEFGHIJ".to_vec(); // 20 bytes
        for key_len in [5usize, 7, 16] {
            let mut plaintext = vec![0u8; RC4_BLOCK_SIZE * 3];
            for (i, b) in plaintext.iter_mut().enumerate() {
                *b = (i % 251) as u8;
            }
            let ciphertext = encrypt_rc4_cryptoapi(&plaintext, &h, key_len, HashAlg::Sha1);

            let mut stream = Vec::new();
            stream.extend_from_slice(&(plaintext.len() as u64).to_le_bytes());
            stream.extend_from_slice(&ciphertext);

            let mut cursor = Cursor::new(stream);
            cursor.seek(SeekFrom::Start(8)).unwrap();

            let mut reader =
                Rc4CryptoApiDecryptReader::new(cursor, plaintext.len() as u64, h.clone(), key_len)
                    .unwrap();

            let seek_pos = (RC4_BLOCK_SIZE as u64) + 0x10;
            reader.seek(SeekFrom::Start(seek_pos)).unwrap();

            let mut buf = [0u8; 64];
            reader.read_exact(&mut buf).unwrap();

            assert_eq!(
                &buf[..],
                &plaintext[seek_pos as usize..seek_pos as usize + buf.len()],
                "seek+read mismatch (key_len={key_len})"
            );
        }
    }

    #[test]
    fn seek_beyond_package_size_behaves_like_eof() {
        let h = b"0123456789ABCDEFGHIJ".to_vec(); // 20 bytes
        for key_len in [5usize, 7, 16] {
            let plaintext = b"hello world".to_vec();
            let ciphertext = encrypt_rc4_cryptoapi(&plaintext, &h, key_len, HashAlg::Sha1);

            let mut stream = Vec::new();
            stream.extend_from_slice(&(plaintext.len() as u64).to_le_bytes());
            stream.extend_from_slice(&ciphertext);

            let mut cursor = Cursor::new(stream);
            cursor.seek(SeekFrom::Start(8)).unwrap();

            let mut reader =
                Rc4CryptoApiDecryptReader::new(cursor, plaintext.len() as u64, h.clone(), key_len)
                    .unwrap();

            // Seek beyond EOF; reader should clamp to EOF and reads should return 0.
            reader
                .seek(SeekFrom::Start(plaintext.len() as u64 + 100))
                .unwrap();

            let mut buf = [0u8; 32];
            let n = reader.read(&mut buf).unwrap();
            assert_eq!(n, 0, "expected EOF (key_len={key_len})");
        }
    }

    #[test]
    fn sha1_40_bit_vector_hello_rc4_cryptoapi() {
        // Deterministic 40-bit vector derived from `docs/offcrypto-standard-cryptoapi-rc4.md`:
        //
        // - H (spun SHA-1 password hash):
        //   1b5972284eab6481eb6565a0985b334b3e65e041
        // - block 0 digest = SHA1(H || LE32(0)):
        //   6ad7dedf2da3514b1d85eabee069d47dd058967f
        // - 40-bit RC4 key material = first 5 bytes of digest:
        //   6ad7dedf2d
        let h: Vec<u8> = vec![
            0x1b, 0x59, 0x72, 0x28, 0x4e, 0xab, 0x64, 0x81, 0xeb, 0x65, 0x65, 0xa0, 0x98, 0x5b,
            0x33, 0x4b, 0x3e, 0x65, 0xe0, 0x41,
        ];
        let key_len = 5usize;
        let plaintext = b"Hello, RC4 CryptoAPI!";
        let expected_ciphertext: Vec<u8> = vec![
            0xd1, 0xfa, 0x44, 0x49, 0x13, 0xb4, 0x83, 0x9b, 0x06, 0xeb, 0x48, 0x51, 0x75,
            0x0a, 0x07, 0x76, 0x10, 0x05, 0xf0, 0x25, 0xbf,
        ];

        // Encrypt and assert the ciphertext matches the known vector.
        let got_ciphertext = encrypt_rc4_cryptoapi(plaintext, &h, key_len, HashAlg::Sha1);
        assert_eq!(got_ciphertext, expected_ciphertext);

        // Decrypt via the streaming reader under test.
        let mut stream = Vec::new();
        stream.extend_from_slice(&(plaintext.len() as u64).to_le_bytes());
        stream.extend_from_slice(&expected_ciphertext);

        let mut cursor = Cursor::new(stream);
        cursor.seek(SeekFrom::Start(8)).unwrap();

        let mut reader =
            Rc4CryptoApiDecryptReader::new(cursor, plaintext.len() as u64, h.clone(), key_len)
                .unwrap();
        let mut out = vec![0u8; plaintext.len()];
        reader.read_exact(&mut out).unwrap();
        assert_eq!(&out[..], plaintext);
    }

    #[test]
    fn md5_sequential_reads_and_seek_work() {
        let h = b"0123456789ABCDEF".to_vec(); // 16 bytes
        let key_len = 16;

        let mut plaintext = vec![0u8; RC4_BLOCK_SIZE * 2 + 64];
        for (i, b) in plaintext.iter_mut().enumerate() {
            *b = (i % 251) as u8;
        }
        let ciphertext = encrypt_rc4_cryptoapi(&plaintext, &h, key_len, HashAlg::Md5);

        // Simulate EncryptedPackage stream layout: [u64 package_size] + ciphertext.
        let mut stream = Vec::new();
        stream.extend_from_slice(&(plaintext.len() as u64).to_le_bytes());
        stream.extend_from_slice(&ciphertext);

        let cursor = Cursor::new(stream);
        let mut reader = Rc4CryptoApiDecryptReader::from_encrypted_package_stream(
            cursor,
            h.clone(),
            (key_len as u32) * 8,
            CALG_MD5,
        )
        .unwrap();

        // Sequential read in small chunks to cross a 0x200 boundary.
        let mut out = vec![0u8; plaintext.len()];
        let out_len = out.len();
        let mut read = 0usize;
        while read < out_len {
            let end = read + 33.min(out_len - read);
            let n = reader.read(&mut out[read..end]).unwrap();
            assert!(n > 0, "unexpected EOF while reading");
            read += n;
        }
        assert_eq!(out, plaintext);

        // Seek into the middle of a block and read.
        let seek_pos = (RC4_BLOCK_SIZE as u64) + 0x10;
        reader.seek(SeekFrom::Start(seek_pos)).unwrap();
        let mut buf = [0u8; 64];
        reader.read_exact(&mut buf).unwrap();
        assert_eq!(
            &buf[..],
            &plaintext[seek_pos as usize..seek_pos as usize + buf.len()]
        );
    }

    #[test]
    fn md5_sequential_reads_and_seek_work_with_40_bit_key() {
        let h = b"0123456789ABCDEF".to_vec(); // 16 bytes
        let key_len = 5; // 40-bit (5-byte RC4 key)

        let mut plaintext = vec![0u8; RC4_BLOCK_SIZE * 2 + 64];
        for (i, b) in plaintext.iter_mut().enumerate() {
            *b = (i % 251) as u8;
        }
        let ciphertext = encrypt_rc4_cryptoapi(&plaintext, &h, key_len, HashAlg::Md5);

        // Simulate EncryptedPackage stream layout: [u64 package_size] + ciphertext.
        let mut stream = Vec::new();
        stream.extend_from_slice(&(plaintext.len() as u64).to_le_bytes());
        stream.extend_from_slice(&ciphertext);

        // Exercise `keySize=0` (=> 40-bit) + CALG_MD5 parsing/validation path.
        let cursor = Cursor::new(stream);
        let mut reader = Rc4CryptoApiDecryptReader::from_encrypted_package_stream(
            cursor,
            h.clone(),
            0, // keySize (bits) => 40-bit per MS-OFFCRYPTO
            CALG_MD5,
        )
        .unwrap();

        // Sequential read in small chunks to cross a 0x200 boundary.
        let mut out = vec![0u8; plaintext.len()];
        let out_len = out.len();
        let mut read = 0usize;
        while read < out_len {
            let end = read + 33.min(out_len - read);
            let n = reader.read(&mut out[read..end]).unwrap();
            assert!(n > 0, "unexpected EOF while reading");
            read += n;
        }
        assert_eq!(out, plaintext);

        // Seek into the middle of a block and read.
        let seek_pos = (RC4_BLOCK_SIZE as u64) + 0x10;
        reader.seek(SeekFrom::Start(seek_pos)).unwrap();
        let mut buf = [0u8; 64];
        reader.read_exact(&mut buf).unwrap();
        assert_eq!(
            &buf[..],
            &plaintext[seek_pos as usize..seek_pos as usize + buf.len()]
        );
    }

    #[test]
    fn encrypted_package_errors_on_too_short_stream() {
        let stream = Cursor::new(vec![0u8; 7]);
        let err = Rc4CryptoApiDecryptReader::from_encrypted_package_stream(
            stream,
            vec![0u8; 20],
            128,
            CALG_SHA1,
        )
        .unwrap_err();
        assert!(matches!(
            err,
            Rc4CryptoApiEncryptedPackageError::StreamTooShort { len: 7 }
        ));
    }

    #[test]
    fn encrypted_package_errors_on_bogus_key_size() {
        // Valid EncryptedPackage framing with empty ciphertext and package_size=0.
        let stream = Cursor::new(0u64.to_le_bytes().to_vec());
        let err = Rc4CryptoApiDecryptReader::from_encrypted_package_stream(
            stream,
            vec![0u8; 20],
            7, // not divisible by 8
            CALG_SHA1,
        )
        .unwrap_err();
        assert!(matches!(
            err,
            Rc4CryptoApiEncryptedPackageError::InvalidKeySizeBits { key_size_bits: 7 }
        ));
    }

    #[test]
    fn encrypted_package_errors_on_unsupported_hash_alg() {
        // Valid EncryptedPackage framing with empty ciphertext and package_size=0.
        let stream = Cursor::new(0u64.to_le_bytes().to_vec());
        let err = Rc4CryptoApiDecryptReader::from_encrypted_package_stream(
            stream,
            vec![0u8; 20],
            40,
            0xDEAD_BEEF,
        )
        .unwrap_err();
        assert!(matches!(
            err,
            Rc4CryptoApiEncryptedPackageError::UnsupportedHashAlgorithm {
                alg_id_hash: 0xDEAD_BEEF
            }
        ));
    }

    #[test]
    fn encrypted_package_errors_when_package_size_exceeds_ciphertext() {
        // Header claims 100 bytes of plaintext, but we only have 10 bytes of ciphertext.
        let mut bytes = Vec::new();
        bytes.extend_from_slice(&100u64.to_le_bytes());
        bytes.extend_from_slice(&[0u8; 10]);

        let stream = Cursor::new(bytes);
        let err = Rc4CryptoApiDecryptReader::from_encrypted_package_stream(
            stream,
            vec![0u8; 20],
            128,
            CALG_SHA1,
        )
        .unwrap_err();
        assert!(matches!(
            err,
            Rc4CryptoApiEncryptedPackageError::PackageSizeExceedsCiphertext {
                package_size: 100,
                ciphertext_len: 10
            }
        ));
    }
}
