use std::io::{self, Read, Seek, SeekFrom};

use formula_xlsx::offcrypto::decrypt_aes_cbc_no_padding_in_place;
use sha1::{Digest as _, Sha1};

const AES_BLOCK_SIZE: usize = 16;
const SEGMENT_SIZE: usize = 0x1000;

#[derive(Debug, Clone)]
pub(crate) enum EncryptionMethod {
    /// MS-OFFCRYPTO "Standard" (CryptoAPI) encryption.
    ///
    /// The `EncryptedPackage` ciphertext is encrypted in 4096-byte segments using AES-CBC.
    /// For segment `i`, the IV is `SHA1(salt || LE32(i))[0..16]`.
    StandardCryptoApi {
        /// AES key bytes (16/24/32).
        key: Vec<u8>,
        /// Verifier salt used to derive per-segment IVs.
        salt: Vec<u8>,
    },

    /// MS-OFFCRYPTO "Agile" encryption.
    ///
    /// The `EncryptedPackage` ciphertext is encrypted in 4096-byte segments using AES-CBC.
    /// For segment `i`, the IV is `Truncate(blockSize, Hash(salt || LE32(i)))`.
    Agile {
        /// AES key bytes (16/24/32).
        key: Vec<u8>,
        /// `keyData/@saltValue`.
        salt: Vec<u8>,
        /// `keyData/@hashAlgorithm`.
        hash_alg: formula_xlsx::offcrypto::HashAlgorithm,
        /// `keyData/@blockSize` (expected 16 for AES).
        block_size: usize,
    },
}

/// A `Read + Seek` view over an Office `EncryptedPackage` payload (ciphertext) that decrypts
/// on-demand without allocating the full decrypted ZIP.
///
/// The underlying `inner` stream must be positioned/seekable such that offset 0 corresponds to the
/// start of the ciphertext (i.e. the 8-byte `EncryptedPackage` length header has already been
/// consumed).
pub(crate) struct DecryptedPackageReader<R> {
    inner: R,
    method: EncryptionMethod,
    plaintext_len: u64,
    pos: u64,

    // Ciphertext scratch buffer (reused between segments).
    scratch: Vec<u8>,

    // Cached decrypted segment.
    cached_segment_index: Option<u64>,
    cached_segment_plain: Vec<u8>,
    cached_segment_plain_len: usize,
}

impl<R> DecryptedPackageReader<R> {
    pub(crate) fn new(inner: R, method: EncryptionMethod, plaintext_len: u64) -> Self {
        Self {
            inner,
            method,
            plaintext_len,
            pos: 0,
            scratch: Vec::new(),
            cached_segment_index: None,
            cached_segment_plain: Vec::new(),
            cached_segment_plain_len: 0,
        }
    }
}

impl<R: Read + Seek> Read for DecryptedPackageReader<R> {
    fn read(&mut self, buf: &mut [u8]) -> io::Result<usize> {
        if buf.is_empty() {
            return Ok(0);
        }
        if self.pos >= self.plaintext_len {
            return Ok(0);
        }

        let remaining = (self.plaintext_len - self.pos) as usize;
        let to_read = remaining.min(buf.len());
        self.read_segmented(&mut buf[..to_read])
    }
}

impl<R: Read + Seek> Seek for DecryptedPackageReader<R> {
    fn seek(&mut self, pos: SeekFrom) -> io::Result<u64> {
        let new_pos: i128 = match pos {
            SeekFrom::Start(n) => n as i128,
            SeekFrom::End(off) => self.plaintext_len as i128 + off as i128,
            SeekFrom::Current(off) => self.pos as i128 + off as i128,
        };
        if new_pos < 0 {
            return Err(io::Error::new(
                io::ErrorKind::InvalidInput,
                "invalid seek to a negative position",
            ));
        }
        self.pos = new_pos as u64;
        Ok(self.pos)
    }
}

impl<R: Read + Seek> DecryptedPackageReader<R> {
    fn read_segmented(&mut self, out: &mut [u8]) -> io::Result<usize> {
        let mut remaining = out.len();
        let mut written = 0usize;

        while remaining > 0 {
            let segment_index = self.pos / SEGMENT_SIZE as u64;
            let segment_offset = (self.pos % SEGMENT_SIZE as u64) as usize;

            self.ensure_segment_cached(segment_index)?;

            let seg_plain_len = self.cached_segment_plain_len;
            let available = seg_plain_len.saturating_sub(segment_offset);
            if available == 0 {
                return Err(io::Error::new(
                    io::ErrorKind::InvalidData,
                    "segment cache has no available bytes at current offset",
                ));
            }

            let take = remaining.min(available);
            out[written..written + take]
                .copy_from_slice(&self.cached_segment_plain[segment_offset..segment_offset + take]);

            self.pos += take as u64;
            written += take;
            remaining -= take;
        }

        Ok(written)
    }

    fn ensure_segment_cached(&mut self, segment_index: u64) -> io::Result<()> {
        if self.cached_segment_index == Some(segment_index) {
            return Ok(());
        }

        let seg_plain_start = segment_index
            .checked_mul(SEGMENT_SIZE as u64)
            .ok_or_else(|| io::Error::new(io::ErrorKind::InvalidInput, "segment index overflow"))?;
        if seg_plain_start >= self.plaintext_len {
            return Err(io::Error::new(
                io::ErrorKind::UnexpectedEof,
                "segment index beyond plaintext length",
            ));
        }

        let seg_plain_len =
            (self.plaintext_len - seg_plain_start).min(SEGMENT_SIZE as u64) as usize;
        let seg_cipher_len = round_up_to_multiple(seg_plain_len, AES_BLOCK_SIZE);

        // For both Standard and Agile encryption, the ciphertext segments are laid out so that
        // segment `i` starts at ciphertext offset `i * 0x1000`.
        let seg_cipher_start = seg_plain_start;

        // Reuse the old cached plaintext buffer as scratch space to avoid copies.
        self.scratch.clear();
        std::mem::swap(&mut self.scratch, &mut self.cached_segment_plain);

        self.scratch.resize(seg_cipher_len, 0);
        self.inner.seek(SeekFrom::Start(seg_cipher_start))?;
        self.inner.read_exact(&mut self.scratch)?;

        let seg_index_u32 = u32::try_from(segment_index).map_err(|_| {
            io::Error::new(io::ErrorKind::InvalidInput, "segment index exceeds u32")
        })?;

        match &self.method {
            EncryptionMethod::StandardCryptoApi { key, salt } => {
                let iv = derive_standard_segment_iv(salt, seg_index_u32);
                decrypt_aes_cbc_no_padding_in_place(key, &iv, &mut self.scratch)
                    .map_err(map_aes_cbc_err)?;
            }
            EncryptionMethod::Agile {
                key,
                salt,
                hash_alg,
                block_size,
            } => {
                if *block_size != AES_BLOCK_SIZE {
                    return Err(io::Error::new(
                        io::ErrorKind::InvalidData,
                        format!(
                            "unsupported Agile keyData.blockSize {block_size} (expected {AES_BLOCK_SIZE})"
                        ),
                    ));
                }
                let iv = derive_agile_segment_iv(salt, *hash_alg, seg_index_u32);
                decrypt_aes_cbc_no_padding_in_place(key, &iv, &mut self.scratch)
                    .map_err(map_aes_cbc_err)?;
            }
        }

        self.cached_segment_plain_len = seg_plain_len;
        self.cached_segment_index = Some(segment_index);

        // Move decrypted bytes into the cache (no copy).
        std::mem::swap(&mut self.scratch, &mut self.cached_segment_plain);
        Ok(())
    }
}

fn map_aes_cbc_err(err: formula_xlsx::offcrypto::AesCbcDecryptError) -> io::Error {
    io::Error::new(io::ErrorKind::InvalidData, err)
}

fn round_up_to_multiple(value: usize, multiple: usize) -> usize {
    if multiple == 0 {
        return value;
    }
    let rem = value % multiple;
    if rem == 0 {
        value
    } else {
        value + (multiple - rem)
    }
}

fn derive_standard_segment_iv(salt: &[u8], segment_index: u32) -> [u8; AES_BLOCK_SIZE] {
    let mut hasher = Sha1::new();
    hasher.update(salt);
    hasher.update(segment_index.to_le_bytes());
    let digest = hasher.finalize();

    let mut iv = [0u8; AES_BLOCK_SIZE];
    iv.copy_from_slice(&digest[..AES_BLOCK_SIZE]);
    iv
}

fn derive_agile_segment_iv(
    salt: &[u8],
    hash_alg: formula_xlsx::offcrypto::HashAlgorithm,
    segment_index: u32,
) -> [u8; AES_BLOCK_SIZE] {
    let mut buf = Vec::with_capacity(salt.len() + 4);
    buf.extend_from_slice(salt);
    buf.extend_from_slice(&segment_index.to_le_bytes());

    let digest = hash_bytes(hash_alg, &buf);

    let mut iv = [0u8; AES_BLOCK_SIZE];
    iv.copy_from_slice(&digest[..AES_BLOCK_SIZE]);
    iv
}

fn hash_bytes(alg: formula_xlsx::offcrypto::HashAlgorithm, data: &[u8]) -> Vec<u8> {
    match alg {
        formula_xlsx::offcrypto::HashAlgorithm::Sha1 => sha1::Sha1::digest(data).to_vec(),
        formula_xlsx::offcrypto::HashAlgorithm::Sha256 => sha2::Sha256::digest(data).to_vec(),
        formula_xlsx::offcrypto::HashAlgorithm::Sha384 => sha2::Sha384::digest(data).to_vec(),
        formula_xlsx::offcrypto::HashAlgorithm::Sha512 => sha2::Sha512::digest(data).to_vec(),
    }
}

#[cfg(test)]
mod tests {
    use super::*;
    use aes::{Aes128, Aes192, Aes256};
    use cbc::cipher::block_padding::NoPadding;
    use cbc::cipher::{BlockEncryptMut, KeyIvInit};
    use std::io::Cursor;

    fn patterned_bytes(len: usize) -> Vec<u8> {
        (0..len)
            .map(|i| (i.wrapping_mul(31) ^ (i >> 3)) as u8)
            .collect()
    }

    fn pkcs7_pad(mut plaintext: Vec<u8>) -> Vec<u8> {
        if plaintext.is_empty() {
            return plaintext;
        }
        let mut pad_len = AES_BLOCK_SIZE - (plaintext.len() % AES_BLOCK_SIZE);
        if pad_len == 0 {
            pad_len = AES_BLOCK_SIZE;
        }
        plaintext.extend(std::iter::repeat(pad_len as u8).take(pad_len));
        plaintext
    }

    fn encrypt_segment_aes_cbc_no_padding(
        key: &[u8],
        iv: &[u8; AES_BLOCK_SIZE],
        plaintext: &[u8],
    ) -> Vec<u8> {
        assert!(plaintext.len() % AES_BLOCK_SIZE == 0);

        let mut buf = plaintext.to_vec();
        match key.len() {
            16 => {
                cbc::Encryptor::<Aes128>::new_from_slices(key, iv)
                    .unwrap()
                    .encrypt_padded_mut::<NoPadding>(&mut buf, plaintext.len())
                    .unwrap();
            }
            24 => {
                cbc::Encryptor::<Aes192>::new_from_slices(key, iv)
                    .unwrap()
                    .encrypt_padded_mut::<NoPadding>(&mut buf, plaintext.len())
                    .unwrap();
            }
            32 => {
                cbc::Encryptor::<Aes256>::new_from_slices(key, iv)
                    .unwrap()
                    .encrypt_padded_mut::<NoPadding>(&mut buf, plaintext.len())
                    .unwrap();
            }
            _ => panic!("unsupported key length"),
        }
        buf
    }

    fn encrypt_standard_cryptoapi_stream(key: &[u8], salt: &[u8], plaintext: &[u8]) -> Vec<u8> {
        let orig_size = plaintext.len() as u64;
        let mut out = Vec::new();
        out.extend_from_slice(&orig_size.to_le_bytes());

        if plaintext.is_empty() {
            return out;
        }

        let padded = pkcs7_pad(plaintext.to_vec());
        for (i, chunk) in padded.chunks(SEGMENT_SIZE).enumerate() {
            let iv = derive_standard_segment_iv(salt, i as u32);
            let ciphertext = encrypt_segment_aes_cbc_no_padding(key, &iv, chunk);
            out.extend_from_slice(&ciphertext);
        }
        out
    }

    fn encrypt_agile_segments(
        plaintext: &[u8],
        key: &[u8],
        salt: &[u8],
        hash_alg: formula_xlsx::offcrypto::HashAlgorithm,
    ) -> Vec<u8> {
        let mut out = Vec::new();
        let mut segment_index: u32 = 0;
        let mut offset = 0usize;
        while offset < plaintext.len() {
            let seg_plain_len = (plaintext.len() - offset).min(SEGMENT_SIZE);
            let seg_cipher_len = round_up_to_multiple(seg_plain_len, AES_BLOCK_SIZE);

            let mut seg = vec![0u8; seg_cipher_len];
            seg[..seg_plain_len].copy_from_slice(&plaintext[offset..offset + seg_plain_len]);

            let iv = derive_agile_segment_iv(salt, hash_alg, segment_index);
            let ciphertext = encrypt_segment_aes_cbc_no_padding(key, &iv, &seg);
            out.extend_from_slice(&ciphertext);

            offset += seg_plain_len;
            segment_index += 1;
        }
        out
    }

    #[test]
    fn standard_cryptoapi_read_seek_matches_plaintext() {
        let plaintext = patterned_bytes(10_000);
        let key = [7u8; 16];
        let salt = [0x11u8; 16];

        let encrypted_stream = encrypt_standard_cryptoapi_stream(&key, &salt, &plaintext);
        let ciphertext = &encrypted_stream[8..];

        let inner = Cursor::new(ciphertext.to_vec());
        let mut reader = DecryptedPackageReader::new(
            inner,
            EncryptionMethod::StandardCryptoApi {
                key: key.to_vec(),
                salt: salt.to_vec(),
            },
            plaintext.len() as u64,
        );

        // Read a middle range.
        reader.seek(SeekFrom::Start(1234)).unwrap();
        let mut buf = vec![0u8; 777];
        reader.read_exact(&mut buf).unwrap();
        assert_eq!(&buf, &plaintext[1234..1234 + buf.len()]);

        // Cross-segment boundary.
        reader
            .seek(SeekFrom::Start(SEGMENT_SIZE as u64 - 10))
            .unwrap();
        let mut buf = vec![0u8; 40];
        reader.read_exact(&mut buf).unwrap();
        assert_eq!(&buf, &plaintext[SEGMENT_SIZE - 10..SEGMENT_SIZE - 10 + 40]);

        // Read at end (short).
        reader.seek(SeekFrom::End(-10)).unwrap();
        let mut buf = vec![0u8; 32];
        let n = reader.read(&mut buf).unwrap();
        assert_eq!(n, 10);
        assert_eq!(&buf[..10], &plaintext[plaintext.len() - 10..]);

        // Seek past end => EOF.
        reader
            .seek(SeekFrom::Start(plaintext.len() as u64 + 5))
            .unwrap();
        let mut buf = vec![0u8; 1];
        assert_eq!(reader.read(&mut buf).unwrap(), 0);

        // Cross-check against full-stream decryption helper.
        let decrypted = crate::offcrypto::decrypt_standard_encrypted_package_stream(
            &encrypted_stream,
            &key,
            &salt,
        )
        .expect("full decrypt");
        assert_eq!(decrypted, plaintext);
    }

    #[test]
    fn agile_read_seek_matches_plaintext() {
        let plaintext = patterned_bytes(10_000);
        let key = [9u8; 32];
        let salt = [3u8; 16];
        let hash_alg = formula_xlsx::offcrypto::HashAlgorithm::Sha256;

        let ciphertext = encrypt_agile_segments(&plaintext, &key, &salt, hash_alg);

        let inner = Cursor::new(ciphertext);
        let mut reader = DecryptedPackageReader::new(
            inner,
            EncryptionMethod::Agile {
                key: key.to_vec(),
                salt: salt.to_vec(),
                hash_alg,
                block_size: AES_BLOCK_SIZE,
            },
            plaintext.len() as u64,
        );

        // Cross-segment boundary.
        reader
            .seek(SeekFrom::Start(SEGMENT_SIZE as u64 - 10))
            .unwrap();
        let mut buf = vec![0u8; 40];
        reader.read_exact(&mut buf).unwrap();
        assert_eq!(&buf, &plaintext[SEGMENT_SIZE - 10..SEGMENT_SIZE - 10 + 40]);

        // Middle range read.
        reader.seek(SeekFrom::Start(1234)).unwrap();
        let mut buf = vec![0u8; 777];
        reader.read_exact(&mut buf).unwrap();
        assert_eq!(&buf, &plaintext[1234..1234 + buf.len()]);

        // Read at end (short).
        reader.seek(SeekFrom::End(-10)).unwrap();
        let mut buf = vec![0u8; 32];
        let n = reader.read(&mut buf).unwrap();
        assert_eq!(n, 10);
        assert_eq!(&buf[..10], &plaintext[plaintext.len() - 10..]);
    }
}
