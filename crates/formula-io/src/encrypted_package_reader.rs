use std::fmt;
use std::io::{self, Read, Seek, SeekFrom};

use aes::cipher::{generic_array::GenericArray, BlockDecrypt, KeyInit};
use aes::{Aes128, Aes192, Aes256};
use formula_xlsx::offcrypto::decrypt_aes_cbc_no_padding_in_place;
use sha1::Digest as _;
use zeroize::{Zeroize, Zeroizing};

const AES_BLOCK_SIZE: usize = 16;
const SEGMENT_SIZE: usize = 0x1000;

pub(crate) enum EncryptionMethod {
    /// MS-OFFCRYPTO "Standard" (CryptoAPI) encryption: AES-ECB `EncryptedPackage`.
    ///
    /// In the baseline Standard/CryptoAPI AES scheme, the `EncryptedPackage` stream is encrypted
    /// using AES-ECB (no IV, no chaining).
    StandardAesEcb {
        /// AES key bytes (16/24/32).
        key: Zeroizing<Vec<u8>>,
    },

    /// MS-OFFCRYPTO "Standard" (CryptoAPI) encryption: segmented AES-CBC `EncryptedPackage`.
    ///
    /// Some producers encrypt `EncryptedPackage` in 4096-byte segments using AES-CBC, with a
    /// per-segment IV derived as `SHA1(salt || LE32(segment_index))[0..16]`.
    ///
    StandardCryptoApi {
        /// AES key bytes (16/24/32).
        key: Zeroizing<Vec<u8>>,
        /// `EncryptionVerifier.salt` from the Standard `EncryptionInfo` payload.
        salt: Vec<u8>,
    },

    /// MS-OFFCRYPTO "Agile" encryption.
    ///
    /// The `EncryptedPackage` ciphertext is encrypted in 4096-byte segments using AES-CBC.
    /// For segment `i`, the IV is `Truncate(blockSize, Hash(salt || LE32(i)))`.
    Agile {
        /// AES key bytes (16/24/32).
        key: Zeroizing<Vec<u8>>,
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

impl<R> Drop for DecryptedPackageReader<R> {
    fn drop(&mut self) {
        // Best-effort: wipe cached plaintext and any key material held by `EncryptionMethod`.
        zeroize_vec_u8_full(&mut self.cached_segment_plain);
        zeroize_vec_u8_full(&mut self.scratch);
    }
}

impl fmt::Debug for EncryptionMethod {
    fn fmt(&self, f: &mut fmt::Formatter<'_>) -> fmt::Result {
        match self {
            EncryptionMethod::StandardAesEcb { key } => f
                .debug_struct("StandardAesEcb")
                .field("key_len", &key.len())
                .finish(),
            EncryptionMethod::StandardCryptoApi { key, salt } => f
                .debug_struct("StandardCryptoApi")
                .field("key_len", &key.len())
                .field("salt_len", &salt.len())
                .finish(),
            EncryptionMethod::Agile {
                key,
                salt,
                hash_alg,
                block_size,
            } => f
                .debug_struct("Agile")
                .field("key_len", &key.len())
                .field("salt_len", &salt.len())
                .field("hash_alg", hash_alg)
                .field("block_size", block_size)
                .finish(),
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

        // For both Standard and Agile encryption, ciphertext segments are laid out so that segment
        // `i` starts at ciphertext offset `i * 0x1000`.
        let seg_cipher_start = seg_plain_start;

        // Reuse the old cached plaintext buffer as scratch space to avoid copies.
        self.scratch.clear();
        std::mem::swap(&mut self.scratch, &mut self.cached_segment_plain);

        // `self.scratch` now contains the previous segment's plaintext bytes. Wipe them before
        // resizing/reading the next ciphertext segment, so sensitive data doesn't linger in the
        // Vec's spare capacity (e.g. when the next segment is shorter).
        zeroize_vec_u8_full(&mut self.scratch);

        self.scratch.resize(seg_cipher_len, 0);
        self.inner.seek(SeekFrom::Start(seg_cipher_start))?;
        self.inner.read_exact(&mut self.scratch)?;

        match &self.method {
            EncryptionMethod::StandardAesEcb { key } => {
                aes_ecb_decrypt_in_place(key.as_slice(), &mut self.scratch)?;
            }
            EncryptionMethod::StandardCryptoApi { key, salt } => {
                let seg_index_u32 = u32::try_from(segment_index).map_err(|_| {
                    io::Error::new(io::ErrorKind::InvalidInput, "segment index exceeds u32")
                })?;
                let iv = derive_standard_segment_iv(salt, seg_index_u32);
                decrypt_aes_cbc_no_padding_in_place(key.as_slice(), &iv, &mut self.scratch)
                    .map_err(map_aes_cbc_err)?;
            }
            EncryptionMethod::Agile {
                key,
                salt,
                hash_alg,
                block_size,
            } => {
                let seg_index_u32 = u32::try_from(segment_index).map_err(|_| {
                    io::Error::new(io::ErrorKind::InvalidInput, "segment index exceeds u32")
                })?;
                if *block_size != AES_BLOCK_SIZE {
                    return Err(io::Error::new(
                        io::ErrorKind::InvalidData,
                        format!(
                            "unsupported Agile keyData.blockSize {block_size} (expected {AES_BLOCK_SIZE})"
                        ),
                    ));
                }
                let iv = derive_agile_segment_iv(salt, *hash_alg, seg_index_u32);
                decrypt_aes_cbc_no_padding_in_place(key.as_slice(), &iv, &mut self.scratch)
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

fn zeroize_vec_u8_full(buf: &mut Vec<u8>) {
    buf.zeroize();
    for slot in buf.spare_capacity_mut() {
        slot.write(0);
    }
}

fn aes_ecb_decrypt_in_place(key: &[u8], buf: &mut [u8]) -> io::Result<()> {
    if buf.len() % AES_BLOCK_SIZE != 0 {
        return Err(io::Error::new(
            io::ErrorKind::InvalidData,
            format!(
                "ciphertext length {len} is not a multiple of AES block size ({AES_BLOCK_SIZE})",
                len = buf.len()
            ),
        ));
    }

    fn decrypt_with<C>(key: &[u8], buf: &mut [u8]) -> io::Result<()>
    where
        C: BlockDecrypt + KeyInit,
    {
        let cipher = C::new_from_slice(key).map_err(|_| {
            io::Error::new(
                io::ErrorKind::InvalidData,
                format!(
                    "unsupported AES key length {len} (expected 16/24/32)",
                    len = key.len()
                ),
            )
        })?;
        for block in buf.chunks_mut(AES_BLOCK_SIZE) {
            cipher.decrypt_block(GenericArray::from_mut_slice(block));
        }
        Ok(())
    }

    match key.len() {
        16 => decrypt_with::<Aes128>(key, buf),
        24 => decrypt_with::<Aes192>(key, buf),
        32 => decrypt_with::<Aes256>(key, buf),
        other => Err(io::Error::new(
            io::ErrorKind::InvalidData,
            format!("unsupported AES key length {other} (expected 16/24/32)"),
        )),
    }
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
    let mut hasher = sha1::Sha1::new();
    hasher.update(salt);
    hasher.update(&segment_index.to_le_bytes());
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
    use aes::cipher::{generic_array::GenericArray, BlockEncrypt, KeyInit};
    use aes::{Aes128, Aes192, Aes256};
    use cbc::cipher::block_padding::NoPadding;
    use cbc::cipher::{BlockEncryptMut, KeyIvInit};
    use proptest::prelude::*;
    use std::io::Cursor;

    fn patterned_bytes(len: usize) -> Vec<u8> {
        (0..len)
            .map(|i| (i.wrapping_mul(31) ^ (i >> 3)) as u8)
            .collect()
    }

    fn aes_ecb_encrypt_in_place(key: &[u8], buf: &mut [u8]) {
        assert!(buf.len() % AES_BLOCK_SIZE == 0);
        fn encrypt_with<C>(key: &[u8], buf: &mut [u8])
        where
            C: BlockEncrypt + KeyInit,
        {
            let cipher = C::new_from_slice(key).expect("valid key length");
            for block in buf.chunks_mut(AES_BLOCK_SIZE) {
                cipher.encrypt_block(GenericArray::from_mut_slice(block));
            }
        }

        match key.len() {
            16 => encrypt_with::<Aes128>(key, buf),
            24 => encrypt_with::<Aes192>(key, buf),
            32 => encrypt_with::<Aes256>(key, buf),
            other => panic!("unsupported AES key length {other}"),
        }
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

        let mut segment_index = 0u32;
        let mut offset = 0usize;
        while offset < plaintext.len() {
            let seg_plain_len = (plaintext.len() - offset).min(SEGMENT_SIZE);
            let seg_cipher_len = round_up_to_multiple(seg_plain_len, AES_BLOCK_SIZE);

            let mut seg = vec![0u8; seg_cipher_len];
            seg[..seg_plain_len].copy_from_slice(&plaintext[offset..offset + seg_plain_len]);

            let iv = derive_standard_segment_iv(&salt, segment_index);
            let ciphertext = encrypt_segment_aes_cbc_no_padding(key, &iv, &seg);
            out.extend_from_slice(&ciphertext);

            offset += seg_plain_len;
            segment_index += 1;
        }
        out
    }

    fn encrypt_standard_aes_ecb_stream(key: &[u8], plaintext: &[u8]) -> Vec<u8> {
        let orig_size = plaintext.len() as u64;
        let mut out = Vec::new();
        out.extend_from_slice(&orig_size.to_le_bytes());

        if plaintext.is_empty() {
            return out;
        }

        let mut offset = 0usize;
        while offset < plaintext.len() {
            let seg_plain_len = (plaintext.len() - offset).min(SEGMENT_SIZE);
            let seg_cipher_len = round_up_to_multiple(seg_plain_len, AES_BLOCK_SIZE);

            let mut seg = vec![0u8; seg_cipher_len];
            seg[..seg_plain_len].copy_from_slice(&plaintext[offset..offset + seg_plain_len]);

            aes_ecb_encrypt_in_place(key, &mut seg);
            out.extend_from_slice(&seg);

            offset += seg_plain_len;
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
        let plaintext = patterned_bytes(10_123);
        let key = [7u8; 16];
        let salt = [0x11u8; 16];

        let encrypted_stream = encrypt_standard_cryptoapi_stream(&key, &salt, &plaintext);
        let ciphertext = &encrypted_stream[8..];

        let inner = Cursor::new(ciphertext.to_vec());
        let mut reader = DecryptedPackageReader::new(
            inner,
            EncryptionMethod::StandardCryptoApi {
                key: Zeroizing::new(key.to_vec()),
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

        // Cross-check against full-stream CBC-segmented decrypt helper.
        let decrypted = crate::offcrypto::decrypt_encrypted_package_standard_aes_sha1(
            &encrypted_stream,
            &key,
            &salt,
        )
        .expect("full decrypt");
        assert_eq!(decrypted, plaintext);
    }

    #[derive(Debug, Clone)]
    enum ReadSeekOp {
        SeekStart(u64),
        SeekCurrent(i64),
        SeekEnd(i64),
        Read(usize),
    }

    fn read_seek_op_strategy() -> impl Strategy<Value = ReadSeekOp> {
        // Keep the ranges conservative so the test is fast and doesn't explore pathological i64
        // corner cases. The goal is to stress the segmented-cache logic across random access
        // patterns.
        let seek_abs_max = 25_000u64;
        let seek_rel_max = 25_000i64;
        let read_len_max = SEGMENT_SIZE * 2;

        prop_oneof![
            (0u64..=seek_abs_max).prop_map(ReadSeekOp::SeekStart),
            (-seek_rel_max..=seek_rel_max).prop_map(ReadSeekOp::SeekCurrent),
            (-seek_rel_max..=seek_rel_max).prop_map(ReadSeekOp::SeekEnd),
            (0usize..=read_len_max).prop_map(ReadSeekOp::Read),
        ]
    }

    proptest! {
        #![proptest_config(ProptestConfig {
            cases: 32,
            max_shrink_iters: 0,
            .. ProptestConfig::default()
        })]

        #[test]
        fn prop_standard_cryptoapi_reader_matches_plaintext(
            plaintext in proptest::collection::vec(any::<u8>(), 0..=20_000),
            ops in proptest::collection::vec(read_seek_op_strategy(), 0..=64),
        ) {
            let key = [7u8; 16];
            let salt = [0x11u8; 16];

            let encrypted_stream = encrypt_standard_cryptoapi_stream(&key, &salt, &plaintext);
            let ciphertext = encrypted_stream.get(8..).unwrap_or(&[]);

            let inner = Cursor::new(ciphertext.to_vec());
            let mut reader = DecryptedPackageReader::new(
                inner,
                EncryptionMethod::StandardCryptoApi {
                    key: Zeroizing::new(key.to_vec()),
                    salt: salt.to_vec(),
                },
                plaintext.len() as u64,
            );

            let mut expected_pos: u64 = 0;
            let plaintext_len = plaintext.len() as u64;

            for op in ops {
                match op {
                    ReadSeekOp::SeekStart(pos) => {
                        let res = reader.seek(SeekFrom::Start(pos));
                        prop_assert!(res.is_ok());
                        expected_pos = pos;
                    }
                    ReadSeekOp::SeekCurrent(off) => {
                        let new_pos = expected_pos as i128 + off as i128;
                        let res = reader.seek(SeekFrom::Current(off));
                        if new_pos < 0 {
                            prop_assert!(res.is_err());
                        } else {
                            prop_assert_eq!(res.unwrap(), new_pos as u64);
                            expected_pos = new_pos as u64;
                        }
                    }
                    ReadSeekOp::SeekEnd(off) => {
                        let new_pos = plaintext_len as i128 + off as i128;
                        let res = reader.seek(SeekFrom::End(off));
                        if new_pos < 0 {
                            prop_assert!(res.is_err());
                        } else {
                            prop_assert_eq!(res.unwrap(), new_pos as u64);
                            expected_pos = new_pos as u64;
                        }
                    }
                    ReadSeekOp::Read(len) => {
                        let mut buf = vec![0u8; len];
                        let n = reader.read(&mut buf).expect("read should not error");

                        let expected_n = if expected_pos >= plaintext_len {
                            0usize
                        } else {
                            let remaining = (plaintext_len - expected_pos) as usize;
                            remaining.min(len)
                        };
                        prop_assert_eq!(n, expected_n);

                        if n > 0 {
                            prop_assert_eq!(
                                &buf[..n],
                                &plaintext[expected_pos as usize..expected_pos as usize + n]
                            );
                            expected_pos += n as u64;
                        }
                    }
                }
            }
        }
    }

    #[test]
    #[cfg(not(target_arch = "wasm32"))]
    fn standard_cryptoapi_fixture_read_seek_decrypts_zip() {
        use std::io::Read as _;

        let fixture_path = std::path::Path::new(env!("CARGO_MANIFEST_DIR"))
            .join("tests/fixtures/offcrypto_standard_cryptoapi_password.xlsx");
        let file = std::fs::File::open(&fixture_path).expect("open fixture");
        let mut ole = cfb::CompoundFile::open(file).expect("parse OLE");

        let mut encryption_info = Vec::new();
        ole.open_stream("EncryptionInfo")
            .expect("open EncryptionInfo")
            .read_to_end(&mut encryption_info)
            .expect("read EncryptionInfo");

        let mut encrypted_package = Vec::new();
        ole.open_stream("EncryptedPackage")
            .expect("open EncryptedPackage")
            .read_to_end(&mut encrypted_package)
            .expect("read EncryptedPackage");

        let info = match formula_offcrypto::parse_encryption_info(&encryption_info)
            .expect("parse encryption info")
        {
            formula_offcrypto::EncryptionInfo::Standard {
                header, verifier, ..
            } => formula_offcrypto::StandardEncryptionInfo { header, verifier },
            other => panic!("expected Standard encryption, got {other:?}"),
        };

        let password = "password";
        let key =
            formula_offcrypto::standard_derive_key_zeroizing(&info, password).expect("derive key");
        formula_offcrypto::standard_verify_key(&info, key.as_slice()).expect("verify key");

        let plaintext_len = crate::parse_encrypted_package_original_size(&encrypted_package)
            .expect("EncryptedPackage size header");
        let ciphertext = encrypted_package[8..].to_vec();

        let method = if plaintext_len == 0 {
            EncryptionMethod::StandardAesEcb { key: key.clone() }
        } else {
            let first = ciphertext.get(..AES_BLOCK_SIZE).expect("first AES block");
            let mut ecb = first.to_vec();
            let ecb_ok = aes_ecb_decrypt_in_place(&key, &mut ecb).is_ok() && ecb.starts_with(b"PK");

            let mut cbc = first.to_vec();
            let iv = derive_standard_segment_iv(&info.verifier.salt, 0);
            let cbc_ok = decrypt_aes_cbc_no_padding_in_place(&key, &iv, &mut cbc).is_ok()
                && cbc.starts_with(b"PK");

            if ecb_ok {
                EncryptionMethod::StandardAesEcb { key: key.clone() }
            } else if cbc_ok {
                EncryptionMethod::StandardCryptoApi {
                    key: key.clone(),
                    salt: info.verifier.salt.clone(),
                }
            } else {
                panic!("unable to detect Standard EncryptedPackage cipher mode for fixture");
            }
        };

        let mut reader =
            DecryptedPackageReader::new(Cursor::new(ciphertext), method, plaintext_len);

        // Decrypted ZIP local file header starts with `PK`.
        let mut sig = [0u8; 2];
        reader.read_exact(&mut sig).expect("read signature");
        assert_eq!(&sig, b"PK");

        // Ensure ZipArchive can read the central directory using Seek.
        reader.seek(SeekFrom::Start(0)).expect("rewind");
        let mut zip = zip::ZipArchive::new(reader).expect("open decrypted ZIP");
        let mut part = zip
            .by_name("[Content_Types].xml")
            .expect("read [Content_Types].xml");
        let mut xml = String::new();
        part.read_to_string(&mut xml).expect("read xml");
        assert!(
            xml.contains("<Types"),
            "expected [Content_Types].xml to contain <Types, got: {xml:?}"
        );
    }

    #[test]
    fn standard_aes_ecb_read_seek_matches_plaintext() {
        let plaintext = patterned_bytes(10_000);
        let key = [5u8; 16];

        let encrypted_stream = encrypt_standard_aes_ecb_stream(&key, &plaintext);
        let ciphertext = &encrypted_stream[8..];

        let inner = Cursor::new(ciphertext.to_vec());
        let mut reader = DecryptedPackageReader::new(
            inner,
            EncryptionMethod::StandardAesEcb {
                key: Zeroizing::new(key.to_vec()),
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

        // Cross-check against full-stream AES-ECB decrypt helper.
        let decrypted = crate::offcrypto::decrypt_standard_encrypted_package_stream(
            &encrypted_stream,
            &key,
            &[],
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
                key: Zeroizing::new(key.to_vec()),
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

    #[test]
    fn encryption_method_debug_redacts_key_material() {
        let key_bytes = b"super_secret_key".to_vec();
        let key_debug = format!("{key_bytes:?}");

        let standard_ecb = EncryptionMethod::StandardAesEcb {
            key: Zeroizing::new(key_bytes.clone()),
        };
        let standard_ecb_dbg = format!("{standard_ecb:?}");
        assert!(
            !standard_ecb_dbg.contains(&key_debug),
            "Debug leaked Standard ECB key bytes: {standard_ecb_dbg}"
        );

        let standard_cbc = EncryptionMethod::StandardCryptoApi {
            key: Zeroizing::new(key_bytes.clone()),
            salt: vec![0u8; 16],
        };
        let standard_cbc_dbg = format!("{standard_cbc:?}");
        assert!(
            !standard_cbc_dbg.contains(&key_debug),
            "Debug leaked Standard CBC key bytes: {standard_cbc_dbg}"
        );

        let agile = EncryptionMethod::Agile {
            key: Zeroizing::new(key_bytes),
            salt: vec![0u8; 16],
            hash_alg: formula_xlsx::offcrypto::HashAlgorithm::Sha256,
            block_size: AES_BLOCK_SIZE,
        };
        let agile_dbg = format!("{agile:?}");
        assert!(
            !agile_dbg.contains(&key_debug),
            "Debug leaked Agile key bytes: {agile_dbg}"
        );
    }
}
