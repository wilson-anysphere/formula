use std::io::{Read, Seek, SeekFrom};
use std::num::NonZeroUsize;

use lru::LruCache;
use zeroize::Zeroizing;

use crate::agile::{derive_agile_package_key, parse_agile_encryption_info};
use crate::crypto::{
    aes_cbc_decrypt, aes_ecb_decrypt_in_place, derive_iv, rc4_xor_in_place, HashAlgorithm,
    StandardKeyDerivation, StandardKeyDeriver,
};
use crate::error::OfficeCryptoError;
use crate::standard::{parse_standard_encryption_info, EncryptionHeader, EncryptionVerifier};
use crate::util::{ct_eq, parse_encryption_info_header, EncryptionInfoKind};

const ENCRYPTED_PACKAGE_SIZE_PREFIX_LEN: u64 = 8;
const ENCRYPTED_PACKAGE_SEGMENT_LEN: u64 = 4096;
const AES_BLOCK_LEN: usize = 16;

// CryptoAPI algorithm identifiers (MS-OFFCRYPTO Standard / CryptoAPI encryption).
const CALG_RC4: u32 = 0x0000_6801;
const CALG_AES_128: u32 = 0x0000_660E;
const CALG_AES_192: u32 = 0x0000_660F;
const CALG_AES_256: u32 = 0x0000_6610;

const STANDARD_RC4_ENCRYPTED_PACKAGE_BLOCK_SIZE: usize = 0x200;
const STANDARD_RC4_BLOCKS_PER_SEGMENT: u32 =
    (ENCRYPTED_PACKAGE_SEGMENT_LEN as u32) / (STANDARD_RC4_ENCRYPTED_PACKAGE_BLOCK_SIZE as u32);

#[inline]
fn padded_aes_len(len: usize) -> usize {
    let rem = len % AES_BLOCK_LEN;
    if rem == 0 {
        len
    } else {
        len + (AES_BLOCK_LEN - rem)
    }
}

enum PackageDecryptor {
    /// Fallback for malformed `EncryptionInfo` streams where the `EncryptedPackage` payload is
    /// already plaintext ZIP bytes (still prefixed by an 8-byte size).
    Plaintext,
    Agile {
        package_key: Zeroizing<Vec<u8>>,
        hash_alg: HashAlgorithm,
        salt: Vec<u8>,
        block_size: usize,
    },
    StandardAes {
        /// AES key derived for `block=0`.
        key0: Zeroizing<Vec<u8>>,
    },
    StandardRc4 {
        /// Key-derivation helper for per-0x200-block RC4 keys.
        deriver: StandardKeyDeriver,
        key_bits: u32,
    },
}

pub struct EncryptedPackageReader<R: Read + Seek + Send + Sync> {
    ole: cfb::CompoundFile<R>,
    decrypted_len: u64,
    pos: u64,
    decryptor: PackageDecryptor,
    cache: LruCache<u32, Vec<u8>>,
    cache_bytes: usize,
    max_cache_bytes: usize,
}

impl<R: Read + Seek + Send + Sync> EncryptedPackageReader<R> {
    pub(crate) fn new(
        mut ole: cfb::CompoundFile<R>,
        password: &str,
        options: crate::DecryptOptions,
    ) -> Result<Self, OfficeCryptoError> {
        let mut encryption_info = Vec::new();
        crate::open_stream_case_tolerant(&mut ole, "EncryptionInfo")?
            .read_to_end(&mut encryption_info)?;
        let header = parse_encryption_info_header(&encryption_info)?;

        // Read the EncryptedPackage size prefix (unencrypted).
        let decrypted_len = {
            let mut s = crate::open_stream_case_tolerant(&mut ole, "EncryptedPackage")?;
            let mut b = [0u8; 8];
            s.read_exact(&mut b)?;
            u64::from_le_bytes(b)
        };

        let decryptor = match header.kind {
            EncryptionInfoKind::Agile => {
                // Some malformed containers may include only the 8-byte `EncryptionVersionInfo`
                // header with no Agile XML payload. If the `EncryptedPackage` bytes already look
                // like a ZIP, treat them as plaintext.
                if header.header_size == 0 {
                    if decrypted_len < 2 {
                        return Err(OfficeCryptoError::InvalidFormat(
                            "missing Agile EncryptionInfo XML payload".to_string(),
                        ));
                    }
                    let pt =
                        read_encrypted_package_exact(&mut ole, ENCRYPTED_PACKAGE_SIZE_PREFIX_LEN, 2)?;
                    if pt != b"PK" {
                        return Err(OfficeCryptoError::InvalidFormat(
                            "missing Agile EncryptionInfo XML and EncryptedPackage is not plaintext ZIP"
                                .to_string(),
                        ));
                    }
                    PackageDecryptor::Plaintext
                } else {
                    let info = parse_agile_encryption_info(&encryption_info, &header)?;

                    // `spinCount` is attacker-controlled; enforce limits up front to avoid CPU DoS.
                    if info.password_key_encryptor.spin_count > options.max_spin_count {
                        return Err(OfficeCryptoError::SpinCountTooLarge {
                            spin_count: info.password_key_encryptor.spin_count,
                            max: options.max_spin_count,
                        });
                    }

                    let package_key = derive_agile_package_key(&info, password)?;

                    // Fast sanity: decrypt first block and confirm it looks like a ZIP.
                    if decrypted_len >= 2 {
                        let ct = read_encrypted_package_exact(
                            &mut ole,
                            ENCRYPTED_PACKAGE_SIZE_PREFIX_LEN,
                            16,
                        )?;
                        let iv = derive_iv(
                            info.key_data.hash_algorithm,
                            &info.key_data.salt,
                            &0u32.to_le_bytes(),
                            info.key_data.block_size,
                        );
                        let pt = aes_cbc_decrypt(&package_key, &iv, &ct)?;
                        if pt.len() < 2 || &pt[..2] != b"PK" {
                            return Err(OfficeCryptoError::InvalidFormat(
                                "decrypted package does not look like a ZIP (missing PK signature)"
                                    .to_string(),
                            ));
                        }
                    }

                    PackageDecryptor::Agile {
                        package_key,
                        hash_alg: info.key_data.hash_algorithm,
                        salt: info.key_data.salt,
                        block_size: info.key_data.block_size,
                    }
                }
            }
            EncryptionInfoKind::Standard => {
                let info = parse_standard_encryption_info(&encryption_info, &header)?;
                let hash_alg = HashAlgorithm::from_cryptoapi_alg_id_hash(info.header.alg_id_hash)?;

                match info.header.alg_id {
                    CALG_RC4 => {
                        if info.header.key_bits % 8 != 0 {
                            return Err(OfficeCryptoError::InvalidFormat(format!(
                                "EncryptionHeader keyBits must be divisible by 8 (got {})",
                                info.header.key_bits
                            )));
                        }
                        let key_len = (info.header.key_bits / 8) as usize;
                        if !matches!(key_len, 5 | 7 | 16) {
                            return Err(OfficeCryptoError::UnsupportedEncryption(format!(
                                "unsupported RC4 key length {key_len} bytes (keyBits={})",
                                info.header.key_bits
                            )));
                        }

                        let deriver = StandardKeyDeriver::new(
                            hash_alg,
                            info.header.key_bits,
                            &info.verifier.salt,
                            password,
                            StandardKeyDerivation::Rc4,
                        );
                        let key0 = deriver.derive_key_for_block(0)?;
                        verify_standard_password_with_key(&info.header, &info.verifier, hash_alg, &key0)?;

                        if decrypted_len >= 2 {
                            let ct = read_encrypted_package_exact(
                                &mut ole,
                                ENCRYPTED_PACKAGE_SIZE_PREFIX_LEN,
                                2,
                            )?;
                            let mut pt = ct;
                            if info.header.key_bits == 40 {
                                if key0.len() != 5 {
                                    return Err(OfficeCryptoError::InvalidFormat(format!(
                                        "derived RC4 key for keySize=40 must be 5 bytes (got {})",
                                        key0.len()
                                    )));
                                }
                                let mut padded = [0u8; 16];
                                padded[..5].copy_from_slice(&key0[..5]);
                                rc4_xor_in_place(&padded, &mut pt)?;
                            } else {
                                rc4_xor_in_place(&key0, &mut pt)?;
                            }
                            if pt != b"PK" {
                                return Err(OfficeCryptoError::InvalidFormat(
                                    "decrypted package does not look like a ZIP (missing PK signature)"
                                        .to_string(),
                                ));
                            }
                        }

                        PackageDecryptor::StandardRc4 {
                            deriver,
                            key_bits: info.header.key_bits,
                        }
                    }
                    CALG_AES_128 | CALG_AES_192 | CALG_AES_256 => {
                        if info.header.key_bits % 8 != 0 {
                            return Err(OfficeCryptoError::InvalidFormat(format!(
                                "EncryptionHeader keyBits must be divisible by 8 (got {})",
                                info.header.key_bits
                            )));
                        }
                        let deriver = StandardKeyDeriver::new(
                            hash_alg,
                            info.header.key_bits,
                            &info.verifier.salt,
                            password,
                            StandardKeyDerivation::Aes,
                        );
                        let key0 = deriver.derive_key_for_block(0)?;
                        verify_standard_password_with_key(&info.header, &info.verifier, hash_alg, &key0)?;

                        // Fast sanity: decrypt first block and confirm it looks like a ZIP.
                        if decrypted_len >= 2 {
                            let ct = read_encrypted_package_exact(
                                &mut ole,
                                ENCRYPTED_PACKAGE_SIZE_PREFIX_LEN,
                                16,
                            )?;
                            let mut pt = ct;
                            aes_ecb_decrypt_in_place(&key0, &mut pt)?;
                            if pt.len() < 2 || &pt[..2] != b"PK" {
                                return Err(OfficeCryptoError::InvalidFormat(
                                    "decrypted package does not look like a ZIP (missing PK signature)"
                                        .to_string(),
                                ));
                            }
                        }

                        PackageDecryptor::StandardAes { key0 }
                    }
                    other => {
                        return Err(OfficeCryptoError::UnsupportedEncryption(format!(
                            "unsupported cipher AlgID {other:#x} for EncryptedPackage"
                        )));
                    }
                }
            }
        };

        let cap_segments = std::cmp::max(
            1usize,
            options.max_cache_bytes / ENCRYPTED_PACKAGE_SEGMENT_LEN as usize,
        );
        let cache = LruCache::new(NonZeroUsize::new(cap_segments).unwrap_or(NonZeroUsize::MIN));

        Ok(Self {
            ole,
            decrypted_len,
            pos: 0,
            decryptor,
            cache,
            cache_bytes: 0,
            max_cache_bytes: options.max_cache_bytes,
        })
    }

    fn segment_plain_len(&self, segment_index: u32) -> usize {
        let start = (segment_index as u64) * ENCRYPTED_PACKAGE_SEGMENT_LEN;
        if start >= self.decrypted_len {
            return 0;
        }
        let remaining = self.decrypted_len - start;
        remaining.min(ENCRYPTED_PACKAGE_SEGMENT_LEN) as usize
    }

    fn segment_cipher_offset(&self, segment_index: u32) -> u64 {
        ENCRYPTED_PACKAGE_SIZE_PREFIX_LEN + (segment_index as u64) * ENCRYPTED_PACKAGE_SEGMENT_LEN
    }

    fn decrypt_segment(&mut self, segment_index: u32) -> Result<Vec<u8>, OfficeCryptoError> {
        let plain_len = self.segment_plain_len(segment_index);
        if plain_len == 0 {
            return Ok(Vec::new());
        }

        let cipher_offset = self.segment_cipher_offset(segment_index);

        match &mut self.decryptor {
            PackageDecryptor::Plaintext => {
                read_encrypted_package_exact(&mut self.ole, cipher_offset, plain_len)
            }
            PackageDecryptor::Agile {
                package_key,
                hash_alg,
                salt,
                block_size,
            } => {
                let cipher_len = padded_aes_len(plain_len);
                let ciphertext = read_encrypted_package_exact(&mut self.ole, cipher_offset, cipher_len)?;
                let iv = derive_iv(*hash_alg, salt, &segment_index.to_le_bytes(), *block_size);
                let padded_plain = aes_cbc_decrypt(package_key, &iv, &ciphertext)?;
                Ok(padded_plain[..plain_len].to_vec())
            }
            PackageDecryptor::StandardAes { key0 } => {
                let cipher_len = padded_aes_len(plain_len);
                let mut ciphertext =
                    read_encrypted_package_exact(&mut self.ole, cipher_offset, cipher_len)?;
                aes_ecb_decrypt_in_place(key0.as_slice(), &mut ciphertext)?;
                Ok(ciphertext[..plain_len].to_vec())
            }
            PackageDecryptor::StandardRc4 { deriver, key_bits } => {
                let mut out = read_encrypted_package_exact(&mut self.ole, cipher_offset, plain_len)?;

                let block_base = segment_index
                    .checked_mul(STANDARD_RC4_BLOCKS_PER_SEGMENT)
                    .ok_or_else(|| {
                        OfficeCryptoError::InvalidFormat("RC4 block index overflow".to_string())
                    })?;
                let mut block = 0u32;
                for chunk in out.chunks_mut(STANDARD_RC4_ENCRYPTED_PACKAGE_BLOCK_SIZE) {
                    let block_index = block_base.checked_add(block).ok_or_else(|| {
                        OfficeCryptoError::InvalidFormat("RC4 block index overflow".to_string())
                    })?;
                    let key = deriver.derive_key_for_block(block_index)?;
                    if *key_bits == 40 {
                        if key.len() != 5 {
                            return Err(OfficeCryptoError::InvalidFormat(format!(
                                "derived RC4 key for keySize=40 must be 5 bytes (got {})",
                                key.len()
                            )));
                        }
                        let mut padded = [0u8; 16];
                        padded[..5].copy_from_slice(&key[..5]);
                        rc4_xor_in_place(&padded, chunk)?;
                    } else {
                        rc4_xor_in_place(&key, chunk)?;
                    }
                    block = block.checked_add(1).ok_or_else(|| {
                        OfficeCryptoError::InvalidFormat("RC4 block index overflow".to_string())
                    })?;
                }

                Ok(out)
            }
        }
    }

    fn cache_insert(&mut self, segment_index: u32, bytes: Vec<u8>) {
        if bytes.is_empty() {
            return;
        }
        let len = bytes.len();
        if let Some(evicted) = self.cache.put(segment_index, bytes) {
            self.cache_bytes = self.cache_bytes.saturating_sub(evicted.len());
        }
        self.cache_bytes = self.cache_bytes.saturating_add(len);

        while self.cache_bytes > self.max_cache_bytes && !self.cache.is_empty() {
            if let Some((_k, v)) = self.cache.pop_lru() {
                self.cache_bytes = self.cache_bytes.saturating_sub(v.len());
            }
        }
    }
}

impl<R: Read + Seek + Send + Sync> Read for EncryptedPackageReader<R> {
    fn read(&mut self, buf: &mut [u8]) -> std::io::Result<usize> {
        if buf.is_empty() {
            return Ok(0);
        }

        if self.pos >= self.decrypted_len {
            return Ok(0);
        }

        let mut total = 0usize;
        while total < buf.len() && self.pos < self.decrypted_len {
            let seg_index = u32::try_from(self.pos / ENCRYPTED_PACKAGE_SEGMENT_LEN).map_err(|_| {
                std::io::Error::new(std::io::ErrorKind::Other, "segment index overflow")
            })?;
            let seg_offset = (self.pos % ENCRYPTED_PACKAGE_SEGMENT_LEN) as usize;

            let segment = if let Some(seg) = self.cache.get(&seg_index) {
                seg.as_slice()
            } else {
                let seg = self
                    .decrypt_segment(seg_index)
                    .map_err(|e| std::io::Error::new(std::io::ErrorKind::Other, e))?;
                let slice = seg.as_slice();
                // Copy before caching so we can move `seg` into the cache.
                let to_copy = slice
                    .len()
                    .saturating_sub(seg_offset)
                    .min(buf.len() - total);
                buf[total..total + to_copy]
                    .copy_from_slice(&slice[seg_offset..seg_offset + to_copy]);
                self.pos = self.pos.saturating_add(to_copy as u64);
                total += to_copy;
                self.cache_insert(seg_index, seg);
                continue;
            };

            let to_copy = segment
                .len()
                .saturating_sub(seg_offset)
                .min(buf.len() - total);
            buf[total..total + to_copy]
                .copy_from_slice(&segment[seg_offset..seg_offset + to_copy]);
            self.pos = self.pos.saturating_add(to_copy as u64);
            total += to_copy;
        }

        Ok(total)
    }
}

impl<R: Read + Seek + Send + Sync> Seek for EncryptedPackageReader<R> {
    fn seek(&mut self, pos: SeekFrom) -> std::io::Result<u64> {
        let new_pos: i128 = match pos {
            SeekFrom::Start(n) => n as i128,
            SeekFrom::End(off) => self.decrypted_len as i128 + off as i128,
            SeekFrom::Current(off) => self.pos as i128 + off as i128,
        };

        if new_pos < 0 {
            return Err(std::io::Error::new(
                std::io::ErrorKind::InvalidInput,
                "invalid seek to a negative position",
            ));
        }

        self.pos = new_pos as u64;
        Ok(self.pos)
    }
}

fn verify_standard_password_with_key(
    header: &EncryptionHeader,
    verifier: &EncryptionVerifier,
    hash_alg: HashAlgorithm,
    key0: &[u8],
) -> Result<(), OfficeCryptoError> {
    let expected_hash_len = verifier.verifier_hash_size as usize;

    match header.alg_id {
        CALG_RC4 => {
            if verifier.encrypted_verifier.len() != 16 {
                return Err(OfficeCryptoError::InvalidFormat(format!(
                    "EncryptionVerifier.encryptedVerifier must be 16 bytes for RC4 (got {})",
                    verifier.encrypted_verifier.len()
                )));
            }

            // RC4 is a stream cipher. CryptoAPI encrypts/decrypts the verifier and verifier hash
            // using the **same** RC4 stream (continuing the keystream), so we must apply RC4 to the
            // concatenated bytes rather than resetting the cipher per field.
            let mut buf = Vec::with_capacity(
                verifier.encrypted_verifier.len() + verifier.encrypted_verifier_hash.len(),
            );
            buf.extend_from_slice(&verifier.encrypted_verifier);
            buf.extend_from_slice(&verifier.encrypted_verifier_hash);

            if header.key_bits == 40 {
                if key0.len() != 5 {
                    return Err(OfficeCryptoError::InvalidFormat(format!(
                        "derived RC4 key for keySize=40 must be 5 bytes (got {})",
                        key0.len()
                    )));
                }
                let mut padded = [0u8; 16];
                padded[..5].copy_from_slice(&key0[..5]);
                rc4_xor_in_place(&padded, &mut buf)?;
            } else {
                rc4_xor_in_place(key0, &mut buf)?;
            }

            let verifier_plain = buf.get(..16).ok_or_else(|| {
                OfficeCryptoError::InvalidFormat("RC4 verifier out of range".to_string())
            })?;
            let verifier_hash_plain_full = buf.get(16..).ok_or_else(|| {
                OfficeCryptoError::InvalidFormat("RC4 verifier hash out of range".to_string())
            })?;
            let verifier_hash_plain =
                verifier_hash_plain_full.get(..expected_hash_len).ok_or_else(|| {
                    OfficeCryptoError::InvalidFormat(format!(
                        "decrypted verifier hash shorter than verifierHashSize (got {}, need {})",
                        verifier_hash_plain_full.len(),
                        expected_hash_len
                    ))
                })?;

            let verifier_hash = hash_alg.digest(verifier_plain);
            let verifier_hash = verifier_hash.get(..expected_hash_len).ok_or_else(|| {
                OfficeCryptoError::InvalidFormat(format!(
                    "hash output shorter than verifierHashSize (got {}, need {})",
                    verifier_hash.len(),
                    expected_hash_len
                ))
            })?;

            if ct_eq(verifier_hash_plain, verifier_hash) {
                Ok(())
            } else {
                Err(OfficeCryptoError::InvalidPassword)
            }
        }
        CALG_AES_128 | CALG_AES_192 | CALG_AES_256 => {
            let mut verifier_plain = verifier.encrypted_verifier.clone();
            aes_ecb_decrypt_in_place(key0, &mut verifier_plain)?;
            let mut verifier_hash_plain_full = verifier.encrypted_verifier_hash.clone();
            aes_ecb_decrypt_in_place(key0, &mut verifier_hash_plain_full)?;

            let verifier_hash_plain =
                verifier_hash_plain_full.get(..expected_hash_len).ok_or_else(|| {
                    OfficeCryptoError::InvalidFormat(format!(
                        "decrypted verifier hash shorter than verifierHashSize (got {}, need {})",
                        verifier_hash_plain_full.len(),
                        expected_hash_len
                    ))
                })?;

            let verifier_hash = hash_alg.digest(&verifier_plain);
            let verifier_hash = verifier_hash.get(..expected_hash_len).ok_or_else(|| {
                OfficeCryptoError::InvalidFormat(format!(
                    "hash output shorter than verifierHashSize (got {}, need {})",
                    verifier_hash.len(),
                    expected_hash_len
                ))
            })?;

            if ct_eq(verifier_hash_plain, verifier_hash) {
                Ok(())
            } else {
                Err(OfficeCryptoError::InvalidPassword)
            }
        }
        other => Err(OfficeCryptoError::UnsupportedEncryption(format!(
            "unsupported cipher AlgID {other:#x}"
        ))),
    }
}

fn read_encrypted_package_exact<R: Read + Seek>(
    ole: &mut cfb::CompoundFile<R>,
    offset: u64,
    len: usize,
) -> Result<Vec<u8>, OfficeCryptoError> {
    if len == 0 {
        return Ok(Vec::new());
    }
    let mut s = crate::open_stream_case_tolerant(ole, "EncryptedPackage")?;
    s.seek(SeekFrom::Start(offset))?;
    let mut out = vec![0u8; len];
    s.read_exact(&mut out)?;
    Ok(out)
}
