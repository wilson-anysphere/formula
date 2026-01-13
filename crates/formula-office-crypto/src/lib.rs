//! Decryption of the standard Office OOXML encryption wrapper (`EncryptionInfo` + `EncryptedPackage`)
//! stored inside an OLE/CFB container.
//!
//! This crate supports:
//! - MS-OFFCRYPTO "Agile Encryption" (XML descriptor, Office 2010+)
//! - MS-OFFCRYPTO / ECMA-376 "Standard Encryption" (binary descriptor, Office 2007-era)
//!
//! The decrypted output is the raw OOXML ZIP/OPC bytes (should start with `PK`).

mod agile;
mod crypto;
mod error;
mod standard;
mod util;

use std::io::{Cursor, Read};

pub use crate::crypto::HashAlgorithm;
pub use crate::error::OfficeCryptoError;

const OLE_MAGIC: [u8; 8] = [0xD0, 0xCF, 0x11, 0xE0, 0xA1, 0xB1, 0x1A, 0xE1];

#[derive(Debug, Copy, Clone, PartialEq, Eq)]
pub enum EncryptionScheme {
    Agile,
    Standard,
}

#[derive(Debug, Clone)]
pub struct EncryptOptions {
    pub scheme: EncryptionScheme,
    pub key_bits: usize,
    pub hash_algorithm: HashAlgorithm,
    pub spin_count: u32,
}

impl Default for EncryptOptions {
    fn default() -> Self {
        Self {
            scheme: EncryptionScheme::Agile,
            key_bits: 256,
            hash_algorithm: HashAlgorithm::Sha512,
            spin_count: 100_000,
        }
    }
}

/// Returns true if the provided bytes look like an OLE/CFB container holding an Office-encrypted
/// OOXML package (streams `EncryptionInfo` and `EncryptedPackage`).
pub fn is_encrypted_ooxml_ole(bytes: &[u8]) -> bool {
    if bytes.len() < OLE_MAGIC.len() || bytes[..OLE_MAGIC.len()] != OLE_MAGIC {
        return false;
    }

    let cursor = Cursor::new(bytes);
    let Ok(mut ole) = cfb::CompoundFile::open(cursor) else {
        return false;
    };

    stream_exists(&mut ole, "EncryptionInfo") && stream_exists(&mut ole, "EncryptedPackage")
}

/// Decrypt an Office-encrypted OOXML OLE/CFB wrapper and return the decrypted raw ZIP bytes.
pub fn decrypt_encrypted_package_ole(
    bytes: &[u8],
    password: &str,
) -> Result<Vec<u8>, OfficeCryptoError> {
    let cursor = Cursor::new(bytes);
    let mut ole = cfb::CompoundFile::open(cursor)?;

    let mut encryption_info = Vec::new();
    ole.open_stream("EncryptionInfo")?
        .read_to_end(&mut encryption_info)?;

    let mut encrypted_package = Vec::new();
    ole.open_stream("EncryptedPackage")?
        .read_to_end(&mut encrypted_package)?;

    decrypt_encrypted_package(&encryption_info, &encrypted_package, password)
}

/// Encrypt a raw OOXML ZIP package into an Office `EncryptedPackage` OLE/CFB wrapper.
///
/// The returned bytes are an OLE/CFB container containing:
/// - `EncryptionInfo` stream (Agile XML descriptor, by default)
/// - `EncryptedPackage` stream (8-byte decrypted size prefix + encrypted payload)
pub fn encrypt_package_to_ole(
    zip_bytes: &[u8],
    password: &str,
    opts: EncryptOptions,
) -> Result<Vec<u8>, OfficeCryptoError> {
    use std::io::Write as _;

    let (encryption_info, encrypted_package) = match opts.scheme {
        EncryptionScheme::Agile => {
            agile::encrypt_agile_encrypted_package(zip_bytes, password, &opts)?
        }
        EncryptionScheme::Standard => {
            return Err(OfficeCryptoError::UnsupportedEncryption(
                "Standard encryption writer not implemented".to_string(),
            ))
        }
    };

    let cursor = Cursor::new(Vec::new());
    let mut ole = cfb::CompoundFile::create(cursor)?;

    ole.create_stream("EncryptionInfo")?
        .write_all(&encryption_info)?;
    ole.create_stream("EncryptedPackage")?
        .write_all(&encrypted_package)?;

    Ok(ole.into_inner().into_inner())
}

fn decrypt_encrypted_package(
    encryption_info: &[u8],
    encrypted_package: &[u8],
    password: &str,
) -> Result<Vec<u8>, OfficeCryptoError> {
    let header = util::parse_encryption_info_header(encryption_info)?;
    match header.kind {
        util::EncryptionInfoKind::Agile => {
            let info = agile::parse_agile_encryption_info(encryption_info, &header)?;
            let out = agile::decrypt_agile_encrypted_package(&info, encrypted_package, password)?;
            validate_decrypted_package(&out)?;
            Ok(out)
        }
        util::EncryptionInfoKind::Standard => {
            let info = standard::parse_standard_encryption_info(encryption_info, &header)?;
            let out =
                standard::decrypt_standard_encrypted_package(&info, encrypted_package, password)?;
            validate_decrypted_package(&out)?;
            Ok(out)
        }
    }
}

fn stream_exists<R: Read + std::io::Seek>(ole: &mut cfb::CompoundFile<R>, name: &str) -> bool {
    ole.open_stream(name).is_ok()
}

fn validate_decrypted_package(bytes: &[u8]) -> Result<(), OfficeCryptoError> {
    if bytes.len() < 2 || &bytes[..2] != b"PK" {
        return Err(OfficeCryptoError::InvalidFormat(
            "decrypted package does not look like a ZIP (missing PK signature)".to_string(),
        ));
    }
    Ok(())
}

#[cfg(test)]
mod tests {
    use super::*;
    use crate::crypto::{HashAlgorithm, StandardKeyDeriver};
    use crate::test_alloc::MAX_ALLOC;
    use std::sync::atomic::Ordering;

    #[test]
    fn detects_encrypted_ooxml_ole_container() {
        let cursor = Cursor::new(Vec::new());
        let mut ole = cfb::CompoundFile::create(cursor).expect("create cfb");
        ole.create_stream("EncryptionInfo")
            .expect("create EncryptionInfo stream");
        ole.create_stream("EncryptedPackage")
            .expect("create EncryptedPackage stream");
        let bytes = ole.into_inner().into_inner();
        assert!(is_encrypted_ooxml_ole(&bytes));
    }

    #[test]
    fn parses_standard_encryption_info_minimal() {
        let info_bytes = standard::tests::standard_encryption_info_fixture();
        let header = util::parse_encryption_info_header(&info_bytes).expect("parse header");
        assert_eq!(header.kind, util::EncryptionInfoKind::Standard);
        let parsed =
            standard::parse_standard_encryption_info(&info_bytes, &header).expect("parse standard");
        assert_eq!(parsed.version_major, 4);
        assert_eq!(parsed.version_minor, 2);
        assert_eq!(parsed.header.key_bits, 128);
        assert_eq!(parsed.verifier.salt.len(), 16);
    }

    #[test]
    fn parses_standard_encryption_info_accepts_version_2_2() {
        // Standard encryption is identified by `versionMinor == 2` and `versionMajor ∈ {2,3,4}` in
        // the wild; ensure we don't over-gate on the common `3.2`/`4.2` pairs.
        let mut info_bytes = standard::tests::standard_encryption_info_fixture();
        info_bytes[..2].copy_from_slice(&2u16.to_le_bytes()); // versionMajor
        info_bytes[2..4].copy_from_slice(&2u16.to_le_bytes()); // versionMinor

        let header = util::parse_encryption_info_header(&info_bytes).expect("parse header");
        assert_eq!(header.kind, util::EncryptionInfoKind::Standard);

        let parsed =
            standard::parse_standard_encryption_info(&info_bytes, &header).expect("parse standard");
        assert_eq!(parsed.version_major, 2);
        assert_eq!(parsed.version_minor, 2);
    }

    #[test]
    fn parses_agile_encryption_info_minimal() {
        let info_bytes = agile::tests::agile_encryption_info_fixture();
        let header = util::parse_encryption_info_header(&info_bytes).expect("parse header");
        assert_eq!(header.kind, util::EncryptionInfoKind::Agile);
        let parsed = agile::parse_agile_encryption_info(&info_bytes, &header).expect("parse agile");
        assert_eq!(parsed.version_major, 4);
        assert_eq!(parsed.version_minor, 4);
        assert_eq!(parsed.key_data.key_bits, 256);
        assert_eq!(parsed.password_key_encryptor.spin_count, 100_000);
    }

    #[test]
    fn standard_key_derivation_matches_vector() {
        // Deterministic vector to catch regressions in key derivation.
        let password = "Password";
        let salt: [u8; 16] = [
            0x00, 0x01, 0x02, 0x03, 0x04, 0x05, 0x06, 0x07, 0x08, 0x09, 0x0A, 0x0B, 0x0C, 0x0D,
            0x0E, 0x0F,
        ];
        let deriver = StandardKeyDeriver::new(HashAlgorithm::Sha1, 128, &salt, password);
        let key_block0 = deriver.derive_key_for_block(0).expect("derive key");
        assert_eq!(
            key_block0.as_slice(),
            &[
                0x5A, 0x93, 0xE0, 0xF1, 0xBC, 0x70, 0xC5, 0xBA, 0x59, 0x46, 0x04, 0xA1, 0x5C, 0xD0,
                0xE8, 0x92,
            ]
        );
    }

    #[test]
    fn oversized_encrypted_package_size_errors_without_large_allocation() {
        let total_size: u64 = if usize::BITS < 64 {
            (usize::MAX as u64) + 1
        } else {
            u64::MAX
        };

        let mut encrypted_package = Vec::new();
        encrypted_package.extend_from_slice(&total_size.to_le_bytes());

        let dummy_standard = standard::StandardEncryptionInfo {
            version_major: 0,
            version_minor: 0,
            flags: 0,
            header: standard::EncryptionHeader {
                alg_id: 0,
                alg_id_hash: 0,
                key_bits: 0,
                provider_type: 0,
                csp_name: String::new(),
            },
            verifier: standard::EncryptionVerifier {
                salt: Vec::new(),
                encrypted_verifier: Vec::new(),
                verifier_hash_size: 0,
                encrypted_verifier_hash: Vec::new(),
            },
        };

        let dummy_agile = agile::AgileEncryptionInfo {
            version_major: 0,
            version_minor: 0,
            flags: 0,
            key_data: agile::AgileKeyData {
                salt: Vec::new(),
                block_size: 16,
                key_bits: 128,
                hash_algorithm: HashAlgorithm::Sha256,
                cipher_algorithm: String::new(),
                cipher_chaining: String::new(),
            },
            data_integrity: agile::AgileDataIntegrity {
                encrypted_hmac_key: Vec::new(),
                encrypted_hmac_value: Vec::new(),
            },
            password_key_encryptor: agile::AgilePasswordKeyEncryptor {
                salt: Vec::new(),
                block_size: 16,
                key_bits: 128,
                spin_count: 0,
                hash_algorithm: HashAlgorithm::Sha256,
                cipher_algorithm: String::new(),
                cipher_chaining: String::new(),
                encrypted_verifier_hash_input: Vec::new(),
                encrypted_verifier_hash_value: Vec::new(),
                encrypted_key_value: Vec::new(),
            },
        };

        MAX_ALLOC.store(0, Ordering::Relaxed);

        let err =
            standard::decrypt_standard_encrypted_package(&dummy_standard, &encrypted_package, "")
                .expect_err("expected size overflow");
        assert!(
            matches!(err, OfficeCryptoError::EncryptedPackageSizeOverflow { total_size: got } if got == total_size),
            "expected EncryptedPackageSizeOverflow({total_size}), got {err:?}"
        );

        let err = agile::decrypt_agile_encrypted_package(&dummy_agile, &encrypted_package, "")
            .expect_err("expected size overflow");
        assert!(
            matches!(err, OfficeCryptoError::EncryptedPackageSizeOverflow { total_size: got } if got == total_size),
            "expected EncryptedPackageSizeOverflow({total_size}), got {err:?}"
        );

        let max_alloc = MAX_ALLOC.load(Ordering::Relaxed);
        assert!(
            max_alloc < 16 * 1024 * 1024,
            "expected no large allocation attempts, observed max allocation request: {max_alloc} bytes"
        );
    }
}

#[cfg(test)]
mod test_alloc {
    use std::alloc::{GlobalAlloc, Layout, System};
    use std::sync::atomic::{AtomicUsize, Ordering};

    pub static MAX_ALLOC: AtomicUsize = AtomicUsize::new(0);

    pub struct TrackingAllocator;

    unsafe impl GlobalAlloc for TrackingAllocator {
        unsafe fn alloc(&self, layout: Layout) -> *mut u8 {
            record(layout.size());
            System.alloc(layout)
        }

        unsafe fn alloc_zeroed(&self, layout: Layout) -> *mut u8 {
            record(layout.size());
            System.alloc_zeroed(layout)
        }

        unsafe fn realloc(&self, ptr: *mut u8, layout: Layout, new_size: usize) -> *mut u8 {
            record(new_size);
            System.realloc(ptr, layout, new_size)
        }

        unsafe fn dealloc(&self, ptr: *mut u8, layout: Layout) {
            System.dealloc(ptr, layout)
        }
    }

    #[inline]
    fn record(size: usize) {
        let mut prev = MAX_ALLOC.load(Ordering::Relaxed);
        while size > prev {
            match MAX_ALLOC.compare_exchange_weak(prev, size, Ordering::Relaxed, Ordering::Relaxed)
            {
                Ok(_) => break,
                Err(next) => prev = next,
            }
        }
    }

    // Ensure tests can assert that huge `total_size` values are rejected *before*
    // attempting allocations.
    #[global_allocator]
    static GLOBAL: TrackingAllocator = TrackingAllocator;
}
