#![allow(unexpected_cfgs)]

use proptest::prelude::*;
use std::panic::{catch_unwind, AssertUnwindSafe};
use std::sync::OnceLock;

use crate::crypto::HashAlgorithm;
use crate::util;

#[cfg(fuzzing)]
const CASES: u32 = 256;
#[cfg(not(fuzzing))]
const CASES: u32 = 32;

#[cfg(fuzzing)]
const MAX_LEN: usize = 256 * 1024;
#[cfg(not(fuzzing))]
const MAX_LEN: usize = 32 * 1024;

fn valid_agile_info() -> &'static crate::agile::AgileEncryptionInfo {
    static CACHE: OnceLock<crate::agile::AgileEncryptionInfo> = OnceLock::new();
    CACHE.get_or_init(|| {
        let password = "pw";
        let plaintext = b"PK\x03\x04hello";
        let opts = crate::EncryptOptions {
            scheme: crate::EncryptionScheme::Agile,
            key_bits: 128,
            hash_algorithm: HashAlgorithm::Sha1,
            spin_count: 1,
        };

        let (encryption_info, _encrypted_package) =
            crate::agile::encrypt_agile_encrypted_package(plaintext, password, &opts)
                .expect("encrypt agile fixture");
        let header = util::parse_encryption_info_header(&encryption_info).expect("parse header");
        crate::agile::parse_agile_encryption_info(&encryption_info, &header).expect("parse agile")
    })
}

proptest! {
    #![proptest_config(ProptestConfig {
        cases: CASES,
        max_shrink_iters: 0,
        .. ProptestConfig::default()
    })]

    #[test]
    fn parse_agile_encryption_info_is_panic_free_and_rejects_garbage(tail in prop::collection::vec(any::<u8>(), 0..=MAX_LEN)) {
        // Build an `EncryptionInfo` stream that is guaranteed to be treated as Agile (4.4) and to
        // contain invalid UTF-8 XML bytes.
        //
        // Ensure we are *not* misdetected as a length-prefixed descriptor by writing a huge
        // candidate length at offset 8.
        let mut bytes = Vec::with_capacity(8 + 4 + 2 + tail.len());
        bytes.extend_from_slice(&4u16.to_le_bytes());
        bytes.extend_from_slice(&4u16.to_le_bytes());
        bytes.extend_from_slice(&0u32.to_le_bytes()); // flags
        bytes.extend_from_slice(&u32::MAX.to_le_bytes()); // bogus length prefix (forces header_offset=8)
        bytes.push(b'<');
        bytes.push(0xFF); // invalid UTF-8
        bytes.extend_from_slice(&tail);

        let outcome = catch_unwind(AssertUnwindSafe(|| {
            let header = util::parse_encryption_info_header(&bytes).expect("header should parse");
            crate::agile::parse_agile_encryption_info(&bytes, &header)
        }));
        prop_assert!(outcome.is_ok(), "parse_agile_encryption_info panicked");
        prop_assert!(outcome.unwrap().is_err(), "garbage input should not parse");
    }

    #[test]
    fn decrypt_agile_encrypted_package_is_panic_free_and_rejects_garbage_ciphertext(
        len_matches in any::<bool>(),
        mut ciphertext in prop::collection::vec(any::<u8>(), 0..=MAX_LEN),
    ) {
        // Ensure AES-CBC framing is valid so we exercise more of the decrypt path.
        ciphertext.truncate(ciphertext.len() - (ciphertext.len() % 16));

        let declared_len = if len_matches {
            ciphertext.len() as u64
        } else {
            (ciphertext.len() as u64).saturating_add(1)
        };

        let mut encrypted_package = Vec::with_capacity(8 + ciphertext.len());
        encrypted_package.extend_from_slice(&declared_len.to_le_bytes());
        encrypted_package.extend_from_slice(&ciphertext);

        let info = valid_agile_info();
        let outcome = catch_unwind(AssertUnwindSafe(|| {
            crate::agile::decrypt_agile_encrypted_package(
                info,
                &encrypted_package,
                "pw",
                &crate::DecryptOptions::default(),
            )
        }));
        prop_assert!(outcome.is_ok(), "decrypt_agile_encrypted_package panicked");
        prop_assert!(outcome.unwrap().is_err(), "garbage input should not decrypt");
    }
}

