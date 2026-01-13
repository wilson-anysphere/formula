use aes::cipher::{generic_array::GenericArray, BlockEncrypt, KeyInit};
use aes::{Aes128, Aes192, Aes256};
use formula_offcrypto::{
    standard_derive_key, standard_verify_key, OffcryptoError, StandardEncryptionHeader,
    StandardEncryptionInfo, StandardEncryptionVerifier,
};
use sha1::{Digest as _, Sha1};

// Known test vector from `msoffcrypto/method/ecma376_standard.py` docstrings.
const PASSWORD: &str = "Password1234_";
const SALT: [u8; 16] = [
    0xe8, 0x82, 0x66, 0x49, 0x0c, 0x5b, 0xd1, 0xee, 0xbd, 0x2b, 0x43, 0x94, 0xe3, 0xf8, 0x30, 0xef,
];
const EXPECTED_KEY_128: [u8; 16] = [
    0x40, 0xb1, 0x3a, 0x71, 0xf9, 0x0b, 0x96, 0x6e, 0x37, 0x54, 0x08, 0xf2, 0xd1, 0x81, 0xa1, 0xaa,
];
const ENCRYPTED_VERIFIER: [u8; 16] = [
    0x51, 0x6f, 0x73, 0x2e, 0x96, 0x6f, 0xac, 0x17, 0xb1, 0xc5, 0xd7, 0xd8, 0xcc, 0x36, 0xc9, 0x28,
];
const ENCRYPTED_VERIFIER_HASH: [u8; 32] = [
    0x2b, 0x61, 0x68, 0xda, 0xbe, 0x29, 0x11, 0xad, 0x2b, 0xd3, 0x7c, 0x17, 0x46, 0x74, 0x5c, 0x14,
    0xd3, 0xcf, 0x1b, 0xb1, 0x40, 0xa4, 0x8f, 0x4e, 0x6f, 0x3d, 0x23, 0x88, 0x08, 0x72, 0xb1, 0x6a,
];

fn standard_info() -> StandardEncryptionInfo {
    StandardEncryptionInfo {
        header: StandardEncryptionHeader {
            flags: 0,
            size_extra: 0,
            alg_id: 0x0000_660E,
            alg_id_hash: 0x0000_8004, // CALG_SHA1
            key_size_bits: 128,
            provider_type: 0x0000_0018, // PROV_RSA_AES
            reserved1: 0,
            reserved2: 0,
            csp_name: String::new(),
        },
        verifier: StandardEncryptionVerifier {
            salt: SALT.to_vec(),
            encrypted_verifier: ENCRYPTED_VERIFIER,
            verifier_hash_size: 20,
            encrypted_verifier_hash: ENCRYPTED_VERIFIER_HASH.to_vec(),
        },
    }
}

fn aes_ecb_encrypt_in_place(key: &[u8], buf: &mut [u8]) {
    assert_eq!(buf.len() % 16, 0);
    match key.len() {
        16 => {
            let cipher = Aes128::new_from_slice(key).expect("valid AES-128 key");
            for block in buf.chunks_mut(16) {
                cipher.encrypt_block(GenericArray::from_mut_slice(block));
            }
        }
        24 => {
            let cipher = Aes192::new_from_slice(key).expect("valid AES-192 key");
            for block in buf.chunks_mut(16) {
                cipher.encrypt_block(GenericArray::from_mut_slice(block));
            }
        }
        32 => {
            let cipher = Aes256::new_from_slice(key).expect("valid AES-256 key");
            for block in buf.chunks_mut(16) {
                cipher.encrypt_block(GenericArray::from_mut_slice(block));
            }
        }
        _ => panic!("unexpected AES key length"),
    }
}

#[test]
fn standard_derive_key_matches_msoffcrypto_vector() {
    let info = standard_info();
    let key = standard_derive_key(&info, PASSWORD).expect("derive key");
    assert_eq!(key.as_slice(), &EXPECTED_KEY_128);
}

#[test]
fn standard_verify_key_matches_msoffcrypto_vector() {
    let info = standard_info();
    let key = standard_derive_key(&info, PASSWORD).expect("derive key");
    standard_verify_key(&info, &key).expect("verify key");

    let wrong_key = [0u8; 16];
    assert_eq!(
        standard_verify_key(&info, &wrong_key),
        Err(OffcryptoError::InvalidPassword)
    );
}

#[test]
fn standard_verify_key_accepts_correct_key_and_rejects_incorrect_key() {
    let base_info = StandardEncryptionInfo {
        header: StandardEncryptionHeader {
            flags: 0,
            size_extra: 0,
            alg_id: 0x0000_660E,
            alg_id_hash: 0x0000_8004, // CALG_SHA1
            key_size_bits: 128,
            provider_type: 0x0000_0018, // PROV_RSA_AES
            reserved1: 0,
            reserved2: 0,
            csp_name: String::new(),
        },
        verifier: StandardEncryptionVerifier {
            salt: SALT.to_vec(),
            encrypted_verifier: [0u8; 16],
            verifier_hash_size: 20,
            encrypted_verifier_hash: vec![0u8; 32],
        },
    };

    let key = standard_derive_key(&base_info, PASSWORD).expect("derive key");

    let verifier_plain: [u8; 16] = [
        0x00, 0x01, 0x02, 0x03, 0x04, 0x05, 0x06, 0x07, 0x08, 0x09, 0x0a, 0x0b, 0x0c, 0x0d, 0x0e, 0x0f,
    ];
    let verifier_hash: [u8; 20] = Sha1::digest(&verifier_plain).into();

    let mut verifier_hash_padded = [0u8; 32];
    verifier_hash_padded[..20].copy_from_slice(&verifier_hash);
    verifier_hash_padded[20..].fill(0xa5);

    let mut encrypted_verifier = verifier_plain;
    aes_ecb_encrypt_in_place(&key, &mut encrypted_verifier);

    let mut encrypted_verifier_hash = verifier_hash_padded;
    aes_ecb_encrypt_in_place(&key, &mut encrypted_verifier_hash);

    let mut info = base_info;
    info.verifier.encrypted_verifier = encrypted_verifier;
    info.verifier.encrypted_verifier_hash = encrypted_verifier_hash.to_vec();

    standard_verify_key(&info, &key).expect("verify key");

    let wrong_key = [0u8; 16];
    assert_eq!(
        standard_verify_key(&info, &wrong_key),
        Err(OffcryptoError::InvalidPassword)
    );
}

#[test]
fn standard_derive_key_rejects_non_sha1_alg_id_hash() {
    let info = StandardEncryptionInfo {
        header: StandardEncryptionHeader {
            flags: 0,
            size_extra: 0,
            alg_id: 0x0000_660E,
            alg_id_hash: 0, // not CALG_SHA1
            key_size_bits: 128,
            provider_type: 0x0000_0018, // PROV_RSA_AES
            reserved1: 0,
            reserved2: 0,
            csp_name: String::new(),
        },
        verifier: StandardEncryptionVerifier {
            salt: SALT.to_vec(),
            encrypted_verifier: [0u8; 16],
            verifier_hash_size: 20,
            encrypted_verifier_hash: vec![0u8; 32],
        },
    };

    let err = standard_derive_key(&info, PASSWORD).unwrap_err();
    assert_eq!(err, OffcryptoError::UnsupportedAlgorithm(0));
}

#[test]
fn standard_derive_key_rejects_key_size_mismatch() {
    let info = StandardEncryptionInfo {
        header: StandardEncryptionHeader {
            flags: 0,
            size_extra: 0,
            alg_id: 0x0000_660E,
            alg_id_hash: 0x0000_8004, // CALG_SHA1
            key_size_bits: 256,       // mismatched
            provider_type: 0x0000_0018, // PROV_RSA_AES
            reserved1: 0,
            reserved2: 0,
            csp_name: String::new(),
        },
        verifier: StandardEncryptionVerifier {
            salt: SALT.to_vec(),
            encrypted_verifier: [0u8; 16],
            verifier_hash_size: 20,
            encrypted_verifier_hash: vec![0u8; 32],
        },
    };

    let err = standard_derive_key(&info, PASSWORD).unwrap_err();
    assert_eq!(err, OffcryptoError::UnsupportedAlgorithm(0x0000_660E));
}

#[test]
fn standard_verify_key_rejects_invalid_salt_len() {
    let info = StandardEncryptionInfo {
        header: StandardEncryptionHeader {
            flags: 0,
            size_extra: 0,
            alg_id: 0x0000_660E,
            alg_id_hash: 0x0000_8004, // CALG_SHA1
            key_size_bits: 128,
            provider_type: 0x0000_0018, // PROV_RSA_AES
            reserved1: 0,
            reserved2: 0,
            csp_name: String::new(),
        },
        verifier: StandardEncryptionVerifier {
            salt: vec![0u8; 15],
            encrypted_verifier: [0u8; 16],
            verifier_hash_size: 20,
            encrypted_verifier_hash: vec![0u8; 32],
        },
    };

    let err = standard_verify_key(&info, &[0u8; 16]).unwrap_err();
    assert_eq!(
        err,
        OffcryptoError::InvalidEncryptionInfo {
            context: "EncryptionVerifier.saltSize must be 16 for Standard encryption"
        }
    );
}
