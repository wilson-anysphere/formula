use rc4::cipher::{KeyInit, StreamCipher};
use rc4::Rc4;

/// Error returned when the provided RC4 key length is not supported by this helper.
///
/// Note: The underlying RC4 algorithm supports 1..=256-byte keys. This helper currently dispatches
/// a small set of key lengths we need for Office (and for the unit test vectors).
#[allow(dead_code)]
#[derive(Debug, thiserror::Error, PartialEq, Eq)]
#[error("unsupported RC4 key length: {0}")]
pub(crate) struct UnsupportedRc4KeyLength(pub(crate) usize);

fn rc4_xor_in_place_with_cipher<C: KeyInit + StreamCipher>(
    key: &[u8],
    data: &mut [u8],
) -> Result<(), UnsupportedRc4KeyLength> {
    let mut cipher = C::new_from_slice(key).map_err(|_| {
        debug_assert!(
            false,
            "RC4 cipher rejected key length {}",
            key.len()
        );
        UnsupportedRc4KeyLength(key.len())
    })?;
    cipher.apply_keystream(data);
    Ok(())
}

/// Apply the RC4 keystream (XOR) to `data` in-place using `key`.
///
/// RC4 encryption and decryption are the same operation: `ciphertext = plaintext XOR keystream`.
#[allow(dead_code)]
pub(crate) fn rc4_xor_in_place(key: &[u8], data: &mut [u8]) -> Result<(), UnsupportedRc4KeyLength> {
    use rc4::cipher::consts::{U16, U3, U4, U5, U6, U7};

    match key.len() {
        3 => rc4_xor_in_place_with_cipher::<Rc4<U3>>(key, data)?,
        4 => rc4_xor_in_place_with_cipher::<Rc4<U4>>(key, data)?,
        5 => rc4_xor_in_place_with_cipher::<Rc4<U5>>(key, data)?,
        6 => rc4_xor_in_place_with_cipher::<Rc4<U6>>(key, data)?,
        7 => rc4_xor_in_place_with_cipher::<Rc4<U7>>(key, data)?,
        16 => rc4_xor_in_place_with_cipher::<Rc4<U16>>(key, data)?,
        other => return Err(UnsupportedRc4KeyLength(other)),
    }

    Ok(())
}

#[cfg(test)]
mod tests {
    use super::rc4_xor_in_place;

    #[test]
    fn rc4_vectors_encrypt_decrypt_symmetry() {
        // Vectors from Wikipedia / common RC4 test vector sets.
        //
        // Raw RC4 (no drop).
        let cases: &[(&[u8], &[u8], &[u8])] = &[
            (
                b"Key",
                b"Plaintext",
                &[0xbb, 0xf3, 0x16, 0xe8, 0xd9, 0x40, 0xaf, 0x0a, 0xd3],
            ),
            (b"Wiki", b"pedia", &[0x10, 0x21, 0xbf, 0x04, 0x20]),
            (
                b"Secret",
                b"Attack at dawn",
                &[
                    0x45, 0xa0, 0x1f, 0x64, 0x5f, 0xc3, 0x5b, 0x38, 0x35, 0x52, 0x54, 0x4b,
                    0x9b, 0xf5,
                ],
            ),
        ];

        for (key, plaintext, expected_ciphertext) in cases {
            // Encrypt.
            let mut ciphertext = plaintext.to_vec();
            rc4_xor_in_place(key, &mut ciphertext).expect("RC4 encrypt");
            assert_eq!(
                ciphertext.as_slice(),
                *expected_ciphertext,
                "encrypt key={:?} plaintext={:?}",
                std::str::from_utf8(key).ok(),
                std::str::from_utf8(plaintext).ok()
            );

            // Decrypt (same operation, new cipher instance).
            rc4_xor_in_place(key, &mut ciphertext).expect("RC4 decrypt");
            assert_eq!(
                ciphertext.as_slice(),
                *plaintext,
                "decrypt key={:?}",
                std::str::from_utf8(key).ok()
            );
        }
    }
}
