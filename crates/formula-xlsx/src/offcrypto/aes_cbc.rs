use aes::{Aes128, Aes192, Aes256};
use cbc::Decryptor;
use cipher::block_padding::NoPadding;
use cipher::{BlockDecryptMut, KeyIvInit};
use thiserror::Error;

pub const AES_BLOCK_SIZE: usize = 16;

#[derive(Debug, Error, PartialEq, Eq)]
pub enum AesCbcDecryptError {
    #[error("unsupported AES key length: {0} bytes (expected 16, 24, or 32)")]
    UnsupportedKeyLength(usize),
    #[error("invalid AES-CBC IV length: {0} bytes (expected 16)")]
    InvalidIvLength(usize),
    #[error("ciphertext length is not a multiple of 16 bytes: {0}")]
    InvalidCiphertextLength(usize),
}

/// Decrypt AES-CBC ciphertext without padding removal.
///
/// MS-OFFCRYPTO "Agile" encryption uses AES-CBC with plaintext pre-padded to a whole number of
/// blocks. The encrypted buffers are always a multiple of 16 bytes, and the caller truncates the
/// decrypted output to the semantic length stored elsewhere in the format.
pub fn decrypt_aes_cbc_no_padding(
    key: &[u8],
    iv: &[u8],
    ciphertext: &[u8],
) -> Result<Vec<u8>, AesCbcDecryptError> {
    let mut out = ciphertext.to_vec();
    decrypt_aes_cbc_no_padding_in_place(key, iv, &mut out)?;
    Ok(out)
}

/// In-place AES-CBC decryption without padding removal.
pub fn decrypt_aes_cbc_no_padding_in_place(
    key: &[u8],
    iv: &[u8],
    buf: &mut [u8],
) -> Result<(), AesCbcDecryptError> {
    if iv.len() != AES_BLOCK_SIZE {
        return Err(AesCbcDecryptError::InvalidIvLength(iv.len()));
    }
    if buf.is_empty() {
        return Ok(());
    }
    let buf_len = buf.len();
    if buf_len % AES_BLOCK_SIZE != 0 {
        return Err(AesCbcDecryptError::InvalidCiphertextLength(buf_len));
    }

    match key.len() {
        16 => {
            let dec = Decryptor::<Aes128>::new_from_slices(key, iv).map_err(|_| {
                    // We already validated the IV length, so treat `InvalidLength` as a key issue.
                    AesCbcDecryptError::UnsupportedKeyLength(key.len())
            })?;
            dec.decrypt_padded_mut::<NoPadding>(buf)
                .map_err(|_| AesCbcDecryptError::InvalidCiphertextLength(buf_len))?;
        }
        24 => {
            let dec = Decryptor::<Aes192>::new_from_slices(key, iv)
                .map_err(|_| AesCbcDecryptError::UnsupportedKeyLength(key.len()))?;
            dec.decrypt_padded_mut::<NoPadding>(buf)
                .map_err(|_| AesCbcDecryptError::InvalidCiphertextLength(buf_len))?;
        }
        32 => {
            let dec = Decryptor::<Aes256>::new_from_slices(key, iv)
                .map_err(|_| AesCbcDecryptError::UnsupportedKeyLength(key.len()))?;
            dec.decrypt_padded_mut::<NoPadding>(buf)
                .map_err(|_| AesCbcDecryptError::InvalidCiphertextLength(buf_len))?;
        }
        other => return Err(AesCbcDecryptError::UnsupportedKeyLength(other)),
    }

    Ok(())
}

#[cfg(test)]
mod tests {
    use super::decrypt_aes_cbc_no_padding;
    use super::decrypt_aes_cbc_no_padding_in_place;
    use super::AesCbcDecryptError;

    fn hex_decode(s: &str) -> Vec<u8> {
        let s = s.trim();
        assert!(
            s.len() % 2 == 0,
            "hex string must have even length (got {})",
            s.len()
        );
        (0..s.len())
            .step_by(2)
            .map(|i| u8::from_str_radix(&s[i..i + 2], 16).expect("valid hex byte"))
            .collect()
    }

    // NIST SP 800-38A F.2.1/F.2.3/F.2.5 test vectors for AES-CBC.
    const PLAINTEXT: &str = concat!(
        "6bc1bee22e409f96e93d7e117393172a",
        "ae2d8a571e03ac9c9eb76fac45af8e51",
        "30c81c46a35ce411e5fbc1191a0a52ef",
        "f69f2445df4f9b17ad2b417be66c3710",
    );
    const IV: &str = "000102030405060708090a0b0c0d0e0f";

    #[test]
    fn aes_128_cbc_decrypt_no_padding_matches_nist_vector() {
        let key = hex_decode("2b7e151628aed2a6abf7158809cf4f3c");
        let ciphertext = hex_decode(concat!(
            "7649abac8119b246cee98e9b12e9197d",
            "5086cb9b507219ee95db113a917678b2",
            "73bed6b8e3c1743b7116e69e22229516",
            "3ff1caa1681fac09120eca307586e1a7",
        ));

        let plaintext = decrypt_aes_cbc_no_padding(&key, &hex_decode(IV), &ciphertext)
            .expect("decrypt should succeed");
        assert_eq!(plaintext, hex_decode(PLAINTEXT));
    }

    #[test]
    fn aes_192_cbc_decrypt_no_padding_matches_nist_vector() {
        let key = hex_decode("8e73b0f7da0e6452c810f32b809079e562f8ead2522c6b7b");
        let ciphertext = hex_decode(concat!(
            "4f021db243bc633d7178183a9fa071e8",
            "b4d9ada9ad7dedf4e5e738763f69145a",
            "571b242012fb7ae07fa9baac3df102e0",
            "08b0e27988598881d920a9e64f5615cd",
        ));

        let plaintext = decrypt_aes_cbc_no_padding(&key, &hex_decode(IV), &ciphertext)
            .expect("decrypt should succeed");
        assert_eq!(plaintext, hex_decode(PLAINTEXT));
    }

    #[test]
    fn aes_256_cbc_decrypt_no_padding_matches_nist_vector() {
        let key = hex_decode(concat!(
            "603deb1015ca71be2b73aef0857d7781",
            "1f352c073b6108d72d9810a30914dff4",
        ));
        let ciphertext = hex_decode(concat!(
            "f58c4c04d6e5f1ba779eabfb5f7bfbd6",
            "9cfc4e967edb808d679f777bc6702c7d",
            "39f23369a9d9bacfa530e26304231461",
            "b2eb05e2c39be9fcda6c19078c6a9d1b",
        ));

        let plaintext = decrypt_aes_cbc_no_padding(&key, &hex_decode(IV), &ciphertext)
            .expect("decrypt should succeed");
        assert_eq!(plaintext, hex_decode(PLAINTEXT));
    }

    #[test]
    fn decrypt_errors_on_unsupported_key_length() {
        let key = [0u8; 17];
        let iv = [0u8; 16];
        let ciphertext = [0u8; 16];

        let err = decrypt_aes_cbc_no_padding(&key, &iv, &ciphertext).expect_err("should fail");
        assert_eq!(err, AesCbcDecryptError::UnsupportedKeyLength(17));
    }

    #[test]
    fn decrypt_errors_on_non_block_multiple_ciphertext() {
        let key = [0u8; 16];
        let iv = [0u8; 16];
        let mut buf = [0u8; 15];

        let err =
            decrypt_aes_cbc_no_padding_in_place(&key, &iv, &mut buf).expect_err("should fail");
        assert_eq!(err, AesCbcDecryptError::InvalidCiphertextLength(15));
    }
}
