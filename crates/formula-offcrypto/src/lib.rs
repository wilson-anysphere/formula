use aes::cipher::{generic_array::GenericArray, BlockDecrypt, KeyInit};

#[derive(Debug, thiserror::Error)]
pub enum OffcryptoError {
    #[error("EncryptedPackage stream is too short: expected at least 8 bytes, got {len}")]
    EncryptedPackageTooShort { len: usize },

    #[error(
        "EncryptedPackage ciphertext length must be a multiple of 16 bytes for AES-ECB, got {len}"
    )]
    InvalidCiphertextLength { len: usize },

    #[error("invalid AES key length {len}; expected 16, 24, or 32 bytes")]
    InvalidKeyLength { len: usize },

    #[error("EncryptedPackage declared plaintext size {total_size} exceeds decrypted length {decrypted_len}")]
    TotalSizeOutOfBounds { total_size: u64, decrypted_len: usize },

    #[error("EncryptedPackage declared plaintext size {total_size} does not fit into usize")]
    TotalSizeTooLarge { total_size: u64 },
}

/// Decrypts the `EncryptedPackage` stream for ECMA-376 Standard encryption.
///
/// The stream layout is:
/// - `total_size` (u64, little-endian) at bytes 0..8
/// - AES-ECB ciphertext at bytes 8..
///
/// The ciphertext is decrypted in full, and the returned plaintext is truncated to `total_size`.
pub fn standard_decrypt_package(
    key: &[u8],
    encrypted_package: &[u8],
) -> Result<Vec<u8>, OffcryptoError> {
    if encrypted_package.len() < 8 {
        return Err(OffcryptoError::EncryptedPackageTooShort {
            len: encrypted_package.len(),
        });
    }

    let total_size = u64::from_le_bytes(
        encrypted_package[0..8]
            .try_into()
            .expect("slice length checked"),
    );

    let ciphertext = &encrypted_package[8..];
    if ciphertext.len() % 16 != 0 {
        return Err(OffcryptoError::InvalidCiphertextLength {
            len: ciphertext.len(),
        });
    }

    let mut decrypted = ciphertext.to_vec();
    match key.len() {
        16 => {
            let cipher = aes::Aes128::new_from_slice(key)
                .map_err(|_| OffcryptoError::InvalidKeyLength { len: key.len() })?;
            for block in decrypted.chunks_mut(16) {
                cipher.decrypt_block(GenericArray::from_mut_slice(block));
            }
        }
        24 => {
            let cipher = aes::Aes192::new_from_slice(key)
                .map_err(|_| OffcryptoError::InvalidKeyLength { len: key.len() })?;
            for block in decrypted.chunks_mut(16) {
                cipher.decrypt_block(GenericArray::from_mut_slice(block));
            }
        }
        32 => {
            let cipher = aes::Aes256::new_from_slice(key)
                .map_err(|_| OffcryptoError::InvalidKeyLength { len: key.len() })?;
            for block in decrypted.chunks_mut(16) {
                cipher.decrypt_block(GenericArray::from_mut_slice(block));
            }
        }
        _ => return Err(OffcryptoError::InvalidKeyLength { len: key.len() }),
    }

    let total_size_usize: usize = total_size
        .try_into()
        .map_err(|_| OffcryptoError::TotalSizeTooLarge { total_size })?;

    if total_size_usize > decrypted.len() {
        return Err(OffcryptoError::TotalSizeOutOfBounds {
            total_size,
            decrypted_len: decrypted.len(),
        });
    }

    decrypted.truncate(total_size_usize);
    Ok(decrypted)
}

#[cfg(test)]
mod tests {
    use super::*;
    use aes::cipher::{BlockEncrypt, KeyInit};

    #[test]
    fn standard_decrypt_package_roundtrip() {
        let key = [0x42u8; 16];

        let plaintext: Vec<u8> = (0u8..=42).collect(); // 43 bytes; exercises truncation
        let total_size = plaintext.len() as u64;

        let mut padded = plaintext.clone();
        padded.resize((padded.len() + 15) / 16 * 16, 0u8);

        let cipher = aes::Aes128::new_from_slice(&key).expect("valid AES-128 key");
        let mut ciphertext = padded.clone();
        for block in ciphertext.chunks_mut(16) {
            cipher.encrypt_block(GenericArray::from_mut_slice(block));
        }

        let mut encrypted_package = total_size.to_le_bytes().to_vec();
        encrypted_package.extend_from_slice(&ciphertext);

        let decrypted = standard_decrypt_package(&key, &encrypted_package).expect("decrypt");
        assert_eq!(decrypted, plaintext);
    }
}

