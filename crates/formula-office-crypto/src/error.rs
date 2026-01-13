use thiserror::Error;

#[derive(Debug, Error)]
pub enum OfficeCryptoError {
    #[error("password required")]
    PasswordRequired,
    #[error("invalid password")]
    InvalidPassword,
    #[error("unsupported encryption: {0}")]
    UnsupportedEncryption(String),
    #[error("invalid encryption options: {0}")]
    InvalidOptions(String),
    #[error("invalid format: {0}")]
    InvalidFormat(String),
    #[error("encrypted package size {total_size} overflows supported allocations")]
    EncryptedPackageSizeOverflow { total_size: u64 },
    #[error("failed to allocate decrypted package buffer of size {total_size}: {source}")]
    EncryptedPackageAllocationFailed {
        total_size: u64,
        #[source]
        source: std::collections::TryReserveError,
    },
    #[error("integrity check failed")]
    IntegrityCheckFailed,
    #[error(transparent)]
    Io(#[from] std::io::Error),
}
