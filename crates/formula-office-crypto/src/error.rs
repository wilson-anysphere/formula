use thiserror::Error;

#[derive(Debug, Error)]
pub enum OfficeCryptoError {
    #[error("password required")]
    PasswordRequired,
    #[error("invalid password")]
    InvalidPassword,
    #[error(
        "spinCount {spin_count} exceeds maximum allowed {max} (refusing to run expensive password KDF)"
    )]
    SpinCountTooLarge { spin_count: u32, max: u32 },
    #[error("unsupported encryption: {0}")]
    UnsupportedEncryption(String),
    #[error("invalid encryption options: {0}")]
    InvalidOptions(String),
    #[error("invalid format: {0}")]
    InvalidFormat(String),
    #[error("{context} exceeds maximum allowed size ({limit} bytes)")]
    SizeLimitExceeded { context: &'static str, limit: usize },
    #[error("{context} exceeds maximum allowed size ({limit} bytes)")]
    SizeLimitExceededU64 { context: &'static str, limit: u64 },
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
