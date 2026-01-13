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
    #[error("integrity check failed")]
    IntegrityCheckFailed,
    #[error(transparent)]
    Io(#[from] std::io::Error),
}
