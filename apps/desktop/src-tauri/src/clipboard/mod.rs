use serde::{Deserialize, Serialize};

#[derive(Clone, Debug, Default, Serialize, Deserialize, PartialEq)]
pub struct ClipboardContent {
    #[serde(skip_serializing_if = "Option::is_none")]
    pub text: Option<String>,
    #[serde(skip_serializing_if = "Option::is_none")]
    pub html: Option<String>,
    #[serde(skip_serializing_if = "Option::is_none")]
    pub rtf: Option<String>,
    /// PNG bytes encoded as base64.
    #[serde(skip_serializing_if = "Option::is_none")]
    pub image_png_base64: Option<String>,
}

#[derive(Clone, Debug, Serialize, Deserialize, PartialEq)]
pub struct ClipboardWritePayload {
    pub text: String,
    #[serde(skip_serializing_if = "Option::is_none")]
    pub html: Option<String>,
    #[serde(skip_serializing_if = "Option::is_none")]
    pub rtf: Option<String>,
    /// PNG bytes encoded as base64.
    #[serde(skip_serializing_if = "Option::is_none")]
    pub image_png_base64: Option<String>,
}

#[derive(Debug, thiserror::Error)]
pub enum ClipboardError {
    #[error("clipboard is not supported on this platform")]
    UnsupportedPlatform,
    #[error("clipboard backend is unavailable: {0}")]
    Unavailable(String),
    #[error("clipboard payload is invalid: {0}")]
    InvalidPayload(String),
    #[error("clipboard operation failed: {0}")]
    OperationFailed(String),
}

#[cfg(target_os = "linux")]
mod linux;

pub fn read() -> Result<ClipboardContent, ClipboardError> {
    #[cfg(target_os = "linux")]
    {
        return linux::read();
    }

    #[allow(unreachable_code)]
    Err(ClipboardError::UnsupportedPlatform)
}

pub fn write(payload: &ClipboardWritePayload) -> Result<(), ClipboardError> {
    #[cfg(target_os = "linux")]
    {
        return linux::write(payload);
    }

    #[allow(unreachable_code)]
    Err(ClipboardError::UnsupportedPlatform)
}
