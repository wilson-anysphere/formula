use super::super::{ClipboardContent, ClipboardError, ClipboardWritePayload};

#[cfg(feature = "desktop")]
pub fn read() -> Result<ClipboardContent, ClipboardError> {
    super::super::macos::read()
}

#[cfg(feature = "desktop")]
pub fn write(payload: &ClipboardWritePayload) -> Result<(), ClipboardError> {
    super::super::macos::write(payload)
}

#[cfg(not(feature = "desktop"))]
pub fn read() -> Result<ClipboardContent, ClipboardError> {
    Err(ClipboardError::Unavailable(
        "macOS clipboard backend requires the `desktop` feature".to_string(),
    ))
}

#[cfg(not(feature = "desktop"))]
pub fn write(_payload: &ClipboardWritePayload) -> Result<(), ClipboardError> {
    Err(ClipboardError::Unavailable(
        "macOS clipboard backend requires the `desktop` feature".to_string(),
    ))
}
