use super::super::{ClipboardContent, ClipboardError, ClipboardWritePayload};

#[cfg(feature = "desktop")]
pub fn read() -> Result<ClipboardContent, ClipboardError> {
    super::super::linux::read()
}

#[cfg(feature = "desktop")]
pub fn write(payload: &ClipboardWritePayload) -> Result<(), ClipboardError> {
    super::super::linux::write(payload)
}

#[cfg(not(feature = "desktop"))]
pub fn read() -> Result<ClipboardContent, ClipboardError> {
    Err(ClipboardError::Unavailable(
        "Linux clipboard backend requires the `desktop` feature".to_string(),
    ))
}

#[cfg(not(feature = "desktop"))]
pub fn write(_payload: &ClipboardWritePayload) -> Result<(), ClipboardError> {
    Err(ClipboardError::Unavailable(
        "Linux clipboard backend requires the `desktop` feature".to_string(),
    ))
}
