use super::super::{ClipboardContent, ClipboardError, ClipboardWritePayload};

pub fn read() -> Result<ClipboardContent, ClipboardError> {
    Err(ClipboardError::UnsupportedPlatform)
}

pub fn write(_payload: &ClipboardWritePayload) -> Result<(), ClipboardError> {
    Err(ClipboardError::UnsupportedPlatform)
}
