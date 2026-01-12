use super::super::{ClipboardContent, ClipboardError, ClipboardWritePayload};

pub fn read() -> Result<ClipboardContent, ClipboardError> {
    Err(ClipboardError::OperationFailed(
        "clipboard_read not implemented on macOS".to_string(),
    ))
}

pub fn write(_payload: &ClipboardWritePayload) -> Result<(), ClipboardError> {
    Err(ClipboardError::OperationFailed(
        "clipboard_write not implemented on macOS".to_string(),
    ))
}
