use super::super::{ClipboardContent, ClipboardError, ClipboardWritePayload};

pub fn read() -> Result<ClipboardContent, ClipboardError> {
    super::super::windows::read()
}

pub fn write(payload: &ClipboardWritePayload) -> Result<(), ClipboardError> {
    super::super::windows::write(payload)
}
