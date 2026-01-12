use super::super::{ClipboardContent, ClipboardError, ClipboardWritePayload};

pub fn read() -> Result<ClipboardContent, ClipboardError> {
    super::super::macos::read()
}

pub fn write(payload: &ClipboardWritePayload) -> Result<(), ClipboardError> {
    super::super::macos::write(payload)
}
