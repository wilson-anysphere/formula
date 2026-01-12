use super::super::{ClipboardContent, ClipboardError, ClipboardWritePayload};

pub fn read() -> Result<ClipboardContent, ClipboardError> {
    super::super::linux::read()
}

pub fn write(payload: &ClipboardWritePayload) -> Result<(), ClipboardError> {
    super::super::linux::write(payload)
}
