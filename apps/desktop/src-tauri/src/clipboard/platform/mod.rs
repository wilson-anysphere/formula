use super::{ClipboardContent, ClipboardError, ClipboardWritePayload};

#[cfg(target_os = "windows")]
mod windows;
#[cfg(target_os = "macos")]
mod macos;
#[cfg(target_os = "linux")]
mod linux;
#[cfg(not(any(target_os = "windows", target_os = "macos", target_os = "linux")))]
mod unsupported;

#[cfg(target_os = "windows")]
pub fn read() -> Result<ClipboardContent, ClipboardError> {
    windows::read()
}

#[cfg(target_os = "windows")]
pub fn write(payload: &ClipboardWritePayload) -> Result<(), ClipboardError> {
    windows::write(payload)
}

#[cfg(target_os = "macos")]
pub fn read() -> Result<ClipboardContent, ClipboardError> {
    macos::read()
}

#[cfg(target_os = "macos")]
pub fn write(payload: &ClipboardWritePayload) -> Result<(), ClipboardError> {
    macos::write(payload)
}

#[cfg(target_os = "linux")]
pub fn read() -> Result<ClipboardContent, ClipboardError> {
    linux::read()
}

#[cfg(target_os = "linux")]
pub fn write(payload: &ClipboardWritePayload) -> Result<(), ClipboardError> {
    linux::write(payload)
}

#[cfg(not(any(target_os = "windows", target_os = "macos", target_os = "linux")))]
pub fn read() -> Result<ClipboardContent, ClipboardError> {
    unsupported::read()
}

#[cfg(not(any(target_os = "windows", target_os = "macos", target_os = "linux")))]
pub fn write(payload: &ClipboardWritePayload) -> Result<(), ClipboardError> {
    unsupported::write(payload)
}
