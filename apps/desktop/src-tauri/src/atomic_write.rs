use anyhow::Context;
use formula_fs::atomic_write_bytes;
use std::path::Path;

/// Atomically writes `bytes` to `path`.
///
/// This creates any missing parent directories, writes to a temp file in the
/// destination directory, then renames into place. Callers should prefer this
/// over `std::fs::write` for user-visible save paths so we never leave partially
/// written files behind if the process crashes mid-save.
pub fn write_file_atomic(path: &Path, bytes: &[u8]) -> anyhow::Result<()> {
    write_file_atomic_io(path, bytes).with_context(|| format!("atomic write {path:?}"))
}

pub(crate) fn write_file_atomic_io(path: &Path, bytes: &[u8]) -> std::io::Result<()> {
    atomic_write_bytes(path, bytes)
}

#[cfg(test)]
mod tests {
    use super::*;

    #[test]
    fn write_file_atomic_io_creates_parent_dirs() {
        let tmp = tempfile::tempdir().expect("temp dir");
        let path = tmp.path().join("nested/dir/file.bin");
        assert!(!path.parent().unwrap().exists(), "parent dir should not exist");

        write_file_atomic_io(&path, b"hello").expect("write_file_atomic_io");
        assert!(path.is_file(), "expected file to exist");
        assert_eq!(std::fs::read(&path).expect("read file"), b"hello");
    }

    #[test]
    fn write_file_atomic_io_replaces_existing_file() {
        let tmp = tempfile::tempdir().expect("temp dir");
        let path = tmp.path().join("file.bin");

        write_file_atomic_io(&path, b"old").expect("write old");
        write_file_atomic_io(&path, b"new").expect("write new");
        assert_eq!(std::fs::read(&path).expect("read file"), b"new");
    }
}
