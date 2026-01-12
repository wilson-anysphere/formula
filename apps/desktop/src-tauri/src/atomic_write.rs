use anyhow::Context;
use std::io::Write;
use std::path::Path;
use tempfile::NamedTempFile;

fn sync_dir_best_effort(path: &Path) {
    // Syncing a directory is platform/filesystem dependent; this is strictly best-effort.
    // On Unix this helps make the final rename durable after a crash/power loss.
    if let Ok(dir) = std::fs::File::open(path) {
        let _ = dir.sync_all();
    }
}

/// Atomically writes `bytes` to `path`.
///
/// This creates any missing parent directories, writes to a temp file in the
/// destination directory, then renames into place. Callers should prefer this
/// over `std::fs::write` for user-visible save paths so we never leave partially
/// written files behind if the process crashes mid-save.
pub(crate) fn write_file_atomic_io(path: &Path, bytes: &[u8]) -> std::io::Result<()> {
    // `Path::parent` returns `None` for paths like `foo.xlsx` (no separators).
    // In that case, treat the current directory as the parent so the temp file
    // is still created alongside the destination.
    let parent = path.parent().unwrap_or_else(|| Path::new("."));
    std::fs::create_dir_all(parent)?;

    let mut tmp =
        NamedTempFile::new_in(parent)?;
    tmp.as_file_mut().write_all(bytes)?;
    tmp.as_file_mut().flush()?;

    // Best-effort durability: ensure bytes have hit the OS before the rename.
    tmp.as_file().sync_all()?;

    match tmp.persist(path) {
        Ok(_) => {
            // Best-effort: ensure the directory entry for the new file is durable.
            sync_dir_best_effort(parent);
            Ok(())
        }
        Err(err) if err.error.kind() == std::io::ErrorKind::AlreadyExists => {
            // Best-effort replacement on platforms/filesystems where rename doesn't clobber.
            let _ = std::fs::remove_file(path);
            err.file.persist(path).map(|_| ()).map_err(|e| e.error)?;
            sync_dir_best_effort(parent);
            Ok(())
        }
        Err(err) => Err(err.error),
    }
}

pub fn write_file_atomic(path: &Path, bytes: &[u8]) -> anyhow::Result<()> {
    write_file_atomic_io(path, bytes).with_context(|| format!("write file atomically to {path:?}"))
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
