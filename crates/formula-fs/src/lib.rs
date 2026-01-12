//! Small filesystem utilities shared across workspace crates.
//!
//! In particular, this provides helpers for atomic file writes:
//! - write to a temp file in the same directory (avoids cross-device renames)
//! - flush + `sync_all`
//! - rename into place with replace semantics (including on Windows)

use std::fs::{self, File};
use std::io::{self, Write};
use std::path::{Path, PathBuf};

use tempfile::NamedTempFile;

#[derive(Debug)]
pub enum AtomicWriteError<E> {
    Io(io::Error),
    Writer(E),
}

impl<E> From<io::Error> for AtomicWriteError<E> {
    fn from(err: io::Error) -> Self {
        Self::Io(err)
    }
}

impl<E: std::fmt::Display> std::fmt::Display for AtomicWriteError<E> {
    fn fmt(&self, f: &mut std::fmt::Formatter<'_>) -> std::fmt::Result {
        match self {
            AtomicWriteError::Io(err) => write!(f, "io error: {err}"),
            AtomicWriteError::Writer(err) => write!(f, "write error: {err}"),
        }
    }
}

impl<E: std::error::Error + 'static> std::error::Error for AtomicWriteError<E> {
    fn source(&self) -> Option<&(dyn std::error::Error + 'static)> {
        match self {
            AtomicWriteError::Io(err) => Some(err),
            AtomicWriteError::Writer(err) => Some(err),
        }
    }
}

fn parent_dir_or_dot(path: &Path) -> &Path {
    // `Path::parent` returns `Some("")` for bare relative file names like `foo.xlsx`.
    // Treat that as the current directory so callers can use relative paths without
    // having to prepend `./`.
    path.parent()
        .filter(|p| !p.as_os_str().is_empty())
        .unwrap_or_else(|| Path::new("."))
}

/// Atomically write a file by:
/// - creating parent directories (if needed)
/// - writing to a temp file in the same directory
/// - flushing + syncing the temp file
/// - renaming it into place with replace semantics
///
/// If `write_fn` returns an error, the destination file is left untouched.
pub fn atomic_write<T, E>(
    dest: impl AsRef<Path>,
    write_fn: impl FnOnce(&mut File) -> Result<T, E>,
) -> Result<T, AtomicWriteError<E>> {
    let dest = dest.as_ref();
    let dir = parent_dir_or_dot(dest);
    fs::create_dir_all(dir).map_err(AtomicWriteError::Io)?;

    let mut tmp = NamedTempFile::new_in(dir).map_err(AtomicWriteError::Io)?;
    let out = write_fn(tmp.as_file_mut()).map_err(AtomicWriteError::Writer)?;

    tmp.as_file_mut().flush().map_err(AtomicWriteError::Io)?;
    tmp.as_file().sync_all().map_err(AtomicWriteError::Io)?;

    let tmp_path = tmp.into_temp_path();
    replace_file(tmp_path.as_ref(), dest).map_err(AtomicWriteError::Io)?;

    // Best-effort: sync directory metadata after the rename.
    // Failures here should not be treated as a write failure (the file is already in place).
    let _ = sync_parent_dir(dest);

    Ok(out)
}

/// Like [`atomic_write`], but passes a temp file *path* to the closure.
///
/// This is useful for libraries that only offer `save_as(path)` APIs.
///
/// Note: The temp file already exists when `write_fn` is called. `write_fn` should be prepared
/// to overwrite/truncate it (e.g. via `File::create`).
pub fn atomic_write_with_path<T, E>(
    dest: impl AsRef<Path>,
    write_fn: impl FnOnce(&Path) -> Result<T, E>,
) -> Result<T, AtomicWriteError<E>> {
    let dest = dest.as_ref();
    let dir = parent_dir_or_dot(dest);
    fs::create_dir_all(dir).map_err(AtomicWriteError::Io)?;

    let tmp = NamedTempFile::new_in(dir).map_err(AtomicWriteError::Io)?;
    let tmp_path = tmp.into_temp_path();
    let tmp_path_ref: &Path = <tempfile::TempPath as AsRef<Path>>::as_ref(&tmp_path);

    let out = write_fn(tmp_path_ref).map_err(AtomicWriteError::Writer)?;

    // Ensure the temp file's contents are durably flushed before the rename.
    File::open(tmp_path_ref)
        .and_then(|f| f.sync_all())
        .map_err(AtomicWriteError::Io)?;

    replace_file(tmp_path_ref, dest).map_err(AtomicWriteError::Io)?;
    let _ = sync_parent_dir(dest);

    Ok(out)
}

/// Convenience helper for atomically writing a full byte slice to disk.
pub fn atomic_write_bytes(dest: impl AsRef<Path>, bytes: &[u8]) -> io::Result<()> {
    atomic_write(dest, |file| file.write_all(bytes)).map_err(|err| match err {
        AtomicWriteError::Io(err) => err,
        AtomicWriteError::Writer(err) => err,
    })
}

fn sync_parent_dir(path: &Path) -> io::Result<()> {
    let parent = parent_dir_or_dot(path);
    // On most Unix platforms, opening a directory as a file is supported.
    // On others (or on Windows), this may fail; callers treat it as best-effort.
    let dir = File::open(parent)?;
    dir.sync_all()
}

fn replace_file(from: &Path, to: &Path) -> io::Result<()> {
    #[cfg(windows)]
    {
        use std::os::windows::ffi::OsStrExt as _;
        use windows_sys::Win32::Storage::FileSystem::{MoveFileExW, MOVEFILE_REPLACE_EXISTING};

        fn to_wide_null(path: &Path) -> Vec<u16> {
            let mut wide: Vec<u16> = path.as_os_str().encode_wide().collect();
            wide.push(0);
            wide
        }

        let from_w = to_wide_null(from);
        let to_w = to_wide_null(to);
        let flags = MOVEFILE_REPLACE_EXISTING;
        let ok = unsafe { MoveFileExW(from_w.as_ptr(), to_w.as_ptr(), flags) };
        if ok == 0 {
            return Err(io::Error::last_os_error());
        }
        Ok(())
    }

    #[cfg(not(windows))]
    {
        fs::rename(from, to)
    }
}

/// Generate a sibling file path in the same directory.
///
/// This is mostly useful for tests that want a deterministic-but-unique temp path.
pub fn sibling_path_with_suffix(path: impl AsRef<Path>, suffix: &str) -> PathBuf {
    let path = path.as_ref();
    let dir = parent_dir_or_dot(path);
    let file_name = path.file_name().unwrap_or_default();
    dir.join(format!("{}{}", file_name.to_string_lossy(), suffix))
}

#[cfg(test)]
mod tests {
    use super::*;
    use std::io;

    static CWD_LOCK: std::sync::Mutex<()> = std::sync::Mutex::new(());

    struct CwdGuard {
        old: std::path::PathBuf,
    }

    impl CwdGuard {
        fn chdir(path: &Path) -> Self {
            let old = std::env::current_dir().expect("current_dir");
            std::env::set_current_dir(path).expect("set_current_dir");
            Self { old }
        }
    }

    impl Drop for CwdGuard {
        fn drop(&mut self) {
            let _ = std::env::set_current_dir(&self.old);
        }
    }

    #[test]
    fn atomic_write_supports_relative_dest_in_current_directory() {
        let _guard = CWD_LOCK.lock().expect("lock");

        let tmp = tempfile::tempdir().expect("temp dir");
        let _cwd = CwdGuard::chdir(tmp.path());

        // `file.bin` has an empty `Path::parent()`; this should still work.
        atomic_write_bytes("file.bin", b"hello").expect("atomic write");
        assert_eq!(
            std::fs::read(tmp.path().join("file.bin")).expect("read file"),
            b"hello"
        );
    }

    #[test]
    fn atomic_write_with_path_does_not_clobber_existing_file_on_write_error() {
        let tmp = tempfile::tempdir().expect("temp dir");
        let dest = tmp.path().join("existing.bin");

        let sentinel = b"sentinel-bytes";
        std::fs::write(&dest, sentinel).expect("write sentinel dest file");

        let err = atomic_write_with_path(&dest, |tmp_path| {
            std::fs::write(tmp_path, b"partial").expect("write to temp file");
            Err::<(), _>(io::Error::new(io::ErrorKind::Other, "simulated write failure"))
        })
        .expect_err("expected atomic_write_with_path to return error");

        // The destination file must remain untouched.
        let got = std::fs::read(&dest).expect("read dest");
        assert_eq!(got, sentinel, "dest file should not be clobbered: {err}");

        // Temp file should be cleaned up.
        let entries: Vec<_> = std::fs::read_dir(tmp.path())
            .expect("read_dir")
            .collect::<Result<Vec<_>, _>>()
            .expect("list dir");
        let names: Vec<_> = entries
            .iter()
            .map(|e| e.path())
            .filter(|p| p.is_file())
            .collect();
        assert_eq!(
            names,
            vec![dest.clone()],
            "expected only the destination file to remain (no temp files)"
        );
    }
}
