use std::io;
use std::path::{Path, PathBuf};

use sha2::{Digest, Sha256};

/// A small helper for persisting downloaded updater payloads to disk so we don't
/// keep large update binaries resident in RAM while waiting for user approval.
///
/// This module intentionally has **no Tauri dependencies** so it can be unit
/// tested in headless builds.
#[derive(Clone, Debug)]
pub struct UpdaterDownloadCache {
    dir: PathBuf,
}

impl UpdaterDownloadCache {
    /// Create a cache rooted under `base_dir`.
    ///
    /// `base_dir` should be an app-owned cache directory when available (e.g.
    /// Tauri's `app_cache_dir()`), but tests can pass any temp directory.
    pub fn new(base_dir: impl Into<PathBuf>) -> Self {
        Self {
            dir: base_dir.into().join("updater").join("download-cache"),
        }
    }

    pub fn dir(&self) -> &Path {
        &self.dir
    }

    /// Compute the cache file path for a given update `version`.
    ///
    /// The filename is derived from a hash of the version string to avoid path
    /// traversal issues and to keep filenames filesystem-safe.
    pub fn path_for_version(&self, version: &str) -> PathBuf {
        self.dir.join(file_name_for_version(version))
    }

    /// Atomically write `bytes` to the cache and return the final path.
    pub fn write(&self, version: &str, bytes: &[u8]) -> io::Result<PathBuf> {
        let dest = self.path_for_version(version);
        formula_fs::atomic_write_bytes(&dest, bytes)?;
        Ok(dest)
    }

    /// Read a cached payload from `path`.
    pub fn read(path: &Path) -> io::Result<Vec<u8>> {
        std::fs::read(path)
    }

    /// Delete a cached payload at `path`.
    ///
    /// Deleting a missing file is treated as success.
    pub fn delete(path: &Path) -> io::Result<()> {
        match std::fs::remove_file(path) {
            Ok(()) => Ok(()),
            Err(err) if err.kind() == io::ErrorKind::NotFound => Ok(()),
            Err(err) => Err(err),
        }
    }
}

fn file_name_for_version(version: &str) -> String {
    // Versioned scheme so we can change the naming strategy in the future
    // without clobbering unrelated cache files.
    const PREFIX: &[u8] = b"formula-updater-download-cache-v1\0";
    let mut hasher = Sha256::new();
    hasher.update(PREFIX);
    hasher.update(version.as_bytes());
    let hash = hex::encode(hasher.finalize());
    format!("payload-{hash}.bin")
}

#[cfg(test)]
mod tests {
    use super::*;

    fn list_files_recursive(dir: &Path) -> Vec<PathBuf> {
        fn visit(dir: &Path, out: &mut Vec<PathBuf>) {
            let entries = match std::fs::read_dir(dir) {
                Ok(entries) => entries,
                Err(err) if err.kind() == io::ErrorKind::NotFound => return,
                Err(err) => panic!("read_dir({dir:?}) failed: {err}"),
            };
            for entry in entries {
                let entry = entry.expect("read_dir entry");
                let path = entry.path();
                let meta = entry.metadata().expect("entry metadata");
                if meta.is_dir() {
                    visit(&path, out);
                } else if meta.is_file() {
                    out.push(path);
                }
            }
        }

        let mut out = Vec::new();
        visit(dir, &mut out);
        out.sort();
        out
    }

    #[test]
    fn write_and_read_roundtrip() {
        let tmp = tempfile::tempdir().expect("tempdir");
        let cache = UpdaterDownloadCache::new(tmp.path());

        let bytes = b"hello update payload";
        let path = cache.write("1.2.3", bytes).expect("write");
        let got = UpdaterDownloadCache::read(&path).expect("read");

        assert_eq!(got, bytes);
    }

    #[test]
    fn delete_cleanup_works() {
        let tmp = tempfile::tempdir().expect("tempdir");
        let cache = UpdaterDownloadCache::new(tmp.path());

        let path = cache.write("1.2.3", b"payload").expect("write");
        assert!(path.is_file(), "expected payload file to exist");

        UpdaterDownloadCache::delete(&path).expect("delete");
        assert!(!path.exists(), "expected payload file to be deleted");

        // Deleting again should be a no-op.
        UpdaterDownloadCache::delete(&path).expect("delete missing should succeed");
    }

    #[test]
    fn multiple_versions_use_distinct_paths() {
        let tmp = tempfile::tempdir().expect("tempdir");
        let cache = UpdaterDownloadCache::new(tmp.path());

        let path_a = cache.write("1.0.0", b"a").expect("write a");
        let path_b = cache.write("2.0.0", b"b").expect("write b");

        assert_ne!(path_a, path_b, "expected distinct cache paths per version");
        assert_eq!(UpdaterDownloadCache::read(&path_a).expect("read a"), b"a");
        assert_eq!(UpdaterDownloadCache::read(&path_b).expect("read b"), b"b");
    }

    #[test]
    fn does_not_leave_temp_files_behind() {
        let tmp = tempfile::tempdir().expect("tempdir");
        let cache = UpdaterDownloadCache::new(tmp.path());

        let path = cache.write("1.2.3", b"payload").expect("write");

        // The only file on disk should be the final payload file (no leftover temp files).
        let files = list_files_recursive(tmp.path());
        assert_eq!(files, vec![path.clone()]);

        UpdaterDownloadCache::delete(&path).expect("delete");

        let files_after = list_files_recursive(tmp.path());
        assert!(files_after.is_empty(), "expected no files after delete");
    }
}

