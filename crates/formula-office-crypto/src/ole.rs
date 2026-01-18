use std::collections::BTreeSet;
use std::io::{Read, Seek, SeekFrom};
use std::path::{Path, PathBuf};

use crate::error::OfficeCryptoError;
use crate::{
    MAX_OLE_PRESERVED_ENTRIES, MAX_OLE_PRESERVED_STREAM_BYTES, MAX_OLE_PRESERVED_TOTAL_BYTES,
};

/// A preserved OLE stream (path + bytes).
#[derive(Debug, Clone, PartialEq, Eq)]
pub struct OleStream {
    pub path: PathBuf,
    pub bytes: Vec<u8>,
}

/// A preserved OLE entry.
#[derive(Debug, Clone, PartialEq, Eq)]
pub enum OleEntry {
    Storage { path: PathBuf },
    Stream(OleStream),
}

/// A collection of OLE entries (storages + streams) suitable for round-trip preservation.
///
/// Note: This structure is intended for copying entries from an existing OLE container into a new
/// one. It intentionally does **not** try to preserve the internal CFB directory entry layout
/// byte-for-byte, only the entry *payloads*.
#[derive(Debug, Clone, Default, PartialEq, Eq)]
pub struct OleEntries {
    pub storages: Vec<PathBuf>,
    pub streams: Vec<OleStream>,
}

fn is_reserved_encryption_stream(path: &Path) -> bool {
    let s = path.to_string_lossy();
    let s = s.strip_prefix('/').unwrap_or(&s);
    s.eq_ignore_ascii_case("EncryptionInfo") || s.eq_ignore_ascii_case("EncryptedPackage")
}

fn strip_leading_slash(path: &Path) -> &Path {
    path.strip_prefix("/").unwrap_or(path)
}

fn is_root_path(path: &Path) -> bool {
    path == Path::new("/") || path.as_os_str().is_empty()
}

/// Enumerate and extract all streams/storages from an open OLE/CFB container.
///
/// This helper is intended for round-trip preservation of non-package streams/metadata (for
/// Office-encrypted workbooks, these are any entries besides `EncryptionInfo` and
/// `EncryptedPackage`).
///
/// The returned [`OleEntries`] **excludes** the `EncryptionInfo` and `EncryptedPackage` streams to
/// avoid duplicating large payloads in memory and because those streams must be regenerated on
/// re-encryption.
pub fn extract_ole_entries<R: Read + Seek>(
    ole: &mut cfb::CompoundFile<R>,
) -> Result<OleEntries, OfficeCryptoError> {
    // Collect paths first to avoid borrow conflicts between `walk()` and `open_stream()`.
    let mut storages: Vec<PathBuf> = Vec::new();
    let mut stream_paths: Vec<PathBuf> = Vec::new();

    let mut entry_count = 0usize;
    for entry in ole.walk() {
        entry_count = entry_count.saturating_add(1);
        if entry_count > MAX_OLE_PRESERVED_ENTRIES {
            return Err(OfficeCryptoError::InvalidFormat(format!(
                "too many OLE entries: {entry_count} exceeds limit {MAX_OLE_PRESERVED_ENTRIES}"
            )));
        }

        let path = entry.path().to_path_buf();
        if entry.is_storage() {
            // Skip the root entry.
            if !is_root_path(&path) && !is_reserved_encryption_stream(&path) {
                storages.push(path);
            }
        } else if entry.is_stream() {
            if !is_reserved_encryption_stream(&path) {
                stream_paths.push(path);
            }
        }
    }

    let mut streams: Vec<OleStream> = Vec::new();
    let _ = streams.try_reserve(stream_paths.len());
    let mut total_bytes: usize = 0;
    for path in stream_paths {
        let mut stream = ole.open_stream(&path)?;
        let len_u64 = stream.seek(SeekFrom::End(0))?;
        stream.seek(SeekFrom::Start(0))?;
        let len = usize::try_from(len_u64).map_err(|_| {
            OfficeCryptoError::InvalidFormat("OLE stream size overflow".to_string())
        })?;
        if len > MAX_OLE_PRESERVED_STREAM_BYTES {
            return Err(OfficeCryptoError::SizeLimitExceeded {
                context: "OLE preserved stream",
                limit: MAX_OLE_PRESERVED_STREAM_BYTES,
            });
        }
        let next_total = total_bytes.saturating_add(len);
        if next_total > MAX_OLE_PRESERVED_TOTAL_BYTES {
            return Err(OfficeCryptoError::SizeLimitExceeded {
                context: "OLE preserved streams total",
                limit: MAX_OLE_PRESERVED_TOTAL_BYTES,
            });
        }

        let mut bytes = Vec::new();
        bytes.resize(len, 0);
        stream.read_exact(&mut bytes)?;
        total_bytes = next_total;
        streams.push(OleStream { path, bytes });
    }

    Ok(OleEntries { storages, streams })
}

/// Copy preserved OLE entries into the destination container.
///
/// This is used when re-wrapping an encrypted OOXML package so that non-package metadata streams
/// (e.g. `\u{0005}SummaryInformation`) survive round-tripping.
///
/// Reserved streams (`EncryptionInfo`, `EncryptedPackage`) are always skipped.
pub(crate) fn copy_entries_into_ole<R: Read + std::io::Write + Seek>(
    ole: &mut cfb::CompoundFile<R>,
    entries: &OleEntries,
) -> Result<(), OfficeCryptoError> {
    use std::io::Write as _;

    // Create all storages first (including those implied by stream paths).
    let mut storage_paths: BTreeSet<PathBuf> = BTreeSet::new();
    for p in &entries.storages {
        if is_root_path(p) || is_reserved_encryption_stream(p) {
            continue;
        }
        let p = strip_leading_slash(p);
        if p.as_os_str().is_empty() {
            continue;
        }
        storage_paths.insert(p.to_path_buf());
    }
    for s in &entries.streams {
        if is_reserved_encryption_stream(&s.path) {
            continue;
        }
        // Add all parent storages to preserve structure.
        let mut cur = strip_leading_slash(&s.path);
        while let Some(parent) = cur.parent() {
            if parent.as_os_str().is_empty() {
                break;
            }
            storage_paths.insert(parent.to_path_buf());
            cur = parent;
        }
    }

    // Ensure parents are created before children by sorting by depth (number of components).
    let mut storage_paths: Vec<PathBuf> = storage_paths.into_iter().collect();
    storage_paths.sort_by_key(|p| p.components().count());
    for storage in storage_paths {
        // Some producers include storages that might already exist; treat that as non-fatal.
        match ole.create_storage(&storage) {
            Ok(_) => {}
            Err(err) if err.kind() == std::io::ErrorKind::AlreadyExists => {}
            Err(err) => return Err(err.into()),
        }
    }

    for s in &entries.streams {
        if is_reserved_encryption_stream(&s.path) {
            continue;
        }
        let path = strip_leading_slash(&s.path);
        if path.as_os_str().is_empty() {
            continue;
        }

        ole.create_stream(path)?.write_all(&s.bytes)?;
    }

    Ok(())
}
