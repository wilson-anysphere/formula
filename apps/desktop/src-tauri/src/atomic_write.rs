use anyhow::Context;
use std::io::Write;
use std::path::Path;
use tempfile::NamedTempFile;

/// Atomically writes `bytes` to `path`.
///
/// This creates any missing parent directories, writes to a temp file in the
/// destination directory, then renames into place. Callers should prefer this
/// over `std::fs::write` for user-visible save paths so we never leave partially
/// written files behind if the process crashes mid-save.
pub fn write_file_atomic(path: &Path, bytes: &[u8]) -> anyhow::Result<()> {
    // `Path::parent` returns `None` for paths like `foo.xlsx` (no separators).
    // In that case, treat the current directory as the parent so the temp file
    // is still created alongside the destination.
    let parent = path.parent().unwrap_or_else(|| Path::new("."));
    std::fs::create_dir_all(parent).with_context(|| format!("create parent directory {parent:?}"))?;

    let mut tmp = NamedTempFile::new_in(parent)
        .with_context(|| format!("create temp file in {parent:?} for {path:?}"))?;
    tmp.as_file_mut()
        .write_all(bytes)
        .with_context(|| format!("write temp file for {path:?}"))?;

    // Best-effort durability: ensure bytes have hit the OS before the rename.
    tmp.as_file()
        .sync_all()
        .with_context(|| format!("sync temp file for {path:?}"))?;

    tmp.persist(path)
        .map(|_| ())
        .with_context(|| format!("persist temp file to {path:?}"))?;

    Ok(())
}

