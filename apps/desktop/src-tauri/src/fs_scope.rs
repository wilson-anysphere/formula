use anyhow::{anyhow, Result};
use directories::{BaseDirs, UserDirs};
use std::path::{Path, PathBuf};

/// Return the canonicalized filesystem roots that the desktop app is allowed to access.
///
/// This mirrors the desktop filesystem scope policy (home + documents) used for other local
/// file-access commands.
pub(crate) fn desktop_allowed_roots() -> Result<Vec<PathBuf>> {
    let base_dirs =
        BaseDirs::new().ok_or_else(|| anyhow!("unable to determine home directory"))?;
    let mut roots = vec![base_dirs.home_dir().to_path_buf()];

    if let Some(user_dirs) = UserDirs::new() {
        if let Some(documents) = user_dirs.document_dir() {
            roots.push(documents.to_path_buf());
        }
    }

    let mut canonical_roots = Vec::new();
    for root in roots {
        if let Ok(canon) = std::fs::canonicalize(&root) {
            if !canonical_roots.iter().any(|existing| *existing == canon) {
                canonical_roots.push(canon);
            }
        }
    }

    Ok(canonical_roots)
}

pub(crate) fn path_in_allowed_roots(path: &Path, allowed_roots: &[PathBuf]) -> bool {
    allowed_roots.iter().any(|root| path.starts_with(root))
}

