use anyhow::{anyhow, Result};
#[cfg(any(feature = "desktop", test))]
use anyhow::Context;
use directories::{BaseDirs, UserDirs};
use std::path::{Path, PathBuf};

/// Return the canonicalized filesystem roots that the desktop app is allowed to access.
///
/// This mirrors the desktop filesystem scope policy used for local file access:
/// - `$HOME/**`
/// - `$DOCUMENT/**`
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
        // `dunce::canonicalize` avoids Windows `\\?\` prefix issues and behaves consistently across
        // platforms for scope comparisons.
        if let Ok(canon) = dunce::canonicalize(&root) {
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

#[derive(Debug, thiserror::Error)]
pub(crate) enum CanonicalizeInAllowedRootsError {
    #[error("failed to canonicalize '{path}'")]
    Canonicalize {
        path: PathBuf,
        #[source]
        source: std::io::Error,
    },
    #[error("path '{path}' is outside the allowed filesystem scope")]
    OutsideScope { path: PathBuf },
}

/// Canonicalize `path` and verify it is contained within `allowed_roots`.
///
/// This variant returns a structured error so callers can distinguish canonicalization failures
/// (e.g. missing files) from out-of-scope denials.
pub(crate) fn canonicalize_in_allowed_roots_with_error(
    path: &Path,
    allowed_roots: &[PathBuf],
) -> std::result::Result<PathBuf, CanonicalizeInAllowedRootsError> {
    let canonical =
        dunce::canonicalize(path).map_err(|e| CanonicalizeInAllowedRootsError::Canonicalize {
            path: path.to_path_buf(),
            source: e,
        })?;
    if path_in_allowed_roots(&canonical, allowed_roots) {
        Ok(canonical)
    } else {
        Err(CanonicalizeInAllowedRootsError::OutsideScope { path: canonical })
    }
}

/// Canonicalize `path` and verify it is contained within `allowed_roots`.
///
/// This is used by IPC commands that proxy filesystem access to the webview. Canonicalization
/// normalizes `..` segments and resolves symlinks, preventing symlink-based scope escapes.
#[cfg(any(feature = "desktop", test))]
pub(crate) fn canonicalize_in_allowed_roots(path: &Path, allowed_roots: &[PathBuf]) -> Result<PathBuf> {
    match canonicalize_in_allowed_roots_with_error(path, allowed_roots) {
        Ok(path) => Ok(path),
        Err(CanonicalizeInAllowedRootsError::OutsideScope { path }) => Err(anyhow!(
            "Refusing to access '{}' because it is outside the allowed filesystem scope",
            path.display()
        )),
        Err(CanonicalizeInAllowedRootsError::Canonicalize { path, source }) => {
            Err(anyhow::Error::new(source)).context(format!("canonicalize {}", path.display()))
        }
    }
}

#[cfg(test)]
mod tests {
    use super::*;
    use std::fs;

    #[test]
    fn canonicalize_in_allowed_roots_allows_paths_under_root() {
        let tmp = tempfile::tempdir().expect("tempdir");
        let allowed_root = tmp.path().join("root");
        fs::create_dir_all(&allowed_root).expect("create root");
        let file_path = allowed_root.join("hello.txt");
        fs::write(&file_path, "hello").expect("write file");

        let allowed_roots = vec![dunce::canonicalize(&allowed_root).expect("canonicalize root")];
        let resolved = canonicalize_in_allowed_roots(&file_path, &allowed_roots).expect("in scope");
        assert!(resolved.is_absolute());
    }

    #[test]
    fn canonicalize_in_allowed_roots_rejects_paths_outside_root() {
        let tmp = tempfile::tempdir().expect("tempdir");
        let allowed_root = tmp.path().join("root");
        let outside_root = tmp.path().join("outside");
        fs::create_dir_all(&allowed_root).expect("create root");
        fs::create_dir_all(&outside_root).expect("create outside root");
        let file_path = outside_root.join("secret.txt");
        fs::write(&file_path, "secret").expect("write file");

        let allowed_roots = vec![dunce::canonicalize(&allowed_root).expect("canonicalize root")];
        let err = canonicalize_in_allowed_roots(&file_path, &allowed_roots).unwrap_err();
        assert!(err.to_string().to_lowercase().contains("outside"));
    }

    #[cfg(unix)]
    #[test]
    fn canonicalize_in_allowed_roots_blocks_symlink_escape() {
        use std::os::unix::fs as unix_fs;

        let tmp = tempfile::tempdir().expect("tempdir");
        let allowed_root = tmp.path().join("root");
        let outside_root = tmp.path().join("outside");
        fs::create_dir_all(&allowed_root).expect("create root");
        fs::create_dir_all(&outside_root).expect("create outside root");
        let outside_file = outside_root.join("secret.txt");
        fs::write(&outside_file, "secret").expect("write outside file");

        let symlink_path = allowed_root.join("escape.txt");
        unix_fs::symlink(&outside_file, &symlink_path).expect("symlink");

        let allowed_roots = vec![dunce::canonicalize(&allowed_root).expect("canonicalize root")];
        let err = canonicalize_in_allowed_roots(&symlink_path, &allowed_roots).unwrap_err();
        assert!(err.to_string().to_lowercase().contains("outside"));
    }
}
