use anyhow::{anyhow, Result};
#[cfg(any(feature = "desktop", test))]
use anyhow::Context;
use directories::{BaseDirs, UserDirs};
#[cfg(any(feature = "desktop", test))]
use std::ffi::OsStr;
use std::path::{Path, PathBuf};

pub(crate) const ACCESS_DENIED_SCOPE: &str =
    "Access denied: path is outside the allowed filesystem scope";

/// Return the canonicalized filesystem roots that the desktop app is allowed to access.
///
/// This mirrors the desktop filesystem scope policy used for local file access:
/// - `$HOME/**`
/// - `$DOCUMENT/**`
/// - `$DOWNLOADS/**` (if the OS/user has a Downloads dir configured and it exists/canonicalizes successfully)
pub(crate) fn desktop_allowed_roots() -> Result<Vec<PathBuf>> {
    let base_dirs =
        BaseDirs::new().ok_or_else(|| anyhow!("unable to determine home directory"))?;
    let mut roots = vec![base_dirs.home_dir().to_path_buf()];

    if let Some(user_dirs) = UserDirs::new() {
        if let Some(documents) = user_dirs.document_dir() {
            roots.push(documents.to_path_buf());
        }
        if let Some(downloads) = user_dirs.download_dir() {
            roots.push(downloads.to_path_buf());
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
    #[error("failed to canonicalize path: {source}")]
    Canonicalize {
        path: PathBuf,
        #[source]
        source: std::io::Error,
    },
    #[error("{ACCESS_DENIED_SCOPE}")]
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

/// Resolve a prospective save destination and verify it is contained within `allowed_roots`.
///
/// Unlike [`canonicalize_in_allowed_roots_with_error`], this helper supports paths that do not
/// exist yet by canonicalizing only the parent directory and re-attaching the requested file name.
///
/// This is intended for "save as" style operations where the destination file may not exist.
#[cfg(any(feature = "desktop", test))]
pub(crate) fn resolve_save_path_in_allowed_roots(
    path: &Path,
    allowed_roots: &[PathBuf],
) -> Result<PathBuf> {
    let Some(file_name) = path.file_name() else {
        return Err(anyhow!("Invalid save path: path must include a file name"));
    };
    if file_name == OsStr::new(".") || file_name == OsStr::new("..") {
        return Err(anyhow!("Invalid save path: file name cannot be '.' or '..'"));
    }

    // `Path::parent()` returns an empty path for single-component relative paths ("foo.xlsx").
    // Canonicalize treats this inconsistently across platforms, so normalize to ".".
    let parent = match path.parent() {
        Some(parent) if !parent.as_os_str().is_empty() => parent,
        _ => Path::new("."),
    };

    let canonical_parent = match canonicalize_in_allowed_roots_with_error(parent, allowed_roots) {
        Ok(parent) => parent,
        Err(CanonicalizeInAllowedRootsError::OutsideScope { .. }) => return Err(anyhow!(ACCESS_DENIED_SCOPE)),
        Err(CanonicalizeInAllowedRootsError::Canonicalize { source, .. }) => {
            return Err(anyhow::Error::new(source)).context("canonicalize path");
        }
    };

    let candidate = canonical_parent.join(Path::new(file_name));

    // If the destination path already exists, ensure it doesn't resolve (via symlink) outside the
    // scope.
    //
    // We intentionally treat dangling symlinks as existing entries (via `symlink_metadata`) since
    // writing through a dangling symlink could create a new file outside the allowed scope.
    if let Ok(meta) = std::fs::symlink_metadata(&candidate) {
        match canonicalize_in_allowed_roots_with_error(&candidate, allowed_roots) {
            Ok(_) => {}
            Err(CanonicalizeInAllowedRootsError::OutsideScope { .. }) => {
                return Err(anyhow!(ACCESS_DENIED_SCOPE));
            }
            Err(CanonicalizeInAllowedRootsError::Canonicalize { source, .. }) => {
                if meta.file_type().is_symlink() {
                    return Err(anyhow!(ACCESS_DENIED_SCOPE));
                }
                return Err(anyhow::Error::new(source)).context("canonicalize path");
            }
        }
    }

    Ok(candidate)
}

/// Canonicalize `path` and verify it is contained within `allowed_roots`.
///
/// This is used by IPC commands that proxy filesystem access to the webview. Canonicalization
/// normalizes `..` segments and resolves symlinks, preventing symlink-based scope escapes.
#[cfg(any(feature = "desktop", test))]
pub(crate) fn canonicalize_in_allowed_roots(path: &Path, allowed_roots: &[PathBuf]) -> Result<PathBuf> {
    match canonicalize_in_allowed_roots_with_error(path, allowed_roots) {
        Ok(path) => Ok(path),
        Err(CanonicalizeInAllowedRootsError::OutsideScope { .. }) => Err(anyhow!(ACCESS_DENIED_SCOPE)),
        Err(CanonicalizeInAllowedRootsError::Canonicalize { source, .. }) => {
            Err(anyhow::Error::new(source)).context("canonicalize path")
        }
    }
}

#[cfg(test)]
mod tests {
    use super::*;
    use std::fs;
    use std::sync::{Mutex, OnceLock};

    static ENV_MUTEX: OnceLock<Mutex<()>> = OnceLock::new();

    fn env_mutex() -> &'static Mutex<()> {
        ENV_MUTEX.get_or_init(|| Mutex::new(()))
    }

    #[cfg(target_os = "linux")]
    use std::ffi::OsString;

    #[cfg(target_os = "linux")]
    struct EnvVarGuard {
        key: &'static str,
        prev: Option<OsString>,
    }

    #[cfg(target_os = "linux")]
    impl EnvVarGuard {
        fn set(key: &'static str, value: impl AsRef<std::ffi::OsStr>) -> Self {
            let prev = std::env::var_os(key);
            std::env::set_var(key, value);
            Self { key, prev }
        }
    }

    #[cfg(target_os = "linux")]
    impl Drop for EnvVarGuard {
        fn drop(&mut self) {
            match &self.prev {
                Some(prev) => std::env::set_var(self.key, prev),
                None => std::env::remove_var(self.key),
            }
        }
    }

    fn tempdir_outside_allowed_roots(allowed_roots: &[PathBuf]) -> tempfile::TempDir {
        fn tempdir_if_outside(
            base: Option<&Path>,
            allowed_roots: &[PathBuf],
        ) -> Option<tempfile::TempDir> {
            let tmp = match base {
                Some(base) => tempfile::tempdir_in(base).ok()?,
                None => tempfile::tempdir().ok()?,
            };
            let canon = dunce::canonicalize(tmp.path()).ok()?;
            if path_in_allowed_roots(&canon, allowed_roots) {
                return None;
            }
            Some(tmp)
        }

        // 1) Default temp dir (fast path on Unix where it is typically outside `$HOME`).
        if let Some(tmp) = tempdir_if_outside(None, allowed_roots) {
            return tmp;
        }

        // 2) Try some OS-specific global temp locations.
        #[cfg(unix)]
        {
            // Some environments set `TMPDIR` under `$HOME`, which would fall inside the default
            // allowed roots. Fall back to well-known global temp directories.
            for base in [Path::new("/tmp"), Path::new("/var/tmp")] {
                if let Some(tmp) = tempdir_if_outside(Some(base), allowed_roots) {
                    return tmp;
                }
            }
        }

        #[cfg(windows)]
        {
            // `%WINDIR%\\Temp` (commonly writable for normal users).
            if let Some(windir) = std::env::var_os("WINDIR").or_else(|| std::env::var_os("SystemRoot"))
            {
                let base = PathBuf::from(windir).join("Temp");
                if let Some(tmp) = tempdir_if_outside(Some(&base), allowed_roots) {
                    return tmp;
                }
            }

            // `%ProgramData%` is typically outside the user profile.
            if let Some(program_data) = std::env::var_os("ProgramData") {
                let base = PathBuf::from(program_data);
                if let Some(tmp) = tempdir_if_outside(Some(&base), allowed_roots) {
                    return tmp;
                }
            }

            // `C:\\Users\\Public` is commonly present and outside `$HOME`.
            if let Some(public) = std::env::var_os("PUBLIC") {
                let base = PathBuf::from(public);
                if let Some(tmp) = tempdir_if_outside(Some(&base), allowed_roots) {
                    return tmp;
                }
            }
        }

        panic!("unable to create a temporary directory outside the allowed filesystem scope");
    }

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

    #[cfg(unix)]
    #[test]
    fn resolve_save_path_in_allowed_roots_rejects_dangling_symlink_escape() {
        use std::os::unix::fs as unix_fs;

        let tmp = tempfile::tempdir().expect("tempdir");
        let allowed_root = tmp.path().join("root");
        let outside_root = tmp.path().join("outside");
        fs::create_dir_all(&allowed_root).expect("create root");
        fs::create_dir_all(&outside_root).expect("create outside root");

        // A dangling symlink inside the allowed root pointing outside. Writing to this path would
        // otherwise create a new file outside the scope.
        let outside_target = outside_root.join("new.xlsx");
        let symlink_path = allowed_root.join("dangling.xlsx");
        unix_fs::symlink(&outside_target, &symlink_path).expect("symlink");

        let allowed_roots = vec![dunce::canonicalize(&allowed_root).expect("canonicalize root")];
        let err = resolve_save_path_in_allowed_roots(&symlink_path, &allowed_roots).unwrap_err();
        assert!(err.to_string().to_lowercase().contains("outside"));
    }

    #[test]
    fn resolve_save_path_in_allowed_roots_rejects_missing_file_name() {
        let tmp = tempfile::tempdir().expect("tempdir");
        let allowed_roots = vec![dunce::canonicalize(tmp.path()).expect("canonicalize root")];

        let err = resolve_save_path_in_allowed_roots(Path::new(""), &allowed_roots).unwrap_err();
        assert!(err.to_string().to_lowercase().contains("file name"));
    }

    #[test]
    fn resolve_save_path_in_allowed_roots_rejects_dot_and_dotdot_file_names() {
        let tmp = tempfile::tempdir().expect("tempdir");
        let allowed_root = tmp.path().join("root");
        fs::create_dir_all(&allowed_root).expect("create root");
        let allowed_roots = vec![dunce::canonicalize(&allowed_root).expect("canonicalize root")];

        let err = resolve_save_path_in_allowed_roots(&allowed_root.join("."), &allowed_roots).unwrap_err();
        let msg = err.to_string().to_lowercase();
        assert!(
            msg.contains("file name") || msg.contains("outside"),
            "unexpected error for dot path: {err}"
        );

        let err =
            resolve_save_path_in_allowed_roots(&allowed_root.join(".."), &allowed_roots).unwrap_err();
        let msg = err.to_string().to_lowercase();
        assert!(
            msg.contains("file name") || msg.contains("outside"),
            "unexpected error for dotdot path: {err}"
        );
    }

    #[test]
    fn desktop_scope_open_validation_allows_home_file_and_rejects_out_of_scope() {
        let _guard = env_mutex()
            .lock()
            .unwrap_or_else(|poisoned| poisoned.into_inner());

        let base_dirs = BaseDirs::new().expect("base dirs");
        let in_scope_tmp = tempfile::tempdir_in(base_dirs.home_dir()).expect("tempdir in home");
        let in_scope_file = in_scope_tmp.path().join("in-scope.xlsx");
        fs::write(&in_scope_file, "hello").expect("write in-scope file");

        let allowed_roots = desktop_allowed_roots().expect("allowed roots");
        let out_scope_tmp = tempdir_outside_allowed_roots(&allowed_roots);
        let out_scope_file = out_scope_tmp.path().join("out-of-scope.xlsx");
        fs::write(&out_scope_file, "secret").expect("write out-of-scope file");

        let resolved_in_scope =
            canonicalize_in_allowed_roots(&in_scope_file, &allowed_roots).expect("in scope");
        assert!(path_in_allowed_roots(&resolved_in_scope, &allowed_roots));

        let err = canonicalize_in_allowed_roots(&out_scope_file, &allowed_roots).unwrap_err();
        assert!(err.to_string().to_lowercase().contains("outside"));
    }

    #[cfg(target_os = "linux")]
    #[test]
    fn desktop_scope_includes_xdg_downloads_dir_outside_home() {
        let _guard = env_mutex()
            .lock()
            .unwrap_or_else(|poisoned| poisoned.into_inner());

        let config_dir = tempfile::tempdir().expect("config tempdir");
        let downloads_dir = tempfile::tempdir().expect("downloads tempdir");
        fs::write(
            config_dir.path().join("user-dirs.dirs"),
            format!("XDG_DOWNLOAD_DIR=\"{}\"\n", downloads_dir.path().display()),
        )
        .expect("write user-dirs.dirs");

        let _xdg_config_home = EnvVarGuard::set("XDG_CONFIG_HOME", config_dir.path());

        // Sanity: this test configures Downloads to resolve outside `$HOME`.
        let base_dirs = BaseDirs::new().expect("base dirs");
        let home_canon = dunce::canonicalize(base_dirs.home_dir()).expect("canonicalize home");
        let downloads_canon = dunce::canonicalize(downloads_dir.path()).expect("canonicalize downloads");
        assert!(
            !downloads_canon.starts_with(&home_canon),
            "test setup error: downloads dir should be outside home"
        );

        let allowed_roots = desktop_allowed_roots().expect("allowed roots");
        assert!(
            allowed_roots.iter().any(|root| root == &downloads_canon),
            "expected desktop_allowed_roots to include downloads root"
        );

        let file_path = downloads_dir.path().join("from_downloads.txt");
        fs::write(&file_path, "hello").expect("write file");

        let resolved =
            canonicalize_in_allowed_roots(&file_path, &allowed_roots).expect("downloads allowed");
        assert!(resolved.starts_with(&downloads_canon));
    }

    #[test]
    fn desktop_scope_save_validation_allows_nonexistent_home_path_and_rejects_out_of_scope() {
        let _guard = env_mutex()
            .lock()
            .unwrap_or_else(|poisoned| poisoned.into_inner());

        let base_dirs = BaseDirs::new().expect("base dirs");
        let in_scope_tmp = tempfile::tempdir_in(base_dirs.home_dir()).expect("tempdir in home");
        let in_scope_dest = in_scope_tmp.path().join("new-workbook.xlsx");
        assert!(!in_scope_dest.exists());

        let allowed_roots = desktop_allowed_roots().expect("allowed roots");
        let out_scope_tmp = tempdir_outside_allowed_roots(&allowed_roots);
        let out_scope_dest = out_scope_tmp.path().join("new-workbook.xlsx");
        assert!(!out_scope_dest.exists());

        let resolved_in_scope =
            resolve_save_path_in_allowed_roots(&in_scope_dest, &allowed_roots).expect("in scope");
        let expected_in_scope_parent = dunce::canonicalize(in_scope_tmp.path()).expect("canon");
        assert_eq!(
            resolved_in_scope,
            expected_in_scope_parent.join("new-workbook.xlsx")
        );

        let err = resolve_save_path_in_allowed_roots(&out_scope_dest, &allowed_roots).unwrap_err();
        assert!(err.to_string().to_lowercase().contains("outside"));
    }
}
