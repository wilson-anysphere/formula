use std::ffi::OsStr;
use std::path::{Path, PathBuf};

/// Restricts filesystem access to a fixed set of allowed roots.
///
/// This is used by Tauri IPC commands that accept a user-controlled path from the
/// webview. These commands must not rely on the Tauri FS plugin's permission
/// scoping since they bypass it by calling `std::fs` directly.
#[derive(Clone, Debug)]
pub struct PathScopePolicy {
    roots: Vec<PathBuf>,
}

impl PathScopePolicy {
    const ACCESS_DENIED: &'static str = "Access denied: path is outside the allowed filesystem scope";

    /// Build a policy with an explicit set of allowed roots.
    ///
    /// This is primarily intended for tests; production code should prefer
    /// `default_desktop()`.
    pub fn new(roots: Vec<PathBuf>) -> Self {
        let mut roots: Vec<PathBuf> = roots
            .into_iter()
            .filter_map(|root| {
                // Roots are best-effort canonicalized. If canonicalization fails (e.g. the root
                // doesn't exist), fall back to the original path so we don't panic during
                // construction. Validation will still be strict (it canonicalizes the target).
                match dunce::canonicalize(&root) {
                    Ok(canonical) => Some(canonical),
                    Err(_) => Some(root),
                }
            })
            .collect();

        roots.sort();
        roots.dedup();

        Self { roots }
    }

    /// Default policy for the desktop app: allow `$HOME/**`, `$DOCUMENT/**`, and (optionally)
    /// `$DOWNLOADS/**`.
    pub fn default_desktop() -> Self {
        let mut roots = Vec::new();

        if let Some(user_dirs) = directories::UserDirs::new() {
            roots.push(user_dirs.home_dir().to_path_buf());
            if let Some(documents) = user_dirs.document_dir() {
                roots.push(documents.to_path_buf());
            }
            // Include Downloads only if it exists and can be canonicalized. On some systems
            // (notably Linux with XDG user dirs), Downloads may be configured outside `$HOME`.
            if let Some(downloads) = user_dirs.download_dir() {
                if let Ok(canonical) = dunce::canonicalize(downloads) {
                    roots.push(canonical);
                }
            }
        } else {
            // Fallbacks if `directories` cannot determine platform user dirs.
            // Keep the scope minimal: only a single home directory root.
            if let Some(home) = std::env::var_os("HOME") {
                roots.push(PathBuf::from(home));
            } else if let Some(profile) = std::env::var_os("USERPROFILE") {
                roots.push(PathBuf::from(profile));
            }
        }

        Self::new(roots)
    }

    /// Validate a path used for reading.
    ///
    /// The returned path is canonicalized (resolves `..` and symlinks).
    pub fn validate_read_path(&self, input: &Path) -> Result<PathBuf, String> {
        let canonical = dunce::canonicalize(input).map_err(|e| e.to_string())?;
        if !self.is_within_scope(&canonical) {
            return Err(Self::ACCESS_DENIED.to_string());
        }
        Ok(canonical)
    }

    /// Validate a path used for writing (the file may not exist yet).
    ///
    /// This canonicalizes the parent directory and then joins the provided file
    /// name. If the target file already exists, it is additionally canonicalized
    /// to prevent symlink escapes.
    pub fn validate_write_path(&self, input: &Path) -> Result<PathBuf, String> {
        let file_name = input
            .file_name()
            .ok_or_else(|| "Invalid path: expected a file path".to_string())?;
        if file_name == OsStr::new(".") || file_name == OsStr::new("..") {
            return Err("Invalid path: expected a file path".to_string());
        }

        let parent = input
            .parent()
            .filter(|parent| !parent.as_os_str().is_empty())
            .ok_or_else(|| "Invalid path: expected a file path with a parent directory".to_string())?;

        let canonical_parent = dunce::canonicalize(parent).map_err(|e| e.to_string())?;
        let candidate = canonical_parent.join(file_name);

        if !self.is_within_scope(&candidate) {
            return Err(Self::ACCESS_DENIED.to_string());
        }

        // If the file already exists, ensure it doesn't resolve (via symlink) outside the scope.
        if candidate.exists() {
            let canonical_target = dunce::canonicalize(&candidate).map_err(|e| e.to_string())?;
            if !self.is_within_scope(&canonical_target) {
                return Err(Self::ACCESS_DENIED.to_string());
            }
        }

        Ok(candidate)
    }

    fn is_within_scope(&self, canonical: &Path) -> bool {
        self.roots.iter().any(|root| canonical.starts_with(root))
    }
}

#[cfg(test)]
mod tests {
    use super::PathScopePolicy;
    use directories::BaseDirs;

    #[cfg(target_os = "linux")]
    use std::ffi::OsString;
    #[cfg(target_os = "linux")]
    use std::sync::{Mutex, OnceLock};

    #[cfg(target_os = "linux")]
    static ENV_MUTEX: OnceLock<Mutex<()>> = OnceLock::new();

    #[cfg(target_os = "linux")]
    fn env_mutex() -> &'static Mutex<()> {
        ENV_MUTEX.get_or_init(|| Mutex::new(()))
    }

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

    #[test]
    fn read_allows_paths_within_allowed_root() {
        let allowed_temp = tempfile::tempdir().unwrap();
        let allowed_path = allowed_temp.path().join("ok.txt");
        std::fs::write(&allowed_path, "hello").unwrap();

        let allowed_root = dunce::canonicalize(allowed_temp.path()).unwrap();
        let policy = PathScopePolicy::new(vec![allowed_root.clone()]);
        let validated = policy.validate_read_path(&allowed_path).unwrap();
        assert!(validated.starts_with(&allowed_root));
    }

    #[test]
    fn read_rejects_paths_outside_allowed_root() {
        let allowed_root = tempfile::tempdir().unwrap();
        let disallowed_root = tempfile::tempdir().unwrap();
        let disallowed_path = disallowed_root.path().join("secret.txt");
        std::fs::write(&disallowed_path, "secret").unwrap();

        let policy = PathScopePolicy::new(vec![allowed_root.path().to_path_buf()]);
        let err = policy.validate_read_path(&disallowed_path).unwrap_err();
        assert_eq!(
            err,
            "Access denied: path is outside the allowed filesystem scope"
        );
    }

    #[cfg(target_os = "linux")]
    #[test]
    fn default_desktop_allows_paths_in_downloads_dir_outside_home() {
        // `directories::UserDirs` relies on process-wide environment variables (HOME/XDG_*), so
        // guard against concurrent tests that might also manipulate them.
        let _guard = env_mutex().lock().unwrap();

        let tmp = tempfile::tempdir().unwrap();
        let home_dir = tmp.path().join("home");
        let downloads_dir = tmp.path().join("downloads");
        let config_dir = tmp.path().join("config");
        std::fs::create_dir_all(&home_dir).unwrap();
        std::fs::create_dir_all(&downloads_dir).unwrap();
        std::fs::create_dir_all(&config_dir).unwrap();

        // Configure Downloads to resolve outside `$HOME` via XDG user-dirs config.
        std::fs::write(
            config_dir.join("user-dirs.dirs"),
            "XDG_DOWNLOAD_DIR=\"$HOME/../downloads\"\n",
        )
        .unwrap();

        let _home = EnvVarGuard::set("HOME", &home_dir);
        let _xdg_config = EnvVarGuard::set("XDG_CONFIG_HOME", &config_dir);

        // Sanity: Downloads is not a subdirectory of HOME.
        assert!(
            !downloads_dir.starts_with(&home_dir),
            "test setup error: downloads_dir should be outside home_dir"
        );

        let file_path = downloads_dir.join("from_downloads.txt");
        std::fs::write(&file_path, "hello").unwrap();

        // PathScopePolicy (read + write path validation) should allow Downloads.
        let policy = PathScopePolicy::default_desktop();
        let validated = policy
            .validate_read_path(&file_path)
            .expect("downloads path should be allowed");
        let downloads_root = dunce::canonicalize(&downloads_dir).unwrap();
        assert!(
            validated.starts_with(&downloads_root),
            "expected validated path to be within downloads root"
        );

        // Lower-level fs scope helpers should allow the same roots.
        let allowed_roots = crate::fs_scope::desktop_allowed_roots().unwrap();
        let downloads_root_std = std::fs::canonicalize(&downloads_dir).unwrap();
        assert!(
            allowed_roots.iter().any(|root| root == &downloads_root_std),
            "expected desktop_allowed_roots to include downloads root"
        );
        crate::fs_scope::canonicalize_in_allowed_roots(&file_path, &allowed_roots)
            .expect("downloads path should be allowed by fs_scope");
    }

    #[test]
    fn default_desktop_allows_home_reads_and_rejects_out_of_scope() {
        let base_dirs = BaseDirs::new().expect("base dirs");
        let in_scope_tmp = tempfile::tempdir_in(base_dirs.home_dir()).expect("tempdir in home");
        let in_scope_file = in_scope_tmp.path().join("in-scope.txt");
        std::fs::write(&in_scope_file, "hello").expect("write in-scope file");

        let out_scope_tmp = tempfile::tempdir().expect("tempdir out-of-scope");
        let out_scope_file = out_scope_tmp.path().join("out-of-scope.txt");
        std::fs::write(&out_scope_file, "secret").expect("write out-of-scope file");

        let policy = PathScopePolicy::default_desktop();
        let validated = policy
            .validate_read_path(&in_scope_file)
            .expect("in-scope read");
        assert!(validated.is_absolute());

        let err = policy.validate_read_path(&out_scope_file).unwrap_err();
        assert_eq!(
            err,
            "Access denied: path is outside the allowed filesystem scope"
        );
    }

    #[test]
    fn write_allows_new_files_within_allowed_root() {
        let allowed_temp = tempfile::tempdir().unwrap();
        let candidate = allowed_temp.path().join("new.xlsx");

        let allowed_root = dunce::canonicalize(allowed_temp.path()).unwrap();
        let policy = PathScopePolicy::new(vec![allowed_root.clone()]);
        let validated = policy.validate_write_path(&candidate).unwrap();
        assert!(validated.starts_with(&allowed_root));
        assert_eq!(validated.file_name().unwrap(), "new.xlsx");
    }

    #[test]
    fn write_rejects_paths_outside_allowed_root() {
        let allowed_root = tempfile::tempdir().unwrap();
        let disallowed_root = tempfile::tempdir().unwrap();
        let candidate = disallowed_root.path().join("nope.xlsx");

        let policy = PathScopePolicy::new(vec![allowed_root.path().to_path_buf()]);
        let err = policy.validate_write_path(&candidate).unwrap_err();
        assert_eq!(
            err,
            "Access denied: path is outside the allowed filesystem scope"
        );
    }

    #[test]
    fn default_desktop_allows_home_writes_and_rejects_out_of_scope() {
        let base_dirs = BaseDirs::new().expect("base dirs");
        let in_scope_tmp = tempfile::tempdir_in(base_dirs.home_dir()).expect("tempdir in home");
        let in_scope_dest = in_scope_tmp.path().join("new.xlsx");
        assert!(!in_scope_dest.exists());

        let out_scope_tmp = tempfile::tempdir().expect("tempdir out-of-scope");
        let out_scope_dest = out_scope_tmp.path().join("new.xlsx");
        assert!(!out_scope_dest.exists());

        let policy = PathScopePolicy::default_desktop();
        let validated = policy
            .validate_write_path(&in_scope_dest)
            .expect("in-scope write");
        let expected_parent = dunce::canonicalize(in_scope_tmp.path()).expect("canonicalize parent");
        assert_eq!(validated, expected_parent.join("new.xlsx"));

        let err = policy.validate_write_path(&out_scope_dest).unwrap_err();
        assert_eq!(
            err,
            "Access denied: path is outside the allowed filesystem scope"
        );
    }

    #[cfg(unix)]
    #[test]
    fn symlink_escape_is_rejected_for_reads() {
        use std::os::unix::fs::symlink;

        let allowed_root = tempfile::tempdir().unwrap();
        let disallowed_root = tempfile::tempdir().unwrap();

        let secret = disallowed_root.path().join("secret.txt");
        std::fs::write(&secret, "secret").unwrap();

        let link_path = allowed_root.path().join("link");
        symlink(disallowed_root.path(), &link_path).unwrap();

        let attempted = link_path.join("secret.txt");

        let policy = PathScopePolicy::new(vec![allowed_root.path().to_path_buf()]);
        let err = policy.validate_read_path(&attempted).unwrap_err();
        assert_eq!(
            err,
            "Access denied: path is outside the allowed filesystem scope"
        );
    }
}
