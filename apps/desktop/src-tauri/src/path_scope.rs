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

    /// Default policy for the desktop app: allow `$HOME/**` and `$DOCUMENT/**`.
    pub fn default_desktop() -> Self {
        let mut roots = Vec::new();

        if let Some(user_dirs) = directories::UserDirs::new() {
            roots.push(user_dirs.home_dir().to_path_buf());
            if let Some(documents) = user_dirs.document_dir() {
                roots.push(documents.to_path_buf());
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

        // `Path::parent()` returns an empty path for single-component relative paths ("foo.xlsx").
        // Canonicalize treats this inconsistently across platforms, so normalize to ".".
        let parent = match input.parent() {
            Some(parent) if !parent.as_os_str().is_empty() => parent,
            _ => Path::new("."),
        };

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
