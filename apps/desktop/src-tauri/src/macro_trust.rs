use std::collections::{HashMap, HashSet};
use std::fs;
use std::path::{Path, PathBuf};

use crate::atomic_write::write_file_atomic;
use anyhow::Context;
use directories::ProjectDirs;
use serde::{Deserialize, Serialize};
use sha2::{Digest, Sha256};

pub type SharedMacroTrustStore = std::sync::Arc<std::sync::Mutex<MacroTrustStore>>;

#[derive(Clone, Debug, Serialize, Deserialize, PartialEq, Eq)]
#[serde(rename_all = "snake_case")]
pub enum MacroTrustDecision {
    /// Default state: macros are blocked.
    Blocked,
    /// Trust this workbook's macros permanently (persisted on disk).
    TrustedAlways,
    /// Trust this workbook's macros for the current app session only.
    TrustedOnce,
    /// Trust only if the workbook is signed and the signature verifies (best-effort).
    TrustedSignedOnly,
}

#[derive(Clone, Debug, Serialize, Deserialize, PartialEq, Eq)]
#[serde(rename_all = "snake_case")]
enum PersistedDecision {
    TrustedAlways,
    TrustedSignedOnly,
}

#[derive(Clone, Debug, Serialize, Deserialize)]
struct TrustStoreFile {
    version: u32,
    entries: HashMap<String, PersistedDecision>,
}

impl Default for TrustStoreFile {
    fn default() -> Self {
        Self {
            version: 1,
            entries: HashMap::new(),
        }
    }
}

#[derive(Debug)]
pub struct MacroTrustStore {
    path: Option<PathBuf>,
    persisted: HashMap<String, PersistedDecision>,
    trusted_once: HashSet<String>,
    /// Whether the persisted on-disk trust decisions have been loaded.
    ///
    /// This exists so the desktop shell can defer filesystem I/O (reading the trust store file)
    /// until after the first window is visible. While `loaded == false`, macro execution is
    /// conservatively blocked regardless of persisted state (default-deny).
    loaded: bool,
}

impl MacroTrustStore {
    pub fn new_ephemeral() -> Self {
        Self {
            path: None,
            persisted: HashMap::new(),
            trusted_once: HashSet::new(),
            loaded: true,
        }
    }

    pub fn load(path: PathBuf) -> anyhow::Result<Self> {
        let persisted = load_trust_file(&path).unwrap_or_default().entries;

        Ok(Self {
            path: Some(path),
            persisted,
            trusted_once: HashSet::new(),
            loaded: true,
        })
    }

    pub fn load_default() -> anyhow::Result<Self> {
        let Some(path) = default_trust_store_path() else {
            return Ok(Self::new_ephemeral());
        };
        Self::load(path)
    }

    /// Construct the default trust store without performing any filesystem reads.
    ///
    /// The store starts in a "not loaded" state. Call [`Self::ensure_loaded`] (typically from a
    /// background task) before relying on persisted trust decisions.
    pub fn new_unloaded_default() -> Self {
        let Some(path) = default_trust_store_path() else {
            return Self::new_ephemeral();
        };
        Self {
            path: Some(path),
            persisted: HashMap::new(),
            trusted_once: HashSet::new(),
            loaded: false,
        }
    }

    pub fn is_loaded(&self) -> bool {
        self.loaded
    }

    /// Load persisted trust decisions from disk if they have not been loaded yet.
    pub fn ensure_loaded(&mut self) {
        if self.loaded {
            return;
        }
        let Some(path) = self.path.clone() else {
            // Ephemeral store: nothing to load.
            self.loaded = true;
            return;
        };
        self.persisted = load_trust_file(&path).unwrap_or_default().entries;
        self.loaded = true;
    }

    pub fn trust_state(&self, fingerprint: &str) -> MacroTrustDecision {
        // Default-deny until we have loaded persisted state from disk. This avoids accidentally
        // trusting macros before the trust store is ready (and allows desktop startup to defer
        // filesystem I/O until after the window is visible).
        if !self.loaded {
            return MacroTrustDecision::Blocked;
        }
        if self.trusted_once.contains(fingerprint) {
            return MacroTrustDecision::TrustedOnce;
        }
        match self.persisted.get(fingerprint) {
            Some(PersistedDecision::TrustedAlways) => MacroTrustDecision::TrustedAlways,
            Some(PersistedDecision::TrustedSignedOnly) => MacroTrustDecision::TrustedSignedOnly,
            None => MacroTrustDecision::Blocked,
        }
    }

    pub fn set_trust(
        &mut self,
        fingerprint: String,
        decision: MacroTrustDecision,
    ) -> anyhow::Result<()> {
        // Avoid clobbering an existing trust store file. If we haven't loaded persisted entries
        // yet, load them now before mutating/saving.
        self.ensure_loaded();

        match decision {
            MacroTrustDecision::Blocked => {
                self.trusted_once.remove(&fingerprint);
                self.persisted.remove(&fingerprint);
            }
            MacroTrustDecision::TrustedOnce => {
                // Session-only trust should not keep any persisted allow-list entry.
                self.persisted.remove(&fingerprint);
                self.trusted_once.insert(fingerprint);
            }
            MacroTrustDecision::TrustedAlways => {
                self.trusted_once.remove(&fingerprint);
                self.persisted
                    .insert(fingerprint, PersistedDecision::TrustedAlways);
            }
            MacroTrustDecision::TrustedSignedOnly => {
                self.trusted_once.remove(&fingerprint);
                self.persisted
                    .insert(fingerprint, PersistedDecision::TrustedSignedOnly);
            }
        }

        self.save()
    }

    fn save(&self) -> anyhow::Result<()> {
        let Some(path) = self.path.as_ref() else {
            return Ok(());
        };
        let file = TrustStoreFile {
            version: 1,
            entries: self.persisted.clone(),
        };
        let json = serde_json::to_vec_pretty(&file).context("serialize macro trust store")?;
        write_file_atomic(path, &json)
            .with_context(|| format!("write macro trust store {path:?}"))?;
        Ok(())
    }
}

pub fn default_trust_store_path() -> Option<PathBuf> {
    // Keep the on-disk format stable and backend-owned. This should not be a user-facing
    // preferences mechanism; it is the desktop app's "Trust Center" datastore.
    let proj = ProjectDirs::from("com", "formula", "Formula")?;
    Some(proj.config_dir().join("macro_trust.json"))
}

pub fn compute_macro_fingerprint(workbook_id: &str, vba_project_bin: &[u8]) -> String {
    // Versioned fingerprint scheme so we can change it in the future without silently
    // reusing old trust decisions.
    const PREFIX: &[u8] = b"formula-macro-fingerprint-v1\0";
    let mut hasher = Sha256::new();
    hasher.update(PREFIX);
    hasher.update(workbook_id.as_bytes());
    hasher.update(b"\0");
    hasher.update(vba_project_bin);
    hex::encode(hasher.finalize())
}

fn load_trust_file(path: &Path) -> anyhow::Result<TrustStoreFile> {
    let bytes = match fs::read(path) {
        Ok(bytes) => bytes,
        Err(err) if err.kind() == std::io::ErrorKind::NotFound => {
            return Ok(TrustStoreFile::default())
        }
        Err(err) => return Err(err).with_context(|| format!("read macro trust store {path:?}")),
    };

    let parsed: TrustStoreFile =
        serde_json::from_slice(&bytes).context("parse macro trust store json")?;
    Ok(parsed)
}

#[cfg(test)]
mod tests {
    use super::*;

    #[test]
    fn fingerprint_is_stable_and_sensitive_to_inputs() {
        let fp1 = compute_macro_fingerprint("wb1", b"vba");
        let fp2 = compute_macro_fingerprint("wb1", b"vba");
        assert_eq!(fp1, fp2);

        let fp_other_id = compute_macro_fingerprint("wb2", b"vba");
        assert_ne!(fp1, fp_other_id);

        let fp_other_vba = compute_macro_fingerprint("wb1", b"vba2");
        assert_ne!(fp1, fp_other_vba);
    }

    #[test]
    fn trusted_once_is_session_only() {
        let dir = tempfile::tempdir().expect("tempdir");
        let path = dir.path().join("trust.json");

        let mut store = MacroTrustStore::load(path.clone()).expect("load store");
        store
            .set_trust("fp".to_string(), MacroTrustDecision::TrustedOnce)
            .expect("set trust");
        assert_eq!(store.trust_state("fp"), MacroTrustDecision::TrustedOnce);

        drop(store);
        let store2 = MacroTrustStore::load(path).expect("reload store");
        assert_eq!(store2.trust_state("fp"), MacroTrustDecision::Blocked);
    }

    #[test]
    fn trusted_always_persists() {
        let dir = tempfile::tempdir().expect("tempdir");
        let path = dir.path().join("trust.json");

        let mut store = MacroTrustStore::load(path.clone()).expect("load store");
        store
            .set_trust("fp".to_string(), MacroTrustDecision::TrustedAlways)
            .expect("set trust");
        assert_eq!(store.trust_state("fp"), MacroTrustDecision::TrustedAlways);

        drop(store);
        let store2 = MacroTrustStore::load(path).expect("reload store");
        assert_eq!(store2.trust_state("fp"), MacroTrustDecision::TrustedAlways);
    }

    #[test]
    fn persisted_trust_is_default_deny_until_store_is_loaded() {
        let dir = tempfile::tempdir().expect("tempdir");
        let path = dir.path().join("trust.json");

        // Persist a trust decision on disk.
        let mut store = MacroTrustStore::load(path.clone()).expect("load store");
        store
            .set_trust("fp".to_string(), MacroTrustDecision::TrustedAlways)
            .expect("set trust");
        drop(store);

        // Simulate the desktop cold-start behavior: create the store without reading from disk.
        let mut unloaded = MacroTrustStore {
            path: Some(path),
            persisted: HashMap::new(),
            trusted_once: HashSet::new(),
            loaded: false,
        };

        // Default-deny until `ensure_loaded` runs.
        assert_eq!(unloaded.trust_state("fp"), MacroTrustDecision::Blocked);

        unloaded.ensure_loaded();
        assert!(unloaded.is_loaded());
        assert_eq!(
            unloaded.trust_state("fp"),
            MacroTrustDecision::TrustedAlways
        );
    }
}
