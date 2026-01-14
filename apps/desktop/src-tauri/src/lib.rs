pub(crate) mod atomic_write;
pub mod asset_protocol_core;
pub mod asset_protocol_policy;
pub mod commands;
pub mod deep_link_schemes;
pub mod ed25519_verifier;
pub mod external_url;
#[cfg(any(feature = "desktop", test))]
pub mod clipboard;
// Clipboard fallback helpers are only used by the Linux clipboard backend, which itself is behind
// the `desktop` feature. Keep this module behind the same gate so non-desktop builds don't pick up
// dead_code warnings.
#[cfg(all(target_os = "linux", any(feature = "desktop", test)))]
mod clipboard_fallback;
pub mod file_io;
mod fs_scope;
#[cfg(any(feature = "desktop", test))]
mod ipc_file_limits;
pub mod ipc_origin;
pub mod ipc_limits;
pub mod macro_trust;
pub mod macros;
#[cfg(any(feature = "desktop", test))]
pub mod network_limits;
pub mod open_file;
pub mod open_file_ipc;
pub mod oauth_redirect_ipc;
pub mod opened_urls;
pub mod oauth_loopback;
pub mod oauth_redirect;
pub mod persistence;
pub mod power_query_validation;
pub mod python;
#[cfg(any(feature = "desktop", test))]
pub mod pyodide_assets;
#[cfg(any(feature = "desktop", test))]
pub mod network_fetch;
pub mod resource_limits;
pub(crate) mod sheet_name;
pub mod sql;
pub mod state;
pub mod storage;
pub mod tauri_origin;
#[cfg(feature = "process-metrics")]
pub mod process_metrics;
#[cfg(feature = "desktop")]
pub mod tray_status;
#[cfg(any(feature = "desktop", test))]
pub mod updater_download_cache;
#[cfg(feature = "desktop")]
pub mod updater;
