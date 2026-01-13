use std::fs;
use std::path::PathBuf;

use serde_json::Value as JsonValue;

fn repo_path(relative: &str) -> PathBuf {
    PathBuf::from(env!("CARGO_MANIFEST_DIR")).join(relative)
}

#[test]
fn windows_bundle_configures_webview2_install_mode() {
    let tauri_conf_path = repo_path("tauri.conf.json");
    let conf_raw = fs::read_to_string(&tauri_conf_path)
        .unwrap_or_else(|err| panic!("failed to read {}: {err}", tauri_conf_path.display()));
    let conf: JsonValue = serde_json::from_str(&conf_raw).unwrap_or_else(|err| {
        panic!(
            "invalid JSON in {}: {err}",
            tauri_conf_path.display()
        )
    });

    let windows_bundle = conf
        .get("bundle")
        .and_then(|bundle| bundle.get("windows"))
        .unwrap_or_else(|| panic!("tauri.conf.json missing `bundle.windows` object"));

    let mode = windows_bundle.get("webviewInstallMode").unwrap_or_else(|| {
        panic!(
            "tauri.conf.json missing `bundle.windows.webviewInstallMode`.\n\n\
             Formula relies on the Microsoft Edge WebView2 Evergreen Runtime on Windows.\n\
             Configure the installer to install WebView2 automatically on clean machines, e.g.:\n\
             - {{ \"type\": \"downloadBootstrapper\" }} (small, downloads runtime during install)\n\
             - {{ \"type\": \"offlineInstaller\" }} / {{ \"type\": \"fixedRuntime\" }} (offline, larger)\n"
        )
    });

    let raw_type = match mode {
        JsonValue::String(s) => s.as_str(),
        JsonValue::Object(obj) => obj
            .get("type")
            .and_then(|v| v.as_str())
            .unwrap_or_else(|| panic!("`bundle.windows.webviewInstallMode` object missing `type`")),
        other => panic!(
            "`bundle.windows.webviewInstallMode` must be a string or object; found: {other}"
        ),
    };

    let trimmed = raw_type.trim();
    assert!(
        !trimmed.is_empty(),
        "`bundle.windows.webviewInstallMode` must not be empty"
    );

    // Keep this allowlist in sync with `scripts/ci/check-webview2-install-mode.mjs`.
    const ALLOWED: &[&str] = &[
        "downloadBootstrapper",
        "embedBootstrapper",
        "offlineInstaller",
        "fixedRuntime",
    ];

    assert_ne!(
        trimmed.to_ascii_lowercase(),
        "skip",
        "`bundle.windows.webviewInstallMode` must not be 'skip' (clean Windows machines may not have WebView2 installed)"
    );

    assert!(
        ALLOWED.contains(&trimmed),
        "`bundle.windows.webviewInstallMode` must be one of {:?}; got {trimmed}",
        ALLOWED
    );
}
