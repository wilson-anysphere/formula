use std::fs;
use std::path::PathBuf;

use serde_json::Value as JsonValue;

fn repo_path(relative: &str) -> PathBuf {
    PathBuf::from(env!("CARGO_MANIFEST_DIR")).join(relative)
}

#[test]
fn windows_bundle_targets_include_msi_and_nsis() {
    let tauri_conf_path = repo_path("tauri.conf.json");
    let conf_raw = fs::read_to_string(&tauri_conf_path)
        .unwrap_or_else(|err| panic!("failed to read {}: {err}", tauri_conf_path.display()));
    let conf: JsonValue = serde_json::from_str(&conf_raw).unwrap_or_else(|err| {
        panic!(
            "invalid JSON in {}: {err}",
            tauri_conf_path.display()
        )
    });

    let bundle = conf
        .get("bundle")
        .unwrap_or_else(|| panic!("tauri.conf.json missing `bundle` object"));

    let targets = bundle.get("targets").unwrap_or_else(|| {
        panic!(
            "tauri.conf.json missing `bundle.targets`.\n\n\
             Windows releases must ship **both** installer formats:\n\
             - WiX/MSI (`.msi`)\n\
             - NSIS (`.exe`)\n\
             \n\
             Keep `bundle.targets` as \"all\" (default) or explicitly list `msi` + `nsis`."
        )
    });

    match targets {
        JsonValue::String(s) => {
            let trimmed = s.trim().to_ascii_lowercase();
            if trimmed == "all" {
                return;
            }
            panic!(
                "`bundle.targets` is set to {s:?}. Expected \"all\" (which includes Windows MSI + NSIS) \
                 or an explicit array containing `msi` and `nsis`/`nsis-web`."
            );
        }
        JsonValue::Array(arr) => {
            let mut values: Vec<String> = arr
                .iter()
                .filter_map(|v| v.as_str())
                .map(|s| s.trim().to_ascii_lowercase())
                .filter(|s| !s.is_empty())
                .collect();
            values.sort();
            values.dedup();

            if values.iter().any(|v| v == "all") {
                return;
            }

            assert!(
                values.iter().any(|v| v == "msi"),
                "`bundle.targets` must include `msi` so Windows releases produce a .msi installer (WiX). Got: {values:?}"
            );
            assert!(
                values.iter().any(|v| v == "nsis") || values.iter().any(|v| v == "nsis-web"),
                "`bundle.targets` must include `nsis` (or `nsis-web`) so Windows releases produce a .exe installer. Got: {values:?}"
            );
        }
        other => panic!(
            "`bundle.targets` must be a string or array; found: {other}. Expected \"all\" or an array containing `msi` + `nsis`."
        ),
    }
}

