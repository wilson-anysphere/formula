use std::fs;
use std::path::PathBuf;

use serde_json::Value as JsonValue;

fn repo_path(relative: &str) -> PathBuf {
    PathBuf::from(env!("CARGO_MANIFEST_DIR")).join(relative)
}

#[test]
fn tauri_macos_signing_identity_is_not_hardcoded() {
    // Guardrail: the committed Tauri config should not hardcode a macOS signing identity.
    //
    // Rationale:
    // - Local `tauri build` should succeed on macOS machines without any Developer ID certificates
    //   installed (unsigned builds).
    // - Release CI should set an explicit identity via `APPLE_SIGNING_IDENTITY` (avoids ambiguous
    //   selection when multiple certs exist).
    //
    // The release workflow may patch `tauri.conf.json` at build time, but the version committed to
    // the repo should keep signing disabled by default.
    let tauri_conf_path = repo_path("tauri.conf.json");
    let conf_raw = fs::read_to_string(&tauri_conf_path)
        .unwrap_or_else(|err| panic!("failed to read {}: {err}", tauri_conf_path.display()));
    let conf: JsonValue =
        serde_json::from_str(&conf_raw).unwrap_or_else(|err| panic!("invalid JSON: {err}"));

    let signing_identity = conf
        .get("bundle")
        .and_then(|bundle| bundle.get("macOS"))
        .and_then(|mac| mac.get("signingIdentity"));

    match signing_identity {
        None | Some(JsonValue::Null) => {}
        Some(other) => panic!(
            "tauri.conf.json must not hardcode bundle.macOS.signingIdentity (expected null/absent, found: {other})"
        ),
    }
}

