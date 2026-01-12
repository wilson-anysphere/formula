use std::collections::BTreeSet;
use std::fs;
use std::path::PathBuf;

use serde_json::Value as JsonValue;

fn repo_path(relative: &str) -> PathBuf {
    PathBuf::from(env!("CARGO_MANIFEST_DIR")).join(relative)
}

fn parse_invoke_handler_commands(main_rs_src: &str) -> BTreeSet<String> {
    let start_marker = ".invoke_handler(tauri::generate_handler![";
    let start = main_rs_src
        .find(start_marker)
        .unwrap_or_else(|| panic!("failed to find `{start_marker}` in src/main.rs"));

    let rest = &main_rs_src[start + start_marker.len()..];
    let end = rest
        .find("])")
        .unwrap_or_else(|| panic!("failed to find end of `generate_handler![...]` block"));

    let block = &rest[..end];

    let mut commands = BTreeSet::new();
    for line in block.lines() {
        let trimmed = line.trim();
        if trimmed.is_empty() || trimmed.starts_with("//") {
            continue;
        }

        // Example lines:
        // - `commands::open_workbook,`
        // - `tray_status::set_tray_status,`
        // - `show_system_notification,`
        let raw = trimmed
            .split(|c: char| c == ',' || c.is_whitespace())
            .next()
            .unwrap_or("")
            .trim();
        if raw.is_empty() {
            continue;
        }

        // Command names are the function name, not the module path.
        let name = raw.rsplit("::").next().unwrap_or(raw).trim();
        if name.is_empty() {
            continue;
        }

        commands.insert(name.to_string());
    }

    if commands.is_empty() {
        panic!("no commands parsed from invoke_handler list; parser likely broke");
    }

    commands
}

fn parse_allow_invoke_permission_commands(permission_file: &JsonValue) -> BTreeSet<String> {
    let permissions = permission_file
        .get("permission")
        .and_then(|p| p.as_array())
        .unwrap_or_else(|| panic!("permission file missing `permission` array"));

    for perm in permissions {
        let id = perm.get("identifier").and_then(|v| v.as_str());
        if id != Some("allow-invoke") {
            continue;
        }

        let list = perm
            .get("commands")
            .and_then(|c| c.get("allow"))
            .and_then(|v| v.as_array())
            .unwrap_or_else(|| panic!("`allow-invoke` permission missing `commands.allow` array"));

        let raw: Vec<String> = list
            .iter()
            .filter_map(|v| v.as_str().map(|s| s.to_string()))
            .collect();

        if raw.is_empty() {
            panic!("`allow-invoke` permission `commands.allow` is empty");
        }

        let set: BTreeSet<String> = raw.iter().cloned().collect();
        if set.len() != raw.len() {
            let mut counts = std::collections::BTreeMap::<&str, usize>::new();
            for cmd in &raw {
                *counts.entry(cmd.as_str()).or_default() += 1;
            }
            let dups: Vec<String> = counts
                .into_iter()
                .filter_map(|(cmd, n)| (n > 1).then(|| format!("{cmd} (x{n})")))
                .collect();
            panic!(
                "`allow-invoke` permission `commands.allow` contains duplicates: {}",
                dups.join(", ")
            );
        }

        return set;
    }

    panic!("permission file missing `allow-invoke` permission entry")
}

#[test]
fn tauri_main_capability_scopes_to_main_window() {
    let capability_path = repo_path("capabilities/main.json");
    let capability_raw = fs::read_to_string(&capability_path)
        .unwrap_or_else(|err| panic!("failed to read {}: {err}", capability_path.display()));
    let capability: JsonValue =
        serde_json::from_str(&capability_raw).unwrap_or_else(|err| panic!("invalid JSON: {err}"));

    assert_eq!(
        capability.get("identifier").and_then(|v| v.as_str()),
        Some("main"),
        "capabilities/main.json must use identifier \"main\""
    );

    let windows = capability
        .get("windows")
        .and_then(|w| w.as_array())
        .unwrap_or_else(|| panic!("capability missing `windows` array"));
    assert!(
        windows.iter().any(|w| w.as_str() == Some("main")),
        "capabilities/main.json must include `windows: [\"main\"]` so it applies to the main window"
    );

    let permissions = capability
        .get("permissions")
        .and_then(|p| p.as_array())
        .unwrap_or_else(|| panic!("capability missing `permissions` array"));
    assert!(
        permissions.iter().any(|p| p.as_str() == Some("allow-invoke")),
        "capabilities/main.json must include the application permission `allow-invoke` so the explicit IPC command allowlist is enforced"
    );
}

#[test]
fn tauri_ipc_allowlist_matches_registered_invoke_handler_commands() {
    // The desktop frontend uses `globalThis.__TAURI__.core.invoke(...)` directly.
    // This test ensures we keep the invokable command surface explicit and in sync with
    // the backend's `invoke_handler` registration list.

    let main_rs_path = repo_path("src/main.rs");
    let permission_path = repo_path("permissions/allow-invoke.json");

    let main_rs_src = fs::read_to_string(&main_rs_path)
        .unwrap_or_else(|err| panic!("failed to read {}: {err}", main_rs_path.display()));
    let expected = parse_invoke_handler_commands(&main_rs_src);

    let permission_raw = fs::read_to_string(&permission_path)
        .unwrap_or_else(|err| panic!("failed to read {}: {err}", permission_path.display()));
    let permission_file: JsonValue =
        serde_json::from_str(&permission_raw).unwrap_or_else(|err| panic!("invalid JSON: {err}"));
    let actual = parse_allow_invoke_permission_commands(&permission_file);

    for cmd in &actual {
        assert!(
            !cmd.starts_with("plugin:"),
            "capability allowlist must not include plugin commands: {cmd}"
        );
    }

    assert_eq!(
        actual, expected,
        "IPC allowlist mismatch.\n\n\
         - Update `src-tauri/permissions/allow-invoke.json` (`allow-invoke` permission `commands.allow`) to match the command list in `src-tauri/src/main.rs`.\n\
         - Or, if you removed a command, remove it from both places.\n"
    );
}
