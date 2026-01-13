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

fn parse_core_allow_invoke_commands(capability_file: &JsonValue) -> Option<BTreeSet<String>> {
    let permissions = capability_file
        .get("permissions")
        .and_then(|p| p.as_array())
        .unwrap_or_else(|| panic!("capability missing `permissions` array"));

    let Some(core_allow_invoke) = permissions.iter().find(|p| {
        p.get("identifier")
            .and_then(|v| v.as_str())
            .is_some_and(|id| id == "core:allow-invoke")
    }) else {
        return None;
    };

    let allow = core_allow_invoke
        .get("allow")
        .and_then(|v| v.as_array())
        .unwrap_or_else(|| panic!("`core:allow-invoke` must include an `allow` array"));
    if allow.is_empty() {
        panic!("`core:allow-invoke` must not have an empty allowlist");
    }

    let mut commands = BTreeSet::new();
    for entry in allow {
        let cmd = entry
            .get("command")
            .and_then(|v| v.as_str())
            .unwrap_or_else(|| panic!("`core:allow-invoke` allow entry missing `command`"));
        assert!(
            !cmd.trim().is_empty(),
            "`core:allow-invoke` allow entry command must not be empty"
        );
        assert!(
            !cmd.contains('*'),
            "`core:allow-invoke` command must not contain wildcard patterns: {cmd}"
        );
        assert!(
            commands.insert(cmd.to_string()),
            "`core:allow-invoke` contains duplicate command: {cmd}"
        );
    }

    Some(commands)
}

fn format_command_bullets(commands: &BTreeSet<String>) -> String {
    if commands.is_empty() {
        "  (none)".to_string()
    } else {
        commands
            .iter()
            .map(|cmd| format!("  - {cmd}"))
            .collect::<Vec<_>>()
            .join("\n")
    }
}

#[test]
fn tauri_main_capability_scopes_to_main_window() {
    let tauri_conf_path = repo_path("tauri.conf.json");
    let conf_raw = fs::read_to_string(&tauri_conf_path)
        .unwrap_or_else(|err| panic!("failed to read {}: {err}", tauri_conf_path.display()));
    let conf: JsonValue =
        serde_json::from_str(&conf_raw).unwrap_or_else(|err| panic!("invalid JSON: {err}"));

    let windows = conf
        .get("app")
        .and_then(|app| app.get("windows"))
        .and_then(|w| w.as_array())
        .unwrap_or_else(|| panic!("tauri.conf.json missing `app.windows` array"));
    let main_window = windows
        .iter()
        .find(|w| w.get("label").and_then(|v| v.as_str()) == Some("main"))
        .unwrap_or_else(|| panic!("tauri.conf.json missing main window (label \"main\")"));

    assert!(
        main_window.get("capabilities").is_none(),
        "tauri.conf.json window-level `capabilities` is not supported by tauri-build; window scoping should be done via `capabilities/main.json` instead"
    );

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

    assert!(
        !permissions.iter().any(|p| p.as_str() == Some("core:allow-invoke")),
        "capabilities/main.json must not include `core:allow-invoke` as a string; if present it must be scoped via the object form with an explicit per-command allowlist"
    );

    // Some Tauri toolchains expose a `core:allow-invoke` permission for per-command allowlisting.
    //
    // We primarily rely on the application permission defined in `permissions/allow-invoke.json`,
    // but if `core:allow-invoke` is present it must be scoped explicitly (no wildcard/pattern
    // matches).
    let _ = parse_core_allow_invoke_commands(&capability);
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

#[test]
fn tauri_core_allow_invoke_is_subset_of_allow_invoke_permission() {
    let capability_path = repo_path("capabilities/main.json");
    let capability_raw = fs::read_to_string(&capability_path)
        .unwrap_or_else(|err| panic!("failed to read {}: {err}", capability_path.display()));
    let capability_file: JsonValue =
        serde_json::from_str(&capability_raw).unwrap_or_else(|err| panic!("invalid JSON: {err}"));

    let Some(core_allow_invoke) = parse_core_allow_invoke_commands(&capability_file) else {
        // Some toolchains may not expose the `core:allow-invoke` permission.
        return;
    };

    let permission_path = repo_path("permissions/allow-invoke.json");
    let permission_raw = fs::read_to_string(&permission_path)
        .unwrap_or_else(|err| panic!("failed to read {}: {err}", permission_path.display()));
    let permission_file: JsonValue =
        serde_json::from_str(&permission_raw).unwrap_or_else(|err| panic!("invalid JSON: {err}"));
    let allow_invoke = parse_allow_invoke_permission_commands(&permission_file);

    // `core:allow-invoke` is a defense-in-depth allowlist, but it must remain *identical* to the
    // canonical application permission allowlist (`allow-invoke`).
    //
    // Subset checks aren't enough: missing commands will break IPC on toolchains that enforce
    // `core:allow-invoke`, and extra commands widen the available surface area.
    let missing_from_core: BTreeSet<String> = allow_invoke
        .difference(&core_allow_invoke)
        .cloned()
        .collect();
    let only_in_core: BTreeSet<String> = core_allow_invoke
        .difference(&allow_invoke)
        .cloned()
        .collect();

    if !missing_from_core.is_empty() || !only_in_core.is_empty() {
        panic!(
            "`core:allow-invoke` command list drift detected.\n\n\
             The per-command allowlist in `src-tauri/capabilities/main.json` (`core:allow-invoke`) must be identical to the canonical application permission allowlist in `src-tauri/permissions/allow-invoke.json` (`allow-invoke` -> `commands.allow`).\n\n\
             Commands missing from `core:allow-invoke` (present in `allow-invoke`):\n{}\n\n\
             Commands present only in `core:allow-invoke` (not in `allow-invoke`):\n{}\n\n\
             Keep these lists identical to avoid subtle capability drift.\n",
            format_command_bullets(&missing_from_core),
            format_command_bullets(&only_in_core)
        );
    }
}
