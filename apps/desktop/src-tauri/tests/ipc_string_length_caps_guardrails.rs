use desktop::commands::LimitedString;

#[test]
fn limited_string_deserialize_enforces_max_len() {
    let ok: LimitedString<4> =
        serde_json::from_str("\"test\"").expect("expected payload to deserialize");
    assert_eq!(ok.as_ref(), "test");

    let err = serde_json::from_str::<LimitedString<4>>("\"tests\"")
        .expect_err("expected oversized payload to fail during deserialization");
    let msg = err.to_string();
    assert!(
        msg.contains("max 4 bytes"),
        "expected error to mention max length, got: {msg}"
    );
}

fn assert_fn_has_ipc_string_cap(src: &str, fn_decl: &str, limit_tokens: &[&str]) {
    let start = src
        .find(fn_decl)
        .unwrap_or_else(|| panic!("expected function declaration {fn_decl:?} to exist"));
    let body = &src[start..];
    // Use a fixed-size slice instead of trying to parse the function header, because Rust const
    // generics (`LimitedString<{ ... }>`) contain `{` braces which can confuse naive header
    // detection.
    let snippet = &body[..body.len().min(2000)];

    let has_limited_string = snippet.contains("LimitedString");
    let has_enforce_helper = snippet.contains("enforce_ipc_string_byte_len(");

    let snippet_has_limits = limit_tokens.iter().all(|t| snippet.contains(t));

    assert!(
        (has_limited_string && snippet_has_limits) || (has_enforce_helper && snippet_has_limits),
        "expected {fn_decl} to enforce IPC string caps via LimitedString<...> or enforce_ipc_string_byte_len(...); \
limits={limit_tokens:?}\n--- snippet ---\n{}",
        &snippet[..snippet.len().min(400)]
    );
}

fn assert_marker_has_ipc_string_cap(src: &str, marker: &str, limit_tokens: &[&str]) {
    let start = src
        .find(marker)
        .unwrap_or_else(|| panic!("expected marker {marker:?} to exist"));
    let body = &src[start..];
    let snippet = &body[..body.len().min(2000)];

    let snippet_has_limits = limit_tokens.iter().all(|t| snippet.contains(t));
    assert!(
        snippet.contains("LimitedString") && snippet_has_limits,
        "expected {marker} to reference LimitedString<...> with limits={limit_tokens:?}\n--- snippet ---\n{}",
        &snippet[..snippet.len().min(400)]
    );
}

#[test]
fn privileged_ipc_commands_have_string_length_caps() {
    let commands_src = include_str!("../src/commands.rs");
    assert_fn_has_ipc_string_cap(
        commands_src,
        "pub async fn open_workbook",
        &["MAX_IPC_PATH_BYTES"],
    );
    assert_fn_has_ipc_string_cap(
        commands_src,
        "pub fn inspect_workbook_encryption",
        &["MAX_IPC_PATH_BYTES"],
    );
    assert_fn_has_ipc_string_cap(
        commands_src,
        "pub async fn save_workbook",
        &["MAX_IPC_PATH_BYTES"],
    );
    assert_fn_has_ipc_string_cap(
        commands_src,
        "pub async fn read_text_file",
        &["MAX_IPC_PATH_BYTES"],
    );
    assert_fn_has_ipc_string_cap(
        commands_src,
        "pub async fn stat_file",
        &["MAX_IPC_PATH_BYTES"],
    );
    assert_fn_has_ipc_string_cap(
        commands_src,
        "pub async fn read_binary_file(",
        &["MAX_IPC_PATH_BYTES"],
    );
    assert_fn_has_ipc_string_cap(
        commands_src,
        "pub async fn read_binary_file_range",
        &["MAX_IPC_PATH_BYTES"],
    );
    assert_fn_has_ipc_string_cap(
        commands_src,
        "pub async fn list_dir",
        &["MAX_IPC_PATH_BYTES"],
    );
    assert_fn_has_ipc_string_cap(
        commands_src,
        "pub async fn open_external_url",
        &["MAX_IPC_URL_BYTES"],
    );
    assert_fn_has_ipc_string_cap(
        commands_src,
        "pub async fn network_fetch",
        &["MAX_IPC_URL_BYTES"],
    );
    assert_fn_has_ipc_string_cap(
        commands_src,
        "pub fn power_query_state_set",
        &["MAX_POWER_QUERY_XML_BYTES"],
    );
    assert_fn_has_ipc_string_cap(
        commands_src,
        "pub async fn collab_token_get",
        &["MAX_IPC_SECURE_STORE_KEY_BYTES"],
    );
    assert_fn_has_ipc_string_cap(
        commands_src,
        "pub async fn collab_token_set",
        &["MAX_IPC_SECURE_STORE_KEY_BYTES"],
    );
    assert_fn_has_ipc_string_cap(
        commands_src,
        "pub async fn collab_token_delete",
        &["MAX_IPC_SECURE_STORE_KEY_BYTES"],
    );
    assert_marker_has_ipc_string_cap(
        commands_src,
        "pub struct CollabTokenEntryIpc",
        &["MAX_IPC_COLLAB_TOKEN_BYTES"],
    );
    assert_fn_has_ipc_string_cap(
        commands_src,
        "pub async fn power_query_credential_get",
        &["MAX_CREDENTIAL_SCOPE_KEY_LEN"],
    );
    assert_fn_has_ipc_string_cap(
        commands_src,
        "pub async fn power_query_credential_set",
        &["MAX_CREDENTIAL_SCOPE_KEY_LEN"],
    );
    assert_fn_has_ipc_string_cap(
        commands_src,
        "pub async fn power_query_credential_delete",
        &["MAX_CREDENTIAL_SCOPE_KEY_LEN"],
    );
    assert_fn_has_ipc_string_cap(
        commands_src,
        "pub async fn power_query_refresh_state_get",
        &["MAX_IPC_SECURE_STORE_KEY_BYTES"],
    );
    assert_fn_has_ipc_string_cap(
        commands_src,
        "pub async fn power_query_refresh_state_set",
        &["MAX_IPC_SECURE_STORE_KEY_BYTES"],
    );
    assert_fn_has_ipc_string_cap(
        commands_src,
        "pub async fn collab_encryption_key_get",
        &["MAX_IPC_SECURE_STORE_KEY_BYTES"],
    );
    assert_fn_has_ipc_string_cap(
        commands_src,
        "pub async fn collab_encryption_key_set",
        &[
            "MAX_IPC_SECURE_STORE_KEY_BYTES",
            "MAX_IPC_COLLAB_ENCRYPTION_KEY_BASE64_BYTES",
        ],
    );
    assert_fn_has_ipc_string_cap(
        commands_src,
        "pub async fn collab_encryption_key_delete",
        &["MAX_IPC_SECURE_STORE_KEY_BYTES"],
    );
    assert_fn_has_ipc_string_cap(
        commands_src,
        "pub async fn collab_encryption_key_list",
        &["MAX_IPC_SECURE_STORE_KEY_BYTES"],
    );
    assert_fn_has_ipc_string_cap(
        commands_src,
        "pub async fn get_macro_security_status",
        &["MAX_IPC_SECURE_STORE_KEY_BYTES"],
    );
    assert_fn_has_ipc_string_cap(
        commands_src,
        "pub async fn set_macro_trust",
        &["MAX_IPC_SECURE_STORE_KEY_BYTES"],
    );
    assert_fn_has_ipc_string_cap(
        commands_src,
        "pub fn get_vba_project",
        &["MAX_IPC_SECURE_STORE_KEY_BYTES"],
    );
    assert_fn_has_ipc_string_cap(
        commands_src,
        "pub fn list_macros",
        &["MAX_IPC_SECURE_STORE_KEY_BYTES"],
    );
    assert_fn_has_ipc_string_cap(
        commands_src,
        "pub fn set_macro_ui_context",
        &["MAX_IPC_SECURE_STORE_KEY_BYTES", "MAX_SHEET_ID_BYTES"],
    );
    assert_fn_has_ipc_string_cap(
        commands_src,
        "pub async fn run_macro",
        &["MAX_IPC_SECURE_STORE_KEY_BYTES"],
    );
    assert_fn_has_ipc_string_cap(
        commands_src,
        "pub async fn run_python_script",
        &["MAX_IPC_SECURE_STORE_KEY_BYTES"],
    );
    assert_fn_has_ipc_string_cap(
        commands_src,
        "pub async fn validate_vba_migration",
        &["MAX_IPC_SECURE_STORE_KEY_BYTES"],
    );
    assert_fn_has_ipc_string_cap(
        commands_src,
        "pub async fn fire_workbook_open",
        &["MAX_IPC_SECURE_STORE_KEY_BYTES"],
    );
    assert_fn_has_ipc_string_cap(
        commands_src,
        "pub async fn fire_workbook_before_close",
        &["MAX_IPC_SECURE_STORE_KEY_BYTES"],
    );
    assert_fn_has_ipc_string_cap(
        commands_src,
        "pub async fn fire_worksheet_change",
        &["MAX_IPC_SECURE_STORE_KEY_BYTES", "MAX_SHEET_ID_BYTES"],
    );
    assert_fn_has_ipc_string_cap(
        commands_src,
        "pub async fn fire_selection_change",
        &["MAX_IPC_SECURE_STORE_KEY_BYTES", "MAX_SHEET_ID_BYTES"],
    );

    let main_src = include_str!("../src/main.rs");
    assert_fn_has_ipc_string_cap(
        main_src,
        "async fn show_system_notification",
        &[
            "MAX_IPC_NOTIFICATION_TITLE_BYTES",
            "MAX_IPC_NOTIFICATION_BODY_BYTES",
        ],
    );
    assert_fn_has_ipc_string_cap(
        main_src,
        "async fn oauth_loopback_listen",
        &["MAX_OAUTH_REDIRECT_URI_BYTES"],
    );

    let tray_src = include_str!("../src/tray_status.rs");
    assert_fn_has_ipc_string_cap(
        tray_src,
        "pub fn set_tray_status",
        &["MAX_IPC_TRAY_STATUS_BYTES"],
    );
}
