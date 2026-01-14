use desktop::commands::LimitedString;

#[test]
fn limited_string_deserialize_enforces_max_len() {
    let ok: LimitedString<4> = serde_json::from_str("\"test\"").expect("expected payload to deserialize");
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
    assert_fn_has_ipc_string_cap(commands_src, "pub async fn list_dir", &["MAX_IPC_PATH_BYTES"]);
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
        &["MAX_IPC_URL_BYTES"],
    );
}
