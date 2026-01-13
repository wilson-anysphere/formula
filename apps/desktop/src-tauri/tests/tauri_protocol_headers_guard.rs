use std::fs;
use std::path::PathBuf;

fn repo_path(relative: &str) -> PathBuf {
    PathBuf::from(env!("CARGO_MANIFEST_DIR")).join(relative)
}

fn load_main_rs_source() -> String {
    let main_rs_path = repo_path("src/main.rs");
    fs::read_to_string(&main_rs_path)
        .unwrap_or_else(|err| panic!("failed to read {}: {err}", main_rs_path.display()))
}

fn extract_tauri_scheme_protocol_handler_block<'a>(main_rs_src: &'a str) -> &'a str {
    // The production desktop binary overrides Tauri's internal asset protocol handler for the
    // `tauri://` scheme so we can inject COOP/COEP headers (required for SharedArrayBuffer /
    // `globalThis.crossOriginIsolated`) while still preserving the CSP computed by Tauri's
    // AssetResolver.
    //
    // This file (`src/main.rs`) only compiles with `--features desktop` because it depends on the
    // system WebView toolchain, so we use a source-level scan to guard against accidental header
    // regressions in headless CI.
    let start_marker = ".register_uri_scheme_protocol(\"tauri\"";
    let start = main_rs_src
        .find(start_marker)
        .unwrap_or_else(|| panic!("failed to find `{start_marker}` in src/main.rs"));

    let rest = &main_rs_src[start..];

    // The `tauri://` protocol handler is registered in the `tauri::Builder` chain and is
    // immediately followed by a `.plugin(...)` call in current production builds.
    //
    // We intentionally avoid brittle brace/paren matching and instead slice the builder chain at
    // the next method call boundary.
    let end = rest
        .find(".plugin(")
        .or_else(|| rest.find(".build("))
        .unwrap_or(rest.len());

    &rest[..end]
}

#[test]
fn tauri_scheme_protocol_handler_injects_cross_origin_isolation_headers_and_preserves_csp() {
    let main_rs_src = load_main_rs_source();

    // 1) Ensure `src/main.rs` registers a custom `tauri://` protocol handler.
    assert!(
        main_rs_src.contains(".register_uri_scheme_protocol(\"tauri\""),
        "expected src/main.rs to register a `tauri://` URI scheme handler via `.register_uri_scheme_protocol(\"tauri\", ...)`"
    );

    let handler_block = extract_tauri_scheme_protocol_handler_block(&main_rs_src);

    // 2) Ensure the handler applies COOP/COEP headers (required for cross-origin isolation).
    //
    // Prefer checking for the helper call (current implementation), but allow direct insertion of
    // either the header names or the associated constants in case the handler is refactored.
    let has_apply_call = handler_block.contains("apply_cross_origin_isolation_headers(");
    let has_header_strings = handler_block.contains("cross-origin-opener-policy")
        && handler_block.contains("cross-origin-embedder-policy");
    let has_header_constants = handler_block.contains("CROSS_ORIGIN_OPENER_POLICY")
        && handler_block.contains("CROSS_ORIGIN_EMBEDDER_POLICY");

    assert!(
        has_apply_call || has_header_strings || has_header_constants,
        "expected the `tauri://` protocol handler to inject COOP/COEP headers.\n\
         Missing `apply_cross_origin_isolation_headers(...)` (or direct insertion of the COOP/COEP header names).\n\
         This is required for `globalThis.crossOriginIsolated` and SharedArrayBuffer in packaged desktop builds.\n\
         Searched within the `.register_uri_scheme_protocol(\"tauri\", ...)` block:\n\
         {handler_block}"
    );

    // 3) Ensure we preserve Tauri's configured CSP header when available.
    //
    // Dropping the CSP in the custom handler would reduce production parity with the stock Tauri
    // asset protocol and can break worker/module loading.
    assert!(
        handler_block.contains("asset.csp_header"),
        "expected the `tauri://` protocol handler to reference `asset.csp_header` (to preserve the CSP computed by Tauri's AssetResolver)"
    );
    assert!(
        handler_block.contains("Content-Security-Policy"),
        "expected the `tauri://` protocol handler to set the `Content-Security-Policy` header when `asset.csp_header` is present"
    );

    let csp_is_conditional = (handler_block.contains("if let Some") && handler_block.contains("asset.csp_header"))
        || handler_block.contains("match asset.csp_header")
        || handler_block.contains("asset.csp_header.map")
        || handler_block.contains("asset.csp_header.is_some")
        || handler_block.contains("asset.csp_header.is_some_and");

    assert!(
        csp_is_conditional,
        "expected the `Content-Security-Policy` header to be set conditionally based on `asset.csp_header` (so we don't emit an empty/malformed CSP when absent).\n\
         Searched within the `.register_uri_scheme_protocol(\"tauri\", ...)` block:\n\
         {handler_block}"
    );
}

