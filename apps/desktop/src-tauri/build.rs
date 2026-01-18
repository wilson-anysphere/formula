fn main() {
    // Only needed when building the desktop binary target. Keeping it in place
    // matches the standard Tauri layout.
    #[cfg(feature = "desktop")]
    {
        // Allow building the Rust desktop binary without running the (expensive) frontend build.
        //
        // `tauri_build::build()` expects `frontendDist` from `tauri.conf.json` to exist at build
        // time so it can bundle the assets. For CI shell-startup benchmarks we don't need the
        // real app frontend, but we still want the binary to compile cleanly.
        //
        // If the dist dir is missing, create a tiny placeholder `index.html`. The runtime
        // `--startup-bench` mode overrides `tauri://` responses anyway, but this keeps the
        // build step lightweight and deterministic.
        let Ok(manifest_dir) = std::env::var("CARGO_MANIFEST_DIR") else {
            println!("cargo:warning=missing CARGO_MANIFEST_DIR; skipping frontendDist placeholder creation");
            tauri_build::build();
            return;
        };
        let manifest_dir = std::path::PathBuf::from(manifest_dir);
        let dist_dir = manifest_dir.join("../dist");
        let index_html = dist_dir.join("index.html");
        if !index_html.exists() {
            if let Err(err) = std::fs::create_dir_all(&dist_dir) {
                println!(
                    "cargo:warning=failed to create placeholder frontendDist directory ({:?}): {err}",
                    dist_dir
                );
            }
            let placeholder = r#"<!doctype html>
<meta charset="utf-8" />
<title>Formula</title>
<body>Formula desktop frontend assets are not bundled in this build.</body>
"#;
            if let Err(err) = std::fs::write(&index_html, placeholder) {
                println!(
                    "cargo:warning=failed to write placeholder frontendDist index.html ({:?}): {err}",
                    index_html
                );
            }
        }

        tauri_build::build();
    }
}
