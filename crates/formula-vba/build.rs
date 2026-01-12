use std::fs;
use std::path::PathBuf;

fn main() {
    // Used by `examples/dump_dir_records.rs` so it can compile both before and after
    // `project_normalized_data_v3` lands in the public API.
    //
    // This is intentionally best-effort and string-based: it's a developer tool hook,
    // not a correctness-critical build step.
    println!("cargo:rustc-check-cfg=cfg(formula_vba_has_project_normalized_data_v3)");

    let manifest_dir = PathBuf::from(std::env::var_os("CARGO_MANIFEST_DIR").unwrap());
    let src_dir = manifest_dir.join("src");

    let mut found = false;
    if let Ok(entries) = fs::read_dir(&src_dir) {
        for entry in entries.flatten() {
            let path = entry.path();
            if path.extension().and_then(|s| s.to_str()) != Some("rs") {
                continue;
            }

            println!("cargo:rerun-if-changed={}", path.display());

            if let Ok(text) = fs::read_to_string(&path) {
                if text.contains("project_normalized_data_v3") {
                    found = true;
                }
            }
        }
    }

    if found {
        println!("cargo:rustc-cfg=formula_vba_has_project_normalized_data_v3");
    }
}

