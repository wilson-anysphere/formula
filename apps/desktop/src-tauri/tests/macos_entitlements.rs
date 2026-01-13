use std::collections::BTreeSet;
use std::path::PathBuf;

fn repo_path(relative: &str) -> PathBuf {
    PathBuf::from(env!("CARGO_MANIFEST_DIR")).join(relative)
}

#[test]
fn macos_entitlements_are_minimal_and_allowlisted() {
    let entitlements_path = repo_path("entitlements.plist");

    let value = plist::Value::from_file(&entitlements_path).unwrap_or_else(|err| {
        panic!(
            "failed to parse {} as a plist: {err}",
            entitlements_path.display()
        )
    });

    let dict = value.as_dictionary().unwrap_or_else(|| {
        panic!(
            "{} must be a plist <dict> at the top level",
            entitlements_path.display()
        )
    });

    let actual_keys: BTreeSet<&str> = dict.keys().map(|k| k.as_str()).collect();
    let expected_keys: BTreeSet<&str> = [
        "com.apple.security.network.client",
        "com.apple.security.cs.allow-jit",
        "com.apple.security.cs.allow-unsigned-executable-memory",
    ]
    .into_iter()
    .collect();

    assert_eq!(
        actual_keys, expected_keys,
        "unexpected macOS entitlements in {} (keep the signed entitlement surface minimal)",
        entitlements_path.display()
    );

    for (key, value) in dict {
        assert_eq!(
            value.as_boolean(),
            Some(true),
            "entitlement {key} must be set to boolean true in {}",
            entitlements_path.display()
        );
    }
}
