use formula_io::OpenOptions;

const PASSWORD: &str = "hunter2";

#[test]
fn open_options_debug_does_not_leak_password() {
    let opts = OpenOptions {
        password: Some(PASSWORD.to_string()),
    };

    let debug = format!("{opts:?}");
    assert!(!debug.contains(PASSWORD));
}

