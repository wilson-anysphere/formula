use desktop::file_io::read_xlsx_blocking;
use std::ffi::OsString;
use std::path::PathBuf;

struct EnvVarGuard {
    key: &'static str,
    prev: Option<OsString>,
}

impl EnvVarGuard {
    fn set(key: &'static str, value: &str) -> Self {
        let prev = std::env::var_os(key);
        std::env::set_var(key, value);
        Self { key, prev }
    }
}

impl Drop for EnvVarGuard {
    fn drop(&mut self) {
        match &self.prev {
            Some(prev) => std::env::set_var(self.key, prev),
            None => std::env::remove_var(self.key),
        }
    }
}

#[test]
fn xlsx_open_skips_origin_bytes_when_over_limit() {
    let _guard = EnvVarGuard::set("FORMULA_MAX_ORIGIN_XLSX_BYTES", "1");

    let fixture = PathBuf::from(env!("CARGO_MANIFEST_DIR"))
        .join("../../../fixtures/xlsx/basic/basic.xlsx");
    let workbook = read_xlsx_blocking(&fixture).expect("read fixture");

    assert!(
        !workbook.sheets.is_empty(),
        "expected workbook to load successfully"
    );
    assert!(
        workbook.origin_xlsx_bytes.is_none(),
        "expected workbook origin_xlsx_bytes to be dropped when FORMULA_MAX_ORIGIN_XLSX_BYTES is exceeded"
    );
}

