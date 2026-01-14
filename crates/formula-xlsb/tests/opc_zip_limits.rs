use formula_xlsb::{Error, OpenOptions, XlsbWorkbook};
use std::io::{Cursor, Write};
use std::sync::Mutex;
use tempfile::NamedTempFile;

mod fixture_builder;

use fixture_builder::XlsbFixtureBuilder;

static ENV_LOCK: Mutex<()> = Mutex::new(());

fn write_temp_xlsb(bytes: &[u8]) -> NamedTempFile {
    let mut file = tempfile::Builder::new()
        .prefix("formula_xlsb_zip_limits_")
        .suffix(".xlsb")
        .tempfile()
        .expect("create temp xlsb");
    file.write_all(bytes).expect("write temp xlsb");
    file.flush().expect("flush temp xlsb");
    file
}

struct EnvVarGuard {
    key: &'static str,
    prev: Option<String>,
}

impl EnvVarGuard {
    fn set(key: &'static str, value: &str) -> Self {
        let prev = std::env::var(key).ok();
        std::env::set_var(key, value);
        Self { key, prev }
    }
}

impl Drop for EnvVarGuard {
    fn drop(&mut self) {
        match self.prev.take() {
            Some(v) => std::env::set_var(self.key, v),
            None => std::env::remove_var(self.key),
        }
    }
}

#[test]
fn open_rejects_oversized_zip_part() {
    let _lock = ENV_LOCK.lock().expect("lock env");
    let _max_part = EnvVarGuard::set("FORMULA_XLSB_MAX_ZIP_PART_BYTES", "1024");
    let _max_total = EnvVarGuard::set("FORMULA_XLSB_MAX_PRESERVED_TOTAL_BYTES", "10000000");

    let mut builder = XlsbFixtureBuilder::new();
    builder.add_extra_zip_part("xl/unknown.bin", vec![0u8; 1025]);
    let bytes = builder.build_bytes();
    let tmp = write_temp_xlsb(&bytes);

    let options = OpenOptions {
        preserve_unknown_parts: true,
        preserve_parsed_parts: false,
        preserve_worksheets: false,
        decode_formulas: true,
    };
    let err = XlsbWorkbook::open_with_options(tmp.path(), options)
        .expect_err("expected open to fail due to oversized part");

    match err {
        Error::PartTooLarge { part, size, max } => {
            assert_eq!(part, "xl/unknown.bin");
            assert_eq!(size, 1025);
            assert_eq!(max, 1024);
        }
        other => panic!("expected PartTooLarge, got {other:?}"),
    }
}

#[test]
fn open_rejects_preserved_parts_total_over_budget() {
    let _lock = ENV_LOCK.lock().expect("lock env");
    let _max_part = EnvVarGuard::set("FORMULA_XLSB_MAX_ZIP_PART_BYTES", "4096");

    let mut builder = XlsbFixtureBuilder::new();
    builder.add_extra_zip_part("xl/unknown1.bin", vec![0u8; 6]);
    builder.add_extra_zip_part("xl/unknown2.bin", vec![0u8; 6]);
    let bytes = builder.build_bytes();

    // Compute how many bytes open will preserve before it starts copying unknown parts.
    let mut zip =
        zip::ZipArchive::new(Cursor::new(&bytes)).expect("open in-memory xlsb zip for sizing");
    let mut known_total = 0u64;
    for name in ["[Content_Types].xml", "_rels/.rels", "xl/_rels/workbook.bin.rels"] {
        let entry = zip.by_name(name).expect("expected part to exist");
        known_total += entry.size();
    }

    // Allow one unknown part (6 bytes) to be preserved, but not both (12 bytes).
    let max_total = known_total + 10;
    let max_total_str = max_total.to_string();
    let _max_total = EnvVarGuard::set("FORMULA_XLSB_MAX_PRESERVED_TOTAL_BYTES", &max_total_str);

    let tmp = write_temp_xlsb(&bytes);
    let options = OpenOptions {
        preserve_unknown_parts: true,
        preserve_parsed_parts: false,
        preserve_worksheets: false,
        decode_formulas: true,
    };
    let err = XlsbWorkbook::open_with_options(tmp.path(), options)
        .expect_err("expected open to fail due to preserved parts size budget");

    match err {
        Error::PreservedPartsTooLarge { total, max } => {
            assert_eq!(max, max_total);
            assert_eq!(total, known_total + 12);
        }
        other => panic!("expected PreservedPartsTooLarge, got {other:?}"),
    }
}

