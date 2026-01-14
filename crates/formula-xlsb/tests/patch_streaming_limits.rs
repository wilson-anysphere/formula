use std::cell::Cell as ThreadCell;
use std::io::{self, Cursor, Read};
use std::sync::Mutex;

use formula_xlsb::{biff12_varint, patch_sheet_bin_streaming, CellEdit, CellValue, Error};

static ENV_LOCK: Mutex<()> = Mutex::new(());

struct EnvVarGuard {
    key: &'static str,
    old: Option<String>,
}

impl EnvVarGuard {
    fn set(key: &'static str, value: &str) -> Self {
        let old = std::env::var(key).ok();
        std::env::set_var(key, value);
        Self { key, old }
    }
}

impl Drop for EnvVarGuard {
    fn drop(&mut self) {
        match self.old.take() {
            Some(v) => std::env::set_var(self.key, v),
            None => std::env::remove_var(self.key),
        }
    }
}

struct PanicAfterReadThreshold {
    inner: Cursor<Vec<u8>>,
    bytes_read: std::rc::Rc<ThreadCell<usize>>,
    max_bytes: usize,
}

impl Read for PanicAfterReadThreshold {
    fn read(&mut self, buf: &mut [u8]) -> io::Result<usize> {
        let n = self.inner.read(buf)?;
        let total = self.bytes_read.get().saturating_add(n);
        self.bytes_read.set(total);
        if total > self.max_bytes {
            panic!(
                "unexpectedly read too much of the worksheet stream: {total} bytes (limit {})",
                self.max_bytes
            );
        }
        Ok(n)
    }
}

#[test]
fn streaming_patcher_rejects_oversized_fallback_sheet_stream_without_reading_whole_input() {
    let _lock = ENV_LOCK.lock().unwrap();
    let _max_part = EnvVarGuard::set("FORMULA_XLSB_MAX_ZIP_PART_BYTES", "20");

    // Craft a minimal BIFF stream that triggers the streaming patcher's in-memory fallback:
    // it sees BrtSheetData before BrtWsDim (DIMENSION), then attempts to buffer the whole stream.
    const SHEETDATA: u32 = 0x0091;
    let mut bytes = Vec::new();
    biff12_varint::write_record_id(&mut bytes, SHEETDATA).expect("write SHEETDATA id");
    biff12_varint::write_record_len(&mut bytes, 0).expect("write SHEETDATA len");
    bytes.extend_from_slice(&[0xA5u8; 128]);

    let edits = [CellEdit {
        row: 0,
        col: 0,
        new_value: CellValue::Number(1.0),
        new_style: None,
        clear_formula: false,
        new_formula: None,
        new_rgcb: None,
        new_formula_flags: None,
        shared_string_index: None,
    }];

    let bytes_read = std::rc::Rc::new(ThreadCell::new(0usize));
    let reader = PanicAfterReadThreshold {
        inner: Cursor::new(bytes),
        bytes_read: bytes_read.clone(),
        // If the patcher were to buffer the entire (128+ header) stream, it would exceed this.
        // With the size guard, it should stop at `max+1` bytes.
        max_bytes: 40,
    };

    let mut out = Vec::new();
    let err = patch_sheet_bin_streaming(reader, &mut out, &edits)
        .err()
        .expect("expected oversized stream error");

    match err {
        Error::PartTooLarge { part, size, max } => {
            assert!(part.contains("worksheet"), "unexpected part name: {part}");
            assert_eq!(max, 20);
            assert!(size > max);
        }
        other => panic!("unexpected error: {other:?}"),
    }

    assert!(
        bytes_read.get() <= 40,
        "expected the patcher to stop early, read {} bytes",
        bytes_read.get()
    );
}

