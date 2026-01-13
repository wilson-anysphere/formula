use std::io::{Cursor, Write};

use calamine::Reader;
use calamine::Xls;

mod common;

use common::xls_fixture_builder;

/// Ensure that a panic inside `calamine` does not abort the process.
///
/// This test uses a BIFF8 fixture that has historically triggered a `calamine` panic when
/// `NAME` records are split across `CONTINUE` records (unless the stream is sanitized first).
///
/// If the current `calamine` version becomes resilient to this input, we still assert that the
/// importer does not panic; we only require a `CalaminePanic` error when the underlying `calamine`
/// call actually panics.
#[test]
fn import_without_biff_never_panics_on_calamine_panic() {
    let bytes = xls_fixture_builder::build_continued_name_record_fixture_xls();

    // Detect whether `calamine` panics on this fixture when parsing defined names.
    // (Older versions have panicked due to unchecked slicing on continued NAME records.)
    let calamine_panicked = std::panic::catch_unwind(|| {
        let workbook: Xls<_> = Xls::new(Cursor::new(bytes.as_slice())).expect("expected xls to open");
        let _ = workbook.defined_names();
    })
    .is_err();

    let mut tmp = tempfile::NamedTempFile::new().expect("temp file");
    tmp.write_all(&bytes).expect("write xls bytes");

    let result = std::panic::catch_unwind(|| formula_xls::import_xls_path_without_biff(tmp.path()));
    assert!(result.is_ok(), "importer should not panic");

    let result = result.unwrap();
    if calamine_panicked {
        let Err(formula_xls::ImportError::CalaminePanic(message)) = result else {
            panic!("expected CalaminePanic error when calamine panics, got: {result:?}");
        };
        // The importer should attach some context about what it was doing when the panic occurred.
        assert!(
            message.contains("opening `.xls` via calamine")
                || message.contains("reading sheet metadata")
                || message.contains("reading defined names")
                || message.contains("reading cell values")
                || message.contains("reading formulas")
                || message.contains("reading merged cells"),
            "expected contextual CalaminePanic message, got: {message}"
        );
    }
}
