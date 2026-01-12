use std::io::{self, Cursor, Write};

use formula_fs::atomic_write;
use formula_io::{save_workbook, Workbook};

#[test]
fn save_workbook_replaces_existing_file_atomically() {
    let mut model = formula_model::Workbook::new();
    model.add_sheet("Sheet1").expect("add sheet");

    let mut cursor = Cursor::new(Vec::new());
    formula_xlsx::write_workbook_to_writer(&model, &mut cursor).expect("write workbook");
    let bytes = cursor.into_inner();
    let pkg = formula_xlsx::XlsxPackage::from_bytes(&bytes).expect("parse generated package");
    let expected = pkg.write_to_bytes().expect("write package bytes");

    let workbook = Workbook::Xlsx(pkg);

    let dir = tempfile::tempdir().expect("temp dir");
    let out_path = dir.path().join("existing.xlsx");
    std::fs::write(&out_path, b"old-bytes").expect("seed existing file");

    save_workbook(&workbook, &out_path).expect("save workbook");
    let written = std::fs::read(&out_path).expect("read written bytes");
    assert_eq!(written, expected);
}

#[test]
fn save_workbook_creates_parent_directories() {
    let tmp = tempfile::tempdir().expect("temp dir");
    let out_path = tmp.path().join("nested/dir/workbook.xlsx");

    assert!(
        !out_path.parent().unwrap().exists(),
        "test precondition: parent dir should not exist"
    );

    let mut model = formula_model::Workbook::new();
    model.add_sheet("Sheet1".to_string()).expect("add sheet");

    let workbook = Workbook::Model(model);
    save_workbook(&workbook, &out_path).expect("save workbook");

    assert!(out_path.exists(), "expected workbook file to be created");
    assert!(
        out_path.parent().unwrap().exists(),
        "expected parent directories to be created"
    );
}

#[test]
fn atomic_write_does_not_clobber_existing_file_on_write_error() {
    let tmp = tempfile::tempdir().expect("temp dir");
    let dest = tmp.path().join("existing.bin");

    let sentinel = b"sentinel-bytes";
    std::fs::write(&dest, sentinel).expect("write sentinel dest file");

    let err = atomic_write(&dest, |file| {
        file.write_all(b"partial").expect("write to temp file");
        Err::<(), _>(io::Error::new(io::ErrorKind::Other, "simulated write failure"))
    })
    .expect_err("expected atomic_write to return error");

    // The destination file must remain untouched.
    let got = std::fs::read(&dest).expect("read dest");
    assert_eq!(got, sentinel, "dest file should not be clobbered: {err}");

    // Temp file should be cleaned up.
    let entries: Vec<_> = std::fs::read_dir(tmp.path())
        .expect("read_dir")
        .collect::<Result<Vec<_>, _>>()
        .expect("list dir");
    let names: Vec<_> = entries
        .iter()
        .map(|e| e.path())
        .filter(|p| p.is_file())
        .collect();
    assert_eq!(
        names,
        vec![dest.clone()],
        "expected only the destination file to remain (no temp files)"
    );
}
