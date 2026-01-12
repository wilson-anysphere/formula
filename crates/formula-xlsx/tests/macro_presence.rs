use std::io::{Cursor, Write};

use formula_model::{Alignment, Range};
use formula_xlsx::{load_from_bytes, MacroPresence, XlsxPackage};

fn build_zip(files: &[(&str, &[u8])]) -> Vec<u8> {
    let cursor = Cursor::new(Vec::new());
    let mut zip = zip::ZipWriter::new(cursor);
    let options = zip::write::FileOptions::<()>::default()
        .compression_method(zip::CompressionMethod::Deflated);

    for (name, bytes) in files {
        zip.start_file(*name, options).unwrap();
        zip.write_all(bytes).unwrap();
    }

    zip.finish().unwrap().into_inner()
}

fn load_minimal_doc_with_parts(parts: &[(&str, &[u8])]) -> formula_xlsx::XlsxDocument {
    let base = formula_xlsx::write_minimal_xlsx(&[] as &[Range], &[] as &[Alignment]).unwrap();
    let mut pkg = XlsxPackage::from_bytes(&base).unwrap();
    for (name, bytes) in parts {
        pkg.set_part(*name, bytes.to_vec());
    }
    let bytes = pkg.write_to_bytes().unwrap();
    load_from_bytes(&bytes).unwrap()
}

fn assert_presence(
    presence: MacroPresence,
    has_vba: bool,
    has_xlm_macrosheets: bool,
    has_dialog_sheets: bool,
) {
    assert_eq!(
        presence,
        MacroPresence {
            has_vba,
            has_xlm_macrosheets,
            has_dialog_sheets,
        }
    );
    assert_eq!(presence.any(), has_vba || has_xlm_macrosheets || has_dialog_sheets);
}

#[test]
fn macro_presence_vba_only() {
    let bytes = build_zip(&[("xl/vbaProject.bin", b"fake-vba-project")]);
    let pkg = XlsxPackage::from_bytes(&bytes).unwrap();
    assert_presence(pkg.macro_presence(), true, false, false);

    let doc = load_minimal_doc_with_parts(&[("xl/vbaProject.bin", b"fake-vba-project")]);
    assert_presence(doc.macro_presence(), true, false, false);
}

#[test]
fn macro_presence_macrosheet_only() {
    let bytes = build_zip(&[("xl/macrosheets/sheet1.xml", b"<macrosheet/>")]);
    let pkg = XlsxPackage::from_bytes(&bytes).unwrap();
    assert_presence(pkg.macro_presence(), false, true, false);

    let doc = load_minimal_doc_with_parts(&[("xl/macrosheets/sheet1.xml", b"<macrosheet/>")]);
    assert_presence(doc.macro_presence(), false, true, false);
}

#[test]
fn macro_presence_dialogsheet_only() {
    let bytes = build_zip(&[("xl/dialogsheets/sheet1.xml", b"<dialogsheet/>")]);
    let pkg = XlsxPackage::from_bytes(&bytes).unwrap();
    assert_presence(pkg.macro_presence(), false, false, true);

    let doc = load_minimal_doc_with_parts(&[("xl/dialogsheets/sheet1.xml", b"<dialogsheet/>")]);
    assert_presence(doc.macro_presence(), false, false, true);
}

#[test]
fn macro_presence_none() {
    let bytes = build_zip(&[("docProps/app.xml", b"<Properties/>")]);
    let pkg = XlsxPackage::from_bytes(&bytes).unwrap();
    assert_presence(pkg.macro_presence(), false, false, false);

    let doc = load_minimal_doc_with_parts(&[]);
    assert_presence(doc.macro_presence(), false, false, false);
}

