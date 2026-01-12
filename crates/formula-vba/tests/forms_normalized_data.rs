use std::io::{Cursor, Write};

use formula_vba::forms_normalized_data;

#[test]
fn forms_normalized_data_pads_stream_to_1023_byte_blocks() {
    let cursor = Cursor::new(Vec::new());
    let mut ole = cfb::CompoundFile::create(cursor).expect("create compound file");

    // Root-level "designer" storage with a single short stream payload.
    ole.create_storage("UserForm1")
        .expect("create designer storage");
    {
        let mut s = ole
            .create_stream("UserForm1/Payload")
            .expect("create stream");
        s.write_all(b"ABC").expect("write stream bytes");
    }

    let vba_project_bin = ole.into_inner().into_inner();
    let normalized = forms_normalized_data(&vba_project_bin).expect("compute FormsNormalizedData");

    // MS-OVBA pads the final block to 1023 bytes with zeros.
    let mut expected = Vec::new();
    expected.extend_from_slice(b"ABC");
    expected.extend(std::iter::repeat(0u8).take(1020));

    assert_eq!(normalized.len(), 1023);
    assert_eq!(normalized, expected);
}

#[test]
fn forms_normalized_data_traverses_nested_storages_in_deterministic_order() {
    let cursor = Cursor::new(Vec::new());
    let mut ole = cfb::CompoundFile::create(cursor).expect("create compound file");

    ole.create_storage("UserForm1")
        .expect("create designer storage");
    ole.create_storage("UserForm1/Child")
        .expect("create nested storage");

    // Write the sibling stream first, then the nested one; the normalization implementation should
    // still process streams in deterministic order.
    {
        let mut s = ole
            .create_stream("UserForm1/Y")
            .expect("create sibling stream");
        s.write_all(b"Y").expect("write Y");
    }
    {
        let mut s = ole
            .create_stream("UserForm1/Child/X")
            .expect("create nested stream");
        s.write_all(b"X").expect("write X");
    }

    let vba_project_bin = ole.into_inner().into_inner();
    let normalized = forms_normalized_data(&vba_project_bin).expect("compute FormsNormalizedData");

    // The library defines traversal order as lexicographic by full OLE path. That yields:
    // - `UserForm1/Child/X`
    // - `UserForm1/Y`
    let mut expected = Vec::new();
    expected.extend_from_slice(b"X");
    expected.extend(std::iter::repeat(0u8).take(1022));
    expected.extend_from_slice(b"Y");
    expected.extend(std::iter::repeat(0u8).take(1022));

    assert_eq!(normalized.len(), 1023 * 2);
    assert_eq!(normalized, expected);
}

