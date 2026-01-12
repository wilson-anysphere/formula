use std::io::{Cursor, Write};

use formula_vba::{compress_container, forms_normalized_data};

fn push_record(out: &mut Vec<u8>, id: u16, data: &[u8]) {
    out.extend_from_slice(&id.to_le_bytes());
    out.extend_from_slice(&(data.len() as u32).to_le_bytes());
    out.extend_from_slice(data);
}

fn build_dir_stream_with_designer_module(module_name: &str) -> Vec<u8> {
    let dir_decompressed = {
        let mut out = Vec::new();
        // PROJECTCODEPAGE (u16 LE)
        push_record(&mut out, 0x0003, &1252u16.to_le_bytes());
        // MODULENAME
        push_record(&mut out, 0x0019, module_name.as_bytes());
        // MODULESTREAMNAME + reserved u16
        let mut stream_name = Vec::new();
        stream_name.extend_from_slice(module_name.as_bytes());
        stream_name.extend_from_slice(&0u16.to_le_bytes());
        push_record(&mut out, 0x001A, &stream_name);
        // MODULETYPE (UserForm)
        push_record(&mut out, 0x0021, &0x0003u16.to_le_bytes());
        // MODULETEXTOFFSET
        push_record(&mut out, 0x0031, &0u32.to_le_bytes());
        out
    };
    compress_container(&dir_decompressed)
}

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

    // Minimal PROJECT and `VBA/dir` so the implementation can discover `UserForm1` as a designer.
    {
        let mut s = ole.create_stream("PROJECT").expect("PROJECT stream");
        s.write_all(b"BaseClass=UserForm1\r\n")
            .expect("write PROJECT");
    }
    ole.create_storage("VBA").expect("VBA storage");
    {
        let dir_container = build_dir_stream_with_designer_module("UserForm1");
        let mut s = ole.create_stream("VBA/dir").expect("dir stream");
        s.write_all(&dir_container).expect("write dir");
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

    // Minimal PROJECT and `VBA/dir` so the implementation can discover `UserForm1` as a designer.
    {
        let mut s = ole.create_stream("PROJECT").expect("PROJECT stream");
        s.write_all(b"BaseClass=UserForm1\r\n")
            .expect("write PROJECT");
    }
    ole.create_storage("VBA").expect("VBA storage");
    {
        let dir_container = build_dir_stream_with_designer_module("UserForm1");
        let mut s = ole.create_stream("VBA/dir").expect("dir stream");
        s.write_all(&dir_container).expect("write dir");
    }

    let vba_project_bin = ole.into_inner().into_inner();
    let normalized = forms_normalized_data(&vba_project_bin).expect("compute FormsNormalizedData");

    // The library defines traversal order as case-insensitive name order per-storage. For this
    // fixture that yields:
    // - recurse into `Child` storage first (stream `X`)
    // - then sibling stream `Y`
    let mut expected = Vec::new();
    expected.extend_from_slice(b"X");
    expected.extend(std::iter::repeat(0u8).take(1022));
    expected.extend_from_slice(b"Y");
    expected.extend(std::iter::repeat(0u8).take(1022));

    assert_eq!(normalized.len(), 1023 * 2);
    assert_eq!(normalized, expected);
}
