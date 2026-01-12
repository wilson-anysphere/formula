use std::io::{Cursor, Write};

use encoding_rs::WINDOWS_1251;
use formula_vba::{compress_container, content_normalized_data};

fn push_record(out: &mut Vec<u8>, id: u16, data: &[u8]) {
    out.extend_from_slice(&id.to_le_bytes());
    out.extend_from_slice(&(data.len() as u32).to_le_bytes());
    out.extend_from_slice(data);
}

#[test]
fn content_normalized_data_decodes_module_stream_names_using_project_codepage() {
    // Cyrillic module name: requires a non-UTF-8 project codepage to decode correctly.
    let module_name = "Модуль1";
    let (module_name_bytes, _, _) = WINDOWS_1251.encode(module_name);

    // Module source is plain ASCII so the test isolates stream name decoding.
    let module_code = b"Sub Hello()\r\nEnd Sub\r\n";
    let module_container = compress_container(module_code);

    // Decompressed `VBA/dir` bytes: a single module record group with MODULENAME and
    // MODULESTREAMNAME encoded in Windows-1251 bytes (+ reserved u16 for 0x001A).
    //
    // Include a conflicting PROJECTCODEPAGE (Windows-1252) to ensure `CodePage=` in the
    // `PROJECT` stream takes precedence.
    let dir_decompressed = {
        let mut out = Vec::new();
        push_record(&mut out, 0x0003, &1252u16.to_le_bytes()); // PROJECTCODEPAGE (conflicts)

        push_record(&mut out, 0x0019, module_name_bytes.as_ref()); // MODULENAME

        let mut stream_name = Vec::new();
        stream_name.extend_from_slice(module_name_bytes.as_ref());
        stream_name.extend_from_slice(&0u16.to_le_bytes());
        push_record(&mut out, 0x001A, &stream_name); // MODULESTREAMNAME

        push_record(&mut out, 0x0021, &0u16.to_le_bytes()); // MODULETYPE (standard)
        push_record(&mut out, 0x0031, &0u32.to_le_bytes()); // MODULETEXTOFFSET
        out
    };
    let dir_container = compress_container(&dir_decompressed);

    let cursor = Cursor::new(Vec::new());
    let mut ole = cfb::CompoundFile::create(cursor).expect("create cfb");
    {
        let mut s = ole.create_stream("PROJECT").expect("PROJECT stream");
        s.write_all(b"CodePage=1251\r\n").expect("write PROJECT");
    }
    ole.create_storage("VBA").expect("VBA storage");
    {
        let mut s = ole.create_stream("VBA/dir").expect("dir stream");
        s.write_all(&dir_container).expect("write dir");
    }
    {
        let stream_path = format!("VBA/{module_name}");
        let mut s = ole.create_stream(&stream_path).expect("module stream");
        s.write_all(&module_container).expect("write module");
    }

    let vba_bin = ole.into_inner().into_inner();

    let normalized = content_normalized_data(&vba_bin).expect("ContentNormalizedData");
    assert!(
        normalized.windows(module_code.len()).any(|w| w == module_code),
        "expected normalized data to include module code bytes"
    );
}

#[test]
fn content_normalized_data_ignores_overflowing_project_codepage_line() {
    // Ensure `content_normalized_data()` uses a later valid `CodePage=` value even if an earlier
    // (malformed/overflowing) `CodePage=` line appears first in the PROJECT stream.
    //
    // This is important because if codepage detection fails, the implementation falls back to
    // PROJECTCODEPAGE in `VBA/dir` (which we deliberately set to 1252 here), and module stream lookup
    // for non-ASCII names will fail.
    let module_name = "Модуль1";
    let (module_name_bytes, _, _) = WINDOWS_1251.encode(module_name);

    let module_code = b"Sub Hello()\r\nEnd Sub\r\n";
    let module_container = compress_container(module_code);

    let dir_decompressed = {
        let mut out = Vec::new();
        // Conflicting PROJECTCODEPAGE: should be ignored when PROJECT has a valid CodePage.
        push_record(&mut out, 0x0003, &1252u16.to_le_bytes());

        push_record(&mut out, 0x0019, module_name_bytes.as_ref()); // MODULENAME
        let mut stream_name = Vec::new();
        stream_name.extend_from_slice(module_name_bytes.as_ref());
        stream_name.extend_from_slice(&0u16.to_le_bytes());
        push_record(&mut out, 0x001A, &stream_name); // MODULESTREAMNAME (+reserved)

        push_record(&mut out, 0x0021, &0u16.to_le_bytes()); // MODULETYPE
        push_record(&mut out, 0x0031, &0u32.to_le_bytes()); // MODULETEXTOFFSET
        out
    };
    let dir_container = compress_container(&dir_decompressed);

    let cursor = Cursor::new(Vec::new());
    let mut ole = cfb::CompoundFile::create(cursor).expect("create cfb");
    {
        let mut s = ole.create_stream("PROJECT").expect("PROJECT stream");
        // First line overflows u32, second line is valid.
        s.write_all(b"CodePage=99999999999999999999999999999999\r\nCodePage=1251\r\n")
            .expect("write PROJECT");
    }
    ole.create_storage("VBA").expect("VBA storage");
    {
        let mut s = ole.create_stream("VBA/dir").expect("dir stream");
        s.write_all(&dir_container).expect("write dir");
    }
    {
        let stream_path = format!("VBA/{module_name}");
        let mut s = ole.create_stream(&stream_path).expect("module stream");
        s.write_all(&module_container).expect("write module");
    }
    let vba_bin = ole.into_inner().into_inner();

    let normalized = content_normalized_data(&vba_bin).expect("ContentNormalizedData");
    assert!(
        normalized.windows(module_code.len()).any(|w| w == module_code),
        "expected normalized data to include module code bytes"
    );
}
