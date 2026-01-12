use std::io::{Cursor, Write};

use formula_vba::{compress_container, v3_content_normalized_data};

fn push_record(out: &mut Vec<u8>, id: u16, data: &[u8]) {
    out.extend_from_slice(&id.to_le_bytes());
    out.extend_from_slice(&(data.len() as u32).to_le_bytes());
    out.extend_from_slice(data);
}

fn utf16le_bytes(s: &str) -> Vec<u8> {
    s.encode_utf16()
        .flat_map(|u| u.to_le_bytes())
        .collect::<Vec<u8>>()
}

#[test]
fn v3_content_normalized_data_uses_modulestreamname_unicode_record_id_0048_for_stream_lookup() {
    // Some producers store a Unicode module stream name in a separate record with id 0x0048
    // immediately after MODULESTREAMNAME (0x001A), even though 0x0048 is canonically used for
    // MODULEDOCSTRINGUNICODE.
    //
    // Ensure `v3_content_normalized_data()` uses this record (when present) for OLE stream lookup.
    let module_stream_name_unicode = "ユニコード名";
    let module_stream_name_unicode_utf16 = utf16le_bytes(module_stream_name_unicode);

    let module_code = b"Sub Hello()\r\nEnd Sub\r\n";
    let module_container = compress_container(module_code);

    let dir_decompressed = {
        let mut out = Vec::new();

        // Module record group.
        push_record(&mut out, 0x0019, b"Module1"); // MODULENAME
        // MODULESTREAMNAME (ANSI/MBCS): deliberately wrong.
        push_record(&mut out, 0x001A, b"Wrong");
        // Nonstandard Unicode stream-name record id (0x0048).
        push_record(&mut out, 0x0048, &module_stream_name_unicode_utf16);
        // MODULETYPE + MODULETEXTOFFSET.
        push_record(&mut out, 0x0021, &0u16.to_le_bytes());
        push_record(&mut out, 0x0031, &0u32.to_le_bytes());

        out
    };
    let dir_container = compress_container(&dir_decompressed);

    let cursor = Cursor::new(Vec::new());
    let mut ole = cfb::CompoundFile::create(cursor).expect("create cfb");
    ole.create_storage("VBA").expect("VBA storage");
    {
        let mut s = ole.create_stream("VBA/dir").expect("dir stream");
        s.write_all(&dir_container).expect("write dir");
    }
    {
        let mut s = ole
            .create_stream(format!("VBA/{module_stream_name_unicode}"))
            .expect("module stream");
        s.write_all(&module_container).expect("write module");
    }
    let vba_bin = ole.into_inner().into_inner();

    let normalized = v3_content_normalized_data(&vba_bin).expect("V3ContentNormalizedData");

    let mut expected = Vec::new();
    // TypeRecord.Id (0x0021) + Reserved (0)
    expected.extend_from_slice(&0x0021u16.to_le_bytes());
    expected.extend_from_slice(&0u16.to_le_bytes());
    // LF-normalized source + trailing module name bytes + LF.
    expected.extend_from_slice(b"Sub Hello()\nEnd Sub\n\nModule1\n");

    assert_eq!(normalized, expected);
}

