use std::io::{Cursor, Write};

use formula_vba::{compress_container, v3_content_normalized_data};

fn push_record(out: &mut Vec<u8>, id: u16, data: &[u8]) {
    out.extend_from_slice(&id.to_le_bytes());
    out.extend_from_slice(&(data.len() as u32).to_le_bytes());
    out.extend_from_slice(data);
}

fn contains_subslice(haystack: &[u8], needle: &[u8]) -> bool {
    if needle.is_empty() {
        return true;
    }
    haystack.windows(needle.len()).any(|w| w == needle)
}

#[test]
fn v3_content_normalized_data_includes_projectmodules_projectcookie_headers_and_dir_trailer() {
    // Regression test for MS-OVBA §2.4.2.5 V3ContentNormalizedData:
    //
    // - PROJECTMODULES: include Id+Size bytes, but *exclude* Count (u16).
    // - PROJECTCOOKIE: include Id+Size bytes, but *exclude* Cookie (u16).
    // - dir trailer: output must end with Terminator (0x0010) + Reserved (0x00000000).
    //
    // This test intentionally builds a minimal decompressed `VBA/dir` stream using the same
    // record framing (`Id(u16) || Size(u32) || Data`) that `v3_content_normalized_data()` parses.
    let module_code = b"Sub Foo()\r\nEnd Sub\r\n";
    let module_container = compress_container(module_code);

    let dir_decompressed = {
        let mut out = Vec::new();

        // PROJECTMODULES (0x000F): Size must be 2 (Count only).
        push_record(&mut out, 0x000F, &1u16.to_le_bytes());

        // PROJECTCOOKIE (0x0013): Size must be 2 (Cookie).
        let cookie_value = 0xBEEFu16;
        push_record(&mut out, 0x0013, &cookie_value.to_le_bytes());

        // Minimal module record group so we exercise the "flush pending module before dir
        // terminator" behavior.
        push_record(&mut out, 0x0019, b"Module1"); // MODULENAME
        let mut stream_name = Vec::new();
        stream_name.extend_from_slice(b"Module1");
        stream_name.extend_from_slice(&0u16.to_le_bytes()); // reserved
        push_record(&mut out, 0x001A, &stream_name); // MODULESTREAMNAME
        push_record(&mut out, 0x0031, &0u32.to_le_bytes()); // MODULETEXTOFFSET

        // dir Terminator/Reserved trailer (treated as a record with Size=0).
        push_record(&mut out, 0x0010, &[]);

        out
    };
    let dir_container = compress_container(&dir_decompressed);

    let cursor = Cursor::new(Vec::new());
    let mut ole = cfb::CompoundFile::create(cursor).expect("create cfb");
    ole.create_storage("VBA").expect("VBA storage");
    {
        let mut s = ole.create_stream("VBA/dir").expect("dir stream");
        s.write_all(&dir_container).expect("write dir bytes");
    }
    {
        let mut s = ole
            .create_stream("VBA/Module1")
            .expect("module stream");
        s.write_all(&module_container)
            .expect("write module bytes");
    }

    let vba_project_bin = ole.into_inner().into_inner();
    let normalized = v3_content_normalized_data(&vba_project_bin).expect("V3ContentNormalizedData");

    let projectmodules_header = [0x0F, 0x00, 0x02, 0x00, 0x00, 0x00];
    let projectcookie_header = [0x13, 0x00, 0x02, 0x00, 0x00, 0x00];
    let dir_trailer = [0x10, 0x00, 0x00, 0x00, 0x00, 0x00];

    assert!(
        contains_subslice(&normalized, &projectmodules_header),
        "expected output to contain PROJECTMODULES.Id+Size bytes (0x000F, 0x00000002 LE)"
    );
    assert!(
        contains_subslice(&normalized, &projectcookie_header),
        "expected output to contain PROJECTCOOKIE.Id+Size bytes (0x0013, 0x00000002 LE)"
    );
    assert!(
        normalized.ends_with(&dir_trailer),
        "expected output to end with dir Terminator/Reserved trailer (0x0010, 0x00000000 LE)"
    );

    // Ensure Count/Cookie values are not accidentally included (as contiguous header+value bytes).
    let projectmodules_header_plus_count = [0x0F, 0x00, 0x02, 0x00, 0x00, 0x00, 0x01, 0x00];
    let projectcookie_header_plus_cookie = [0x13, 0x00, 0x02, 0x00, 0x00, 0x00, 0xEF, 0xBE];
    assert!(
        !contains_subslice(&normalized, &projectmodules_header_plus_count),
        "did not expect PROJECTMODULES.Count bytes to be included in output"
    );
    assert!(
        !contains_subslice(&normalized, &projectcookie_header_plus_cookie),
        "did not expect PROJECTCOOKIE.Cookie bytes to be included in output"
    );
}

