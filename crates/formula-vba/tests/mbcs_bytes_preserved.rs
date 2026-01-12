use std::io::{Cursor, Write};

use encoding_rs::WINDOWS_1251;
use formula_vba::{compress_container, project_normalized_data, v3_content_normalized_data};

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
fn v3_content_normalized_data_preserves_non_ascii_mbcs_bytes_verbatim() {
    // Regression test: MS-OVBA "V3ContentNormalizedData" appends many fields as MBCS bytes. A naive
    // implementation might decode to Rust `String` and then append UTF-8 bytes, which corrupts the
    // transcript for non-ASCII projects.
    let project_name = "Привет"; // "hello" in Russian (non-ASCII)
    let (mbcs, _, _) = WINDOWS_1251.encode(project_name);
    let mbcs_bytes = mbcs.as_ref();
    let utf8_bytes = project_name.as_bytes();
    assert_ne!(
        mbcs_bytes, utf8_bytes,
        "test precondition: Windows-1251 bytes must differ from UTF-8 bytes"
    );

    // Keep module source ASCII so the only appearance of `project_name` bytes is via dir records.
    let module_code = b"Sub Foo()\r\nEnd Sub\r\n";
    let module_container = compress_container(module_code);

    let dir_decompressed = {
        let mut out = Vec::new();
        // PROJECTCODEPAGE (u16 LE) = 1251 (Windows-1251).
        push_record(&mut out, 0x0003, &1251u16.to_le_bytes());
        // PROJECTNAME.ProjectName (MBCS bytes in PROJECTCODEPAGE).
        push_record(&mut out, 0x0004, mbcs_bytes);

        // Minimal module record group. Use the same non-ASCII bytes for the module stream name so
        // the transcript must preserve them verbatim.
        push_record(&mut out, 0x0019, mbcs_bytes); // MODULENAME
        let mut stream_name = Vec::new();
        stream_name.extend_from_slice(mbcs_bytes);
        stream_name.extend_from_slice(&0u16.to_le_bytes()); // reserved u16
        push_record(&mut out, 0x001A, &stream_name); // MODULESTREAMNAME
        push_record(&mut out, 0x0021, &0u16.to_le_bytes()); // MODULETYPE (procedural)
        push_record(&mut out, 0x0031, &0u32.to_le_bytes()); // MODULETEXTOFFSET

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
        // The OLE stream name is Unicode; the `dir` stream records provide the MBCS bytes.
        let mut s = ole
            .create_stream(&format!("VBA/{project_name}"))
            .expect("module stream");
        s.write_all(&module_container).expect("write module");
    }

    let vba_bin = ole.into_inner().into_inner();
    let normalized = v3_content_normalized_data(&vba_bin).expect("V3ContentNormalizedData");

    assert!(
        contains_subslice(&normalized, mbcs_bytes),
        "expected V3ContentNormalizedData to contain the original MBCS bytes"
    );
    assert!(
        !contains_subslice(&normalized, utf8_bytes),
        "expected V3ContentNormalizedData to NOT contain UTF-8 bytes for the same string"
    );
}

#[test]
fn project_normalized_data_preserves_non_ascii_mbcs_bytes_verbatim() {
    // Regression test: MS-OVBA ProjectNormalizedData (dir-record based) appends project metadata as
    // MBCS bytes. Accidental UTF-8 transcoding would change the binding digest for non-ASCII
    // projects.
    let project_name = "Проект"; // "project" in Russian (non-ASCII)
    let (mbcs, _, _) = WINDOWS_1251.encode(project_name);
    let mbcs_bytes = mbcs.as_ref();
    let utf8_bytes = project_name.as_bytes();
    assert_ne!(
        mbcs_bytes, utf8_bytes,
        "test precondition: Windows-1251 bytes must differ from UTF-8 bytes"
    );

    // Encode the PROJECT stream as Windows-1251 to match `PROJECTCODEPAGE`. Even though our current
    // ProjectNormalizedData implementation does not incorporate `Name=...` outside of Host Extender
    // Info, including this line makes the test representative of real-world projects and guards
    // against future changes that might parse and append it.
    let project_stream_text = format!("CodePage=1251\r\nName=\"{project_name}\"\r\n");
    let (project_stream_bytes, _, _) = WINDOWS_1251.encode(&project_stream_text);

    // Minimal decompressed `VBA/dir` stream: PROJECTCODEPAGE + PROJECTNAME.
    let dir_decompressed = {
        let mut out = Vec::new();
        push_record(&mut out, 0x0003, &1251u16.to_le_bytes()); // PROJECTCODEPAGE
        push_record(&mut out, 0x0004, mbcs_bytes); // PROJECTNAME.ProjectName (MBCS)
        out
    };
    let dir_container = compress_container(&dir_decompressed);

    let cursor = Cursor::new(Vec::new());
    let mut ole = cfb::CompoundFile::create(cursor).expect("create cfb");
    {
        let mut s = ole.create_stream("PROJECT").expect("PROJECT stream");
        s.write_all(project_stream_bytes.as_ref())
            .expect("write PROJECT");
    }
    ole.create_storage("VBA").expect("VBA storage");
    {
        let mut s = ole.create_stream("VBA/dir").expect("dir stream");
        s.write_all(&dir_container).expect("write dir");
    }

    let vba_bin = ole.into_inner().into_inner();
    let normalized = project_normalized_data(&vba_bin).expect("ProjectNormalizedData");

    assert!(
        contains_subslice(&normalized, mbcs_bytes),
        "expected ProjectNormalizedData to contain the original MBCS bytes"
    );
    assert!(
        !contains_subslice(&normalized, utf8_bytes),
        "expected ProjectNormalizedData to NOT contain UTF-8 bytes for the same string"
    );
}

