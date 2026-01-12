use std::io::{Cursor, Write};

use encoding_rs::WINDOWS_1251;
use formula_vba::{compress_container, project_normalized_data_v3, v3_content_normalized_data};

fn push_record(out: &mut Vec<u8>, id: u16, data: &[u8]) {
    out.extend_from_slice(&id.to_le_bytes());
    out.extend_from_slice(&(data.len() as u32).to_le_bytes());
    out.extend_from_slice(data);
}

fn find_subslice(haystack: &[u8], needle: &[u8]) -> Option<usize> {
    if needle.is_empty() {
        return Some(0);
    }
    haystack
        .windows(needle.len())
        .position(|window| window == needle)
}

#[test]
fn v3_content_normalized_data_decodes_module_stream_name_using_project_codepage() {
    let module_stream_name = "Модуль1";
    let (stream_name_bytes, _, _) = WINDOWS_1251.encode(module_stream_name);

    let module_code = b"Sub Hello()\r\nEnd Sub\r\n";
    let module_container = compress_container(module_code);

    let dir_decompressed = {
        let mut out = Vec::new();
        // PROJECTCODEPAGE = 1251
        push_record(&mut out, 0x0003, &1251u16.to_le_bytes());

        // MODULENAME is ASCII (module identifier), but MODULESTREAMNAME contains the non-ASCII
        // stream name that we must decode using PROJECTCODEPAGE.
        push_record(&mut out, 0x0019, b"Module1");

        let mut stream_name_record = Vec::new();
        stream_name_record.extend_from_slice(stream_name_bytes.as_ref());
        stream_name_record.extend_from_slice(&0u16.to_le_bytes()); // reserved u16
        push_record(&mut out, 0x001A, &stream_name_record);

        push_record(&mut out, 0x0021, &0u16.to_le_bytes()); // MODULETYPE (standard)
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
        let mut s = ole
            .create_stream(&format!("VBA/{module_stream_name}"))
            .expect("module stream");
        s.write_all(&module_container).expect("write module");
    }

    let vba_bin = ole.into_inner().into_inner();
    let normalized = v3_content_normalized_data(&vba_bin).expect("V3ContentNormalizedData");

    // MS-OVBA §2.4.2.5: V3ContentNormalizedData uses `\n` line endings and appends the module name
    // plus `\n` (HashModuleNameFlag) after the module's normalized source bytes.
    let expected_suffix = b"Sub Hello()\nEnd Sub\n\nModule1\n";
    assert!(
        normalized.ends_with(expected_suffix),
        "expected V3ContentNormalizedData to end with the v3-normalized module transcript"
    );
}

#[test]
fn project_normalized_data_v3_decodes_baseclass_using_project_codepage() {
    let module_name = "Форма1";
    let designer_stream_bytes = b"DESIGNER-STORAGE-BYTES";

    let (module_name_bytes, _, _) = WINDOWS_1251.encode(module_name);

    // Encode the PROJECT stream as Windows-1251, including a non-ASCII BaseClass= value.
    let project_stream_text =
        format!("CodePage=1251\r\nName=\"VBAProject\"\r\nBaseClass={module_name}\r\n");
    let (project_stream_bytes, _, _) = WINDOWS_1251.encode(&project_stream_text);

    // Include a conflicting PROJECTCODEPAGE record to ensure we prefer the PROJECT stream's
    // `CodePage=` line.
    let dir_decompressed = {
        let mut out = Vec::new();
        push_record(&mut out, 0x0003, &1252u16.to_le_bytes());

        // Module record group for the designer module referenced by BaseClass=...
        push_record(&mut out, 0x0019, module_name_bytes.as_ref()); // MODULENAME
        let mut stream_name_record = Vec::new();
        stream_name_record.extend_from_slice(module_name_bytes.as_ref());
        stream_name_record.extend_from_slice(&0u16.to_le_bytes()); // reserved u16
        push_record(&mut out, 0x001A, &stream_name_record); // MODULESTREAMNAME
        push_record(&mut out, 0x0021, &0x0003u16.to_le_bytes()); // MODULETYPE (UserForm)
        push_record(&mut out, 0x0031, &0u32.to_le_bytes()); // MODULETEXTOFFSET
        out
    };
    let dir_container = compress_container(&dir_decompressed);

    // Module source (arbitrary, but must exist so V3ContentNormalizedData can read the module stream).
    let module_code = b"Sub Hello()\r\nEnd Sub\r\n";
    let module_container = compress_container(module_code);

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
    {
        let mut s = ole
            .create_stream(&format!("VBA/{module_name}"))
            .expect("module stream");
        s.write_all(&module_container).expect("write module");
    }

    // Root-level designer storage referenced by BaseClass=... (storage name is MODULESTREAMNAME).
    ole.create_storage(module_name).expect("designer storage");
    {
        let mut s = ole
            .create_stream(&format!("{module_name}/Payload"))
            .expect("designer stream");
        s.write_all(designer_stream_bytes)
            .expect("write designer bytes");
    }

    let vba_bin = ole.into_inner().into_inner();
    let normalized = project_normalized_data_v3(&vba_bin).expect("ProjectNormalizedDataV3");

    assert!(
        find_subslice(&normalized, designer_stream_bytes).is_some(),
        "expected ProjectNormalizedDataV3 to include designer storage stream bytes"
    );
}
