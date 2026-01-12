use std::io::{Cursor, Write};

use formula_vba::{compress_container, VBAProject};

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
fn vba_project_parse_uses_modulestreamname_unicode_tail_for_stream_lookup() {
    // This test targets `DirStream::parse_with_encoding` (used by `VBAProject::parse`) rather than
    // the v3 transcript machinery.
    //
    // Real-world MODULESTREAMNAME (0x001A) records can include a Reserved=0x0032 marker followed by
    // an explicit UTF-16LE stream name. The u32 after the Id is SizeOfStreamName (MBCS), not the
    // total record length; if a parser treats it as a generic TLV size it will become misaligned.
    let module_stream_name_unicode = "МодульПоток";
    let module_stream_name_unicode_utf16 = utf16le_bytes(module_stream_name_unicode);

    let module_code = b"Sub Hello()\r\nEnd Sub\r\n";
    let module_container = compress_container(module_code);

    let dir_decompressed = {
        let mut out = Vec::new();

        // PROJECTCODEPAGE (u16 LE) to make encoding resolution deterministic.
        push_record(&mut out, 0x0003, &1251u16.to_le_bytes());

        // MODULENAME is the module identifier; keep it ASCII.
        push_record(&mut out, 0x0019, b"Module1");

        // MODULESTREAMNAME in spec layout:
        //   Id(u16)=0x001A
        //   SizeOfStreamName(u32)
        //   StreamName (MBCS bytes) -- deliberately wrong
        //   Reserved(u16)=0x0032
        //   SizeOfStreamNameUnicode(u32)
        //   StreamNameUnicode(UTF-16LE bytes)
        out.extend_from_slice(&0x001Au16.to_le_bytes());
        out.extend_from_slice(&(b"Wrong".len() as u32).to_le_bytes());
        out.extend_from_slice(b"Wrong");
        out.extend_from_slice(&0x0032u16.to_le_bytes());
        out.extend_from_slice(&(module_stream_name_unicode_utf16.len() as u32).to_le_bytes());
        out.extend_from_slice(&module_stream_name_unicode_utf16);

        // MODULETYPE (standard) + MODULETEXTOFFSET (0)
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
    let project = VBAProject::parse(&vba_bin).expect("parse");
    let module = project
        .modules
        .iter()
        .find(|m| m.name == "Module1")
        .expect("Module1 present");
    assert_eq!(
        module.stream_name, module_stream_name_unicode,
        "expected MODULESTREAMNAME unicode tail to be used for OLE stream lookup"
    );
    assert!(module.code.contains("Sub Hello"));
}

#[test]
fn vba_project_parse_supports_modulestreamname_unicode_record_id_0048() {
    // `DirStream::parse_with_encoding` historically supported a nonstandard layout where a Unicode
    // module stream-name is stored in a separate record with id 0x0048 immediately following the
    // MODULESTREAMNAME (0x001A) record. (0x0048 is normally used for MODULEDOCSTRINGUNICODE, but
    // some files appear to reuse it for the stream-name field.)
    //
    // Ensure MODULESTREAMNAME special-casing does not regress this behavior.
    let module_stream_name_unicode = "ユニコード名";
    let module_stream_name_unicode_utf16 = utf16le_bytes(module_stream_name_unicode);

    let module_code = b"Sub Hello()\r\nEnd Sub\r\n";
    let module_container = compress_container(module_code);

    let dir_decompressed = {
        let mut out = Vec::new();

        // PROJECTCODEPAGE (u16 LE) to make encoding resolution deterministic.
        push_record(&mut out, 0x0003, &1252u16.to_le_bytes());

        // MODULENAME is the module identifier; keep it ASCII.
        push_record(&mut out, 0x0019, b"Module1");

        // MODULESTREAMNAME (ANSI/MBCS): deliberately wrong to ensure we must use the Unicode
        // variant for OLE lookup.
        push_record(&mut out, 0x001A, b"Wrong");

        // Nonstandard Unicode stream name record id (0x0048).
        push_record(&mut out, 0x0048, &module_stream_name_unicode_utf16);

        // MODULETYPE (standard) + MODULETEXTOFFSET (0)
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
    let project = VBAProject::parse(&vba_bin).expect("parse");
    let module = project
        .modules
        .iter()
        .find(|m| m.name == "Module1")
        .expect("Module1 present");
    assert_eq!(
        module.stream_name, module_stream_name_unicode,
        "expected MODULESTREAMNAME unicode record (0x0048) to be used for OLE stream lookup"
    );
    assert!(module.code.contains("Sub Hello"));
}
