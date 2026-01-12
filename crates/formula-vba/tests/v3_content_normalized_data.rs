use std::io::{Cursor, Write};

use formula_vba::{
    compress_container, content_normalized_data, contents_hash_v3, forms_normalized_data,
    project_normalized_data_v3_transcript, v3_content_normalized_data,
};

fn push_record(out: &mut Vec<u8>, id: u16, data: &[u8]) {
    out.extend_from_slice(&id.to_le_bytes());
    out.extend_from_slice(&(data.len() as u32).to_le_bytes());
    out.extend_from_slice(data);
}

fn push_module_stream_name_record(
    out: &mut Vec<u8>,
    stream_name_mbcs: &[u8],
    stream_name_unicode: &[u8],
) {
    // MS-OVBA MODULESTREAMNAME (0x001A):
    //   Id (u16)
    //   SizeOfStreamName (u32) -- bytes
    //   StreamName (MBCS)
    //   Reserved (u16) = 0x0032
    //   SizeOfStreamNameUnicode (u32) -- bytes (even)
    //   StreamNameUnicode (UTF-16LE)
    //
    // Note: this record is not representable with `push_record(id, data)` because the u32 size
    // field is `SizeOfStreamName`, not the length of all subsequent bytes.
    out.extend_from_slice(&0x001Au16.to_le_bytes());
    out.extend_from_slice(&(stream_name_mbcs.len() as u32).to_le_bytes());
    out.extend_from_slice(stream_name_mbcs);
    out.extend_from_slice(&0x0032u16.to_le_bytes());
    out.extend_from_slice(&(stream_name_unicode.len() as u32).to_le_bytes());
    out.extend_from_slice(stream_name_unicode);
}

fn utf16le_bytes(s: &str) -> Vec<u8> {
    s.encode_utf16()
        .flat_map(|u| u.to_le_bytes())
        .collect::<Vec<u8>>()
}

fn contains_subslice(haystack: &[u8], needle: &[u8]) -> bool {
    if needle.is_empty() {
        return true;
    }
    haystack
        .windows(needle.len())
        .any(|window| window == needle)
}

fn build_project_no_designers() -> Vec<u8> {
    // Keep module source already in normalized form to make expected bytes simple.
    let module_code = b"Sub Foo()\r\nEnd Sub\r\n";
    let module_container = compress_container(module_code);

    let dir_decompressed = {
        let mut out = Vec::new();
        push_record(&mut out, 0x0019, b"Module1"); // MODULENAME

        // MODULESTREAMNAME + reserved u16.
        let mut stream_name = Vec::new();
        stream_name.extend_from_slice(b"Module1");
        stream_name.extend_from_slice(&0u16.to_le_bytes());
        push_record(&mut out, 0x001A, &stream_name);

        // MODULETYPE (procedural; TypeRecord.Id=0x0021)
        // Reserved is a u16 and is typically 0x0000.
        push_record(&mut out, 0x0021, &0u16.to_le_bytes());
        push_record(&mut out, 0x0031, &0u32.to_le_bytes()); // MODULETEXTOFFSET (0)
        out
    };
    let dir_container = compress_container(&dir_decompressed);

    let cursor = Cursor::new(Vec::new());
    let mut ole = cfb::CompoundFile::create(cursor).expect("create cfb");
    ole.create_storage("VBA").expect("VBA storage");
    {
        let mut s = ole.create_stream("PROJECT").expect("PROJECT stream");
        s.write_all(b"Name=\"VBAProject\"\r\n")
            .expect("write PROJECT");
    }
    {
        let mut s = ole.create_stream("VBA/dir").expect("dir stream");
        s.write_all(&dir_container).expect("write dir");
    }
    {
        let mut s = ole.create_stream("VBA/Module1").expect("module stream");
        s.write_all(&module_container).expect("write module");
    }

    ole.into_inner().into_inner()
}

fn build_project_with_designer_storage() -> Vec<u8> {
    let userform_module_code = b"Sub FormHello()\r\nEnd Sub\r\n";
    let userform_container = compress_container(userform_module_code);

    // PROJECT must reference the designer module via BaseClass= for FormsNormalizedData.
    let project_stream = b"Name=\"VBAProject\"\r\nBaseClass=\"UserForm1\"\r\n";

    let dir_decompressed = {
        let mut out = Vec::new();
        push_record(&mut out, 0x0019, b"UserForm1"); // MODULENAME

        // MODULESTREAMNAME + reserved u16.
        let mut stream_name = Vec::new();
        stream_name.extend_from_slice(b"UserForm1");
        stream_name.extend_from_slice(&0u16.to_le_bytes());
        push_record(&mut out, 0x001A, &stream_name);

        // MODULETYPE (non-procedural; TypeRecord.Id=0x0022)
        // Reserved is a u16 and is ignored by the v3 transcript pseudocode.
        push_record(&mut out, 0x0022, &0u16.to_le_bytes());
        push_record(&mut out, 0x0031, &0u32.to_le_bytes()); // MODULETEXTOFFSET (0)
        out
    };
    let dir_container = compress_container(&dir_decompressed);

    let cursor = Cursor::new(Vec::new());
    let mut ole = cfb::CompoundFile::create(cursor).expect("create cfb");
    ole.create_storage("VBA").expect("VBA storage");
    ole.create_storage("UserForm1").expect("designer storage");

    {
        let mut s = ole.create_stream("PROJECT").expect("PROJECT stream");
        s.write_all(project_stream).expect("write PROJECT");
    }
    {
        let mut s = ole.create_stream("VBA/dir").expect("dir stream");
        s.write_all(&dir_container).expect("write dir");
    }
    {
        let mut s = ole
            .create_stream("VBA/UserForm1")
            .expect("userform module stream");
        s.write_all(&userform_container)
            .expect("write userform module");
    }
    {
        let mut s = ole
            .create_stream("UserForm1/Payload")
            .expect("designer stream");
        s.write_all(b"ABC").expect("write designer bytes");
    }

    ole.into_inner().into_inner()
}

fn build_project_unicode_only_module_stream_name() -> Vec<u8> {
    // Keep module source already in normalized form to make expected bytes simple.
    let module_code = b"Sub Foo()\r\nEnd Sub\r\n";
    let module_container = compress_container(module_code);

    // Unicode module identifiers + stream names.
    let module_name_unicode = "模块名";
    let module_stream_name_unicode = "模块1";

    let dir_decompressed = {
        let mut out = Vec::new();

        // MODULENAMEUNICODE only (no ANSI MODULENAME).
        push_record(&mut out, 0x0047, &utf16le_bytes(module_name_unicode));

        // MODULESTREAMNAMEUNICODE only (no ANSI MODULESTREAMNAME).
        push_record(&mut out, 0x0032, &utf16le_bytes(module_stream_name_unicode));

        // MODULETYPE (procedural; TypeRecord.Id=0x0021)
        push_record(&mut out, 0x0021, &0u16.to_le_bytes());
        push_record(&mut out, 0x0031, &0u32.to_le_bytes()); // MODULETEXTOFFSET (0)
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

    ole.into_inner().into_inner()
}

fn build_project_unicode_only_module_stream_name_with_project_stream() -> Vec<u8> {
    // Keep module source already in normalized form to make expected bytes simple.
    let module_code = b"Sub Foo()\r\nEnd Sub\r\n";
    let module_container = compress_container(module_code);

    // Unicode module identifiers + stream names.
    let module_name_unicode = "模块名";
    let module_stream_name_unicode = "模块1";

    let dir_decompressed = {
        let mut out = Vec::new();

        // MODULENAMEUNICODE only (no ANSI MODULENAME).
        push_record(&mut out, 0x0047, &utf16le_bytes(module_name_unicode));

        // MODULESTREAMNAMEUNICODE only (no ANSI MODULESTREAMNAME).
        push_record(&mut out, 0x0032, &utf16le_bytes(module_stream_name_unicode));

        // MODULETYPE (procedural; TypeRecord.Id=0x0021)
        push_record(&mut out, 0x0021, &0u16.to_le_bytes());
        push_record(&mut out, 0x0031, &0u32.to_le_bytes()); // MODULETEXTOFFSET (0)
        out
    };
    let dir_container = compress_container(&dir_decompressed);

    let cursor = Cursor::new(Vec::new());
    let mut ole = cfb::CompoundFile::create(cursor).expect("create cfb");
    ole.create_storage("VBA").expect("VBA storage");
    {
        let mut s = ole.create_stream("PROJECT").expect("PROJECT stream");
        s.write_all(b"Name=\"VBAProject\"\r\n")
            .expect("write PROJECT");
    }
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

    ole.into_inner().into_inner()
}

fn build_project_unicode_module_stream_name_with_internal_len_prefix() -> Vec<u8> {
    // Keep module source already in normalized form to make expected bytes simple.
    let module_code = b"Sub Foo()\r\nEnd Sub\r\n";
    let module_container = compress_container(module_code);

    // Unicode module identifiers + stream names.
    let module_name_unicode = "模块名";
    let module_stream_name_unicode = "模块1";

    // Build MODULESTREAMNAMEUNICODE payload with an internal u32 length prefix (byte count).
    let stream_name_utf16 = utf16le_bytes(module_stream_name_unicode);
    let mut stream_name_payload = Vec::new();
    stream_name_payload.extend_from_slice(&(stream_name_utf16.len() as u32).to_le_bytes());
    stream_name_payload.extend_from_slice(&stream_name_utf16);

    let dir_decompressed = {
        let mut out = Vec::new();

        // MODULENAMEUNICODE only (no ANSI MODULENAME).
        push_record(&mut out, 0x0047, &utf16le_bytes(module_name_unicode));

        // MODULESTREAMNAMEUNICODE only (no ANSI MODULESTREAMNAME), with internal length prefix.
        push_record(&mut out, 0x0032, &stream_name_payload);

        // MODULETYPE (procedural; TypeRecord.Id=0x0021)
        push_record(&mut out, 0x0021, &0u16.to_le_bytes());
        push_record(&mut out, 0x0031, &0u32.to_le_bytes()); // MODULETEXTOFFSET (0)
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

    ole.into_inner().into_inner()
}

fn build_project_with_ansi_and_unicode_module_name_records() -> Vec<u8> {
    // Keep module source already in normalized form to make expected bytes simple.
    let module_code = b"Sub Foo()\r\nEnd Sub\r\n";
    let module_container = compress_container(module_code);

    let dir_decompressed = {
        let mut out = Vec::new();

        // MODULENAME (ANSI) + MODULENAMEUNICODE.
        push_record(&mut out, 0x0019, b"Module1");
        push_record(&mut out, 0x0047, &utf16le_bytes("Module1"));

        // MODULESTREAMNAME (ANSI) + reserved u16.
        let mut stream_name = Vec::new();
        stream_name.extend_from_slice(b"Module1");
        stream_name.extend_from_slice(&0u16.to_le_bytes());
        push_record(&mut out, 0x001A, &stream_name);

        // MODULETYPE (procedural; TypeRecord.Id=0x0021)
        push_record(&mut out, 0x0021, &0u16.to_le_bytes());
        push_record(&mut out, 0x0031, &0u32.to_le_bytes()); // MODULETEXTOFFSET (0)
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
        let mut s = ole.create_stream("VBA/Module1").expect("module stream");
        s.write_all(&module_container).expect("write module");
    }

    ole.into_inner().into_inner()
}

fn build_project_with_modulenameunicode_internal_len_prefix() -> Vec<u8> {
    // Keep module source already in normalized form to make expected bytes simple.
    let module_code = b"Sub Foo()\r\nEnd Sub\r\n";
    let module_container = compress_container(module_code);

    let unicode_name_bytes = utf16le_bytes("Module1");
    let mut unicode_name_payload = Vec::new();
    unicode_name_payload.extend_from_slice(&(unicode_name_bytes.len() as u32).to_le_bytes());
    unicode_name_payload.extend_from_slice(&unicode_name_bytes);

    let dir_decompressed = {
        let mut out = Vec::new();

        // MODULENAME (ANSI) + MODULENAMEUNICODE with an internal u32 length prefix.
        push_record(&mut out, 0x0019, b"Module1");
        push_record(&mut out, 0x0047, &unicode_name_payload);

        // MODULESTREAMNAME (ANSI) + reserved u16.
        let mut stream_name = Vec::new();
        stream_name.extend_from_slice(b"Module1");
        stream_name.extend_from_slice(&0u16.to_le_bytes());
        push_record(&mut out, 0x001A, &stream_name);

        // MODULETYPE (procedural; TypeRecord.Id=0x0021)
        push_record(&mut out, 0x0021, &0u16.to_le_bytes());
        push_record(&mut out, 0x0031, &0u32.to_le_bytes()); // MODULETEXTOFFSET (0)
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
        let mut s = ole.create_stream("VBA/Module1").expect("module stream");
        s.write_all(&module_container).expect("write module");
    }

    ole.into_inner().into_inner()
}

fn build_project_with_modulenameunicode_internal_len_prefix_excluding_trailing_nul_byte_count(
) -> Vec<u8> {
    // Keep module source already in normalized form to make expected bytes simple.
    let module_code = b"Sub Foo()\r\nEnd Sub\r\n";
    let module_container = compress_container(module_code);

    let name_utf16 = utf16le_bytes("Module1");
    let mut name_utf16_with_nul = name_utf16.clone();
    name_utf16_with_nul.extend_from_slice(&0u16.to_le_bytes());
    let byte_len_without_nul = name_utf16.len() as u32;

    let mut unicode_name_payload = Vec::new();
    unicode_name_payload.extend_from_slice(&byte_len_without_nul.to_le_bytes());
    unicode_name_payload.extend_from_slice(&name_utf16_with_nul);

    let dir_decompressed = {
        let mut out = Vec::new();

        // MODULENAME (ANSI) + MODULENAMEUNICODE with an internal u32 length prefix that excludes the
        // trailing UTF-16 NUL terminator.
        push_record(&mut out, 0x0019, b"Module1");
        push_record(&mut out, 0x0047, &unicode_name_payload);

        // MODULESTREAMNAME (ANSI) + reserved u16.
        let mut stream_name = Vec::new();
        stream_name.extend_from_slice(b"Module1");
        stream_name.extend_from_slice(&0u16.to_le_bytes());
        push_record(&mut out, 0x001A, &stream_name);

        // MODULETYPE (procedural; TypeRecord.Id=0x0021)
        push_record(&mut out, 0x0021, &0u16.to_le_bytes());
        push_record(&mut out, 0x0031, &0u32.to_le_bytes()); // MODULETEXTOFFSET (0)
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
        let mut s = ole.create_stream("VBA/Module1").expect("module stream");
        s.write_all(&module_container).expect("write module");
    }

    ole.into_inner().into_inner()
}

fn build_project_with_modulenameunicode_internal_len_prefix_excluding_trailing_nul_code_units(
) -> Vec<u8> {
    // Keep module source already in normalized form to make expected bytes simple.
    let module_code = b"Sub Foo()\r\nEnd Sub\r\n";
    let module_container = compress_container(module_code);

    let name_utf16 = utf16le_bytes("Module1");
    let mut name_utf16_with_nul = name_utf16.clone();
    name_utf16_with_nul.extend_from_slice(&0u16.to_le_bytes());
    let code_unit_len_without_nul = (name_utf16.len() / 2) as u32;

    let mut unicode_name_payload = Vec::new();
    unicode_name_payload.extend_from_slice(&code_unit_len_without_nul.to_le_bytes());
    unicode_name_payload.extend_from_slice(&name_utf16_with_nul);

    let dir_decompressed = {
        let mut out = Vec::new();

        // MODULENAME (ANSI) + MODULENAMEUNICODE with an internal u32 length prefix that excludes the
        // trailing UTF-16 NUL terminator.
        push_record(&mut out, 0x0019, b"Module1");
        push_record(&mut out, 0x0047, &unicode_name_payload);

        // MODULESTREAMNAME (ANSI) + reserved u16.
        let mut stream_name = Vec::new();
        stream_name.extend_from_slice(b"Module1");
        stream_name.extend_from_slice(&0u16.to_le_bytes());
        push_record(&mut out, 0x001A, &stream_name);

        // MODULETYPE (procedural; TypeRecord.Id=0x0021)
        push_record(&mut out, 0x0021, &0u16.to_le_bytes());
        push_record(&mut out, 0x0031, &0u32.to_le_bytes()); // MODULETEXTOFFSET (0)
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
        let mut s = ole.create_stream("VBA/Module1").expect("module stream");
        s.write_all(&module_container).expect("write module");
    }

    ole.into_inner().into_inner()
}

fn build_project_with_ansi_and_unicode_module_stream_name_records() -> Vec<u8> {
    // Keep module source already in normalized form to make expected bytes simple.
    let module_code = b"Sub Foo()\r\nEnd Sub\r\n";
    let module_container = compress_container(module_code);

    // Unicode stream name that does not exist in ANSI.
    let module_stream_name_unicode = "模块1";

    let dir_decompressed = {
        let mut out = Vec::new();

        // MODULENAME (ANSI).
        push_record(&mut out, 0x0019, b"Module1");

        // MODULESTREAMNAME (ANSI) points at a stream we will *not* create, to ensure the Unicode
        // variant is preferred for OLE lookup.
        let mut stream_name = Vec::new();
        stream_name.extend_from_slice(b"Module1");
        stream_name.extend_from_slice(&0u16.to_le_bytes());
        push_record(&mut out, 0x001A, &stream_name);

        // MODULESTREAMNAMEUNICODE points at the actual module stream.
        push_record(&mut out, 0x0032, &utf16le_bytes(module_stream_name_unicode));

        // MODULETYPE (procedural; TypeRecord.Id=0x0021)
        push_record(&mut out, 0x0021, &0u16.to_le_bytes());
        push_record(&mut out, 0x0031, &0u32.to_le_bytes()); // MODULETEXTOFFSET (0)
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

    ole.into_inner().into_inner()
}

fn build_project_with_projectcompatversion(include_compat: bool) -> (Vec<u8>, [u8; 4]) {
    // Distinctive compat version payload so the regression assertion is unambiguous.
    let compat_version = 0xDEADBEEFu32.to_le_bytes();
    let project_cookie = 0xBEEFu16.to_le_bytes();

    // Keep module source already in normalized form to make expected bytes simple.
    let module_code = b"Sub Foo()\r\nEnd Sub\r\n";
    let module_container = compress_container(module_code);

    // Build a decompressed `VBA/dir` stream that contains the required ProjectInformation records
    // plus an inserted PROJECTCOMPATVERSION (0x004A), then the minimal module record group.
    //
    // This mirrors real-world `VBA/dir` layouts where PROJECTCOMPATVERSION may appear between
    // PROJECTVERSION and PROJECTCONSTANTS.
    let dir_decompressed = {
        let mut out = Vec::new();

        // ---- ProjectInformation (required records) ----
        // PROJECTSYSKIND (0x0001): SysKind (u32).
        push_record(&mut out, 0x0001, &1u32.to_le_bytes());
        // PROJECTLCID (0x0002): Lcid (u32).
        push_record(&mut out, 0x0002, &0x0409u32.to_le_bytes());
        // PROJECTLCIDINVOKE (0x0014): LcidInvoke (u32).
        push_record(&mut out, 0x0014, &0x0409u32.to_le_bytes());
        // PROJECTCODEPAGE (0x0003): CodePage (u16).
        push_record(&mut out, 0x0003, &1252u16.to_le_bytes());

        // PROJECTNAME (0x0004): MBCS string in the project codepage.
        push_record(&mut out, 0x0004, b"VBAProject");
        // PROJECTDOCSTRING (0x0005): MBCS string.
        push_record(&mut out, 0x0005, b"DocString");
        // PROJECTHELPFILEPATH (0x0006): MBCS string.
        push_record(&mut out, 0x0006, b"C:\\help.chm");
        // PROJECTHELPCONTEXT (0x0007): u32.
        push_record(&mut out, 0x0007, &0u32.to_le_bytes());
        // PROJECTLIBFLAGS (0x0008): u32.
        push_record(&mut out, 0x0008, &0u32.to_le_bytes());

        // PROJECTVERSION (0x0009).
        //
        // MS-OVBA §2.3.4.2.1.11: This record is fixed-length (no u32 Size field):
        //   Id(u16) || Reserved(u32=4) || VersionMajor(u32) || VersionMinor(u16)
        out.extend_from_slice(&0x0009u16.to_le_bytes());
        out.extend_from_slice(&0x00000004u32.to_le_bytes());
        out.extend_from_slice(&1u32.to_le_bytes());
        out.extend_from_slice(&0u16.to_le_bytes());

        // PROJECTCOMPATVERSION (0x004A): present in many real-world files but must be skipped by
        // the MS-OVBA §2.4.2.5 V3ContentNormalizedData pseudocode.
        if include_compat {
            push_record(&mut out, 0x004A, &compat_version);
        }

        // PROJECTCONSTANTS (0x000C): MBCS string.
        push_record(&mut out, 0x000C, b"Constants");

        // ---- ProjectModules ----
        //
        // Include the ProjectModules/ProjectCookie headers and the dir terminator record so this
        // fixture more closely mirrors real-world `VBA/dir` layouts.
        push_record(&mut out, 0x000F, &1u16.to_le_bytes()); // PROJECTMODULES (Count=1)
        push_record(&mut out, 0x0013, &project_cookie); // PROJECTCOOKIE (Cookie; excluded from transcript)

        // ---- Module record group (minimal) ----
        push_record(&mut out, 0x0019, b"Module1"); // MODULENAME

        // MODULESTREAMNAME + reserved u16.
        let mut stream_name = Vec::new();
        stream_name.extend_from_slice(b"Module1");
        stream_name.extend_from_slice(&0u16.to_le_bytes());
        push_record(&mut out, 0x001A, &stream_name);

        push_record(&mut out, 0x0021, &0u16.to_le_bytes()); // MODULETYPE (standard)
        push_record(&mut out, 0x0031, &0u32.to_le_bytes()); // MODULETEXTOFFSET (0)

        // Dir stream terminator + reserved (treated as a record with Size=0 in TLV-style fixtures).
        push_record(&mut out, 0x0010, &[]);
        out
    };
    let dir_container = compress_container(&dir_decompressed);

    let cursor = Cursor::new(Vec::new());
    let mut ole = cfb::CompoundFile::create(cursor).expect("create cfb");
    ole.create_storage("VBA").expect("VBA storage");
    {
        let mut s = ole.create_stream("PROJECT").expect("PROJECT stream");
        s.write_all(b"Name=\"VBAProject\"\r\n")
            .expect("write PROJECT");
    }
    {
        let mut s = ole.create_stream("VBA/dir").expect("dir stream");
        s.write_all(&dir_container).expect("write dir");
    }
    {
        let mut s = ole.create_stream("VBA/Module1").expect("module stream");
        s.write_all(&module_container).expect("write module");
    }

    (ole.into_inner().into_inner(), compat_version)
}

fn build_project_with_project_info_records_only() -> Vec<u8> {
    // ---- Spec-backed project-information records (MS-OVBA §2.4.2.5) ----
    //
    // We deliberately use recognizable values so we can assert which bytes are included vs.
    // excluded by V3ContentNormalizedData.

    let project_name = b"MyV3Project";

    let syskind = 0xA1B2C3D4u32; // value bytes must NOT appear in V3 transcript
    let lcid = 0x11223344u32;
    let lcid_invoke = 0x55667788u32;
    let codepage = 0xCAFEu16; // value bytes must NOT appear in V3 transcript

    let docstring = b"__DOCSTRING_BYTES__"; // must NOT appear in V3 transcript
                                            // UTF-16LE bytes; do not include NULs.
    let docstring_unicode = b"D\0O\0C\0U\0N\0I\0";

    let helpfile1 = b"__HELPFILE1_PATH__"; // must NOT appear in V3 transcript
    let helpfile2 = b"__HELPFILE2_PATH__"; // must NOT appear in V3 transcript

    let helpcontext = 0x0BADF00Du32; // value bytes must NOT appear in V3 transcript

    let project_lib_flags = 0x01020304u32;

    let version_major = 0xCAFEBABEu32;
    let version_minor = 0xBEEFu16;

    let constants = b"ABC=1";
    // UTF-16LE("ABC=1").
    let constants_unicode = b"A\x00B\x00C\x00=\x001\x00";

    // Build a decompressed `VBA/dir` stream containing the project-information records referenced
    // by MS-OVBA §2.4.2.5 `V3ContentNormalizedData`.
    let dir_decompressed = {
        let mut out = Vec::new();

        // PROJECTSYSKIND (0x0001): u32 SysKind
        push_record(&mut out, 0x0001, &syskind.to_le_bytes());

        // PROJECTLCID (0x0002): u32 Lcid
        push_record(&mut out, 0x0002, &lcid.to_le_bytes());

        // PROJECTLCIDINVOKE (0x0014): u32 LcidInvoke
        push_record(&mut out, 0x0014, &lcid_invoke.to_le_bytes());

        // PROJECTCODEPAGE (0x0003): u16 CodePage
        push_record(&mut out, 0x0003, &codepage.to_le_bytes());

        // PROJECTNAME (0x0004): bytes ProjectName
        push_record(&mut out, 0x0004, project_name);

        // PROJECTDOCSTRING (0x0005) and its unicode sub-record (0x0040).
        push_record(&mut out, 0x0005, docstring);
        push_record(&mut out, 0x0040, docstring_unicode);

        // PROJECTHELPFILEPATH (0x0006) and its second path sub-record (0x003D).
        push_record(&mut out, 0x0006, helpfile1);
        push_record(&mut out, 0x003D, helpfile2);

        // PROJECTHELPCONTEXT (0x0007): u32 HelpContext
        push_record(&mut out, 0x0007, &helpcontext.to_le_bytes());

        // PROJECTLIBFLAGS (0x0008): u32 ProjectLibFlags
        push_record(&mut out, 0x0008, &project_lib_flags.to_le_bytes());

        // PROJECTVERSION (0x0009): fixed-length record (no u32 Size field).
        out.extend_from_slice(&0x0009u16.to_le_bytes()); // Id
        out.extend_from_slice(&0x00000004u32.to_le_bytes()); // Reserved (MUST be 4)
        out.extend_from_slice(&version_major.to_le_bytes());
        out.extend_from_slice(&version_minor.to_le_bytes());

        // PROJECTCONSTANTS (0x000C) and its unicode sub-record (0x003C).
        push_record(&mut out, 0x000C, constants);
        push_record(&mut out, 0x003C, constants_unicode);

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

    ole.into_inner().into_inner()
}

fn build_project_with_project_info_records_only_noncanonical_unicode_ids() -> Vec<u8> {
    // Same as `build_project_with_project_info_records_only()`, but uses observed non-canonical
    // record ids for the Unicode/alternate sub-records:
    // - 0x0041 (PROJECTDOCSTRING Unicode)
    // - 0x0042 (PROJECTHELPFILEPATH2)
    // - 0x0043 (PROJECTCONSTANTS Unicode)

    let project_name = b"MyV3Project";

    let syskind = 0xA1B2C3D4u32; // value bytes must NOT appear in V3 transcript
    let lcid = 0x11223344u32;
    let lcid_invoke = 0x55667788u32;
    let codepage = 0xCAFEu16; // value bytes must NOT appear in V3 transcript

    let docstring = b"__DOCSTRING_BYTES__"; // must NOT appear in V3 transcript
                                            // UTF-16LE bytes; do not include NULs.
    let docstring_unicode = b"D\0O\0C\0U\0N\0I\0";

    let helpfile1 = b"__HELPFILE1_PATH__"; // must NOT appear in V3 transcript
    let helpfile2 = b"__HELPFILE2_PATH__"; // must NOT appear in V3 transcript

    let helpcontext = 0x0BADF00Du32; // value bytes must NOT appear in V3 transcript

    let project_lib_flags = 0x01020304u32;

    let version_major = 0xCAFEBABEu32;
    let version_minor = 0xBEEFu16;

    let constants = b"ABC=1";
    let constants_unicode = b"A\x00B\x00C\x00=\x001\x00";

    let dir_decompressed = {
        let mut out = Vec::new();

        push_record(&mut out, 0x0001, &syskind.to_le_bytes()); // PROJECTSYSKIND
        push_record(&mut out, 0x0002, &lcid.to_le_bytes()); // PROJECTLCID
        push_record(&mut out, 0x0014, &lcid_invoke.to_le_bytes()); // PROJECTLCIDINVOKE
        push_record(&mut out, 0x0003, &codepage.to_le_bytes()); // PROJECTCODEPAGE

        push_record(&mut out, 0x0004, project_name); // PROJECTNAME

        // PROJECTDOCSTRING + non-canonical Unicode sub-record.
        push_record(&mut out, 0x0005, docstring);
        push_record(&mut out, 0x0041, docstring_unicode);

        // PROJECTHELPFILEPATH + non-canonical HelpFile2 sub-record.
        push_record(&mut out, 0x0006, helpfile1);
        push_record(&mut out, 0x0042, helpfile2);

        push_record(&mut out, 0x0007, &helpcontext.to_le_bytes()); // PROJECTHELPCONTEXT
        push_record(&mut out, 0x0008, &project_lib_flags.to_le_bytes()); // PROJECTLIBFLAGS

        // PROJECTVERSION (fixed-length record).
        out.extend_from_slice(&0x0009u16.to_le_bytes()); // Id
        out.extend_from_slice(&0x00000004u32.to_le_bytes()); // Reserved (MUST be 4)
        out.extend_from_slice(&version_major.to_le_bytes());
        out.extend_from_slice(&version_minor.to_le_bytes());

        // PROJECTCONSTANTS + non-canonical Unicode sub-record.
        push_record(&mut out, 0x000C, constants);
        push_record(&mut out, 0x0043, constants_unicode);

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

    ole.into_inner().into_inner()
}

fn build_project_with_project_info_records_only_tlv_projectversion() -> Vec<u8> {
    // Same as `build_project_with_project_info_records_only`, but encodes PROJECTVERSION (0x0009)
    // using a TLV framing (`Id || Size || Data`) instead of the spec fixed-length layout.
    //
    // Some real-world producers and fixtures use this encoding; `v3_content_normalized_data` should
    // still be able to scan the dir stream without losing alignment.

    let project_name = b"MyV3Project";

    let syskind = 0xA1B2C3D4u32;
    let lcid = 0x11223344u32;
    let lcid_invoke = 0x55667788u32;
    let codepage = 0xCAFEu16;

    let docstring = b"__DOCSTRING_BYTES__";
    let docstring_unicode = b"D\0O\0C\0U\0N\0I\0";

    let helpfile1 = b"__HELPFILE1_PATH__";
    let helpfile2 = b"__HELPFILE2_PATH__";

    let helpcontext = 0x0BADF00Du32;

    let project_lib_flags = 0x01020304u32;

    let version_major = 0xCAFEBABEu32;
    let version_minor = 0xBEEFu16;

    let constants = b"ABC=1";
    let constants_unicode = b"A\x00B\x00C\x00=\x001\x00";

    let dir_decompressed = {
        let mut out = Vec::new();

        push_record(&mut out, 0x0001, &syskind.to_le_bytes());
        push_record(&mut out, 0x0002, &lcid.to_le_bytes());
        push_record(&mut out, 0x0014, &lcid_invoke.to_le_bytes());
        push_record(&mut out, 0x0003, &codepage.to_le_bytes());
        push_record(&mut out, 0x0004, project_name);
        push_record(&mut out, 0x0005, docstring);
        push_record(&mut out, 0x0040, docstring_unicode);
        push_record(&mut out, 0x0006, helpfile1);
        push_record(&mut out, 0x003D, helpfile2);
        push_record(&mut out, 0x0007, &helpcontext.to_le_bytes());
        push_record(&mut out, 0x0008, &project_lib_flags.to_le_bytes());

        // PROJECTVERSION (0x0009) encoded as TLV: Data == Reserved(u32) || VersionMajor(u32) ||
        // VersionMinor(u16).
        let mut version = Vec::new();
        version.extend_from_slice(&0x00000004u32.to_le_bytes());
        version.extend_from_slice(&version_major.to_le_bytes());
        version.extend_from_slice(&version_minor.to_le_bytes());
        push_record(&mut out, 0x0009, &version);

        push_record(&mut out, 0x000C, constants);
        push_record(&mut out, 0x003C, constants_unicode);
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

    ole.into_inner().into_inner()
}

#[test]
fn v3_content_normalized_data_includes_module_metadata_even_without_designers() {
    let vba_bin = build_project_no_designers();

    let v3 = v3_content_normalized_data(&vba_bin).expect("V3ContentNormalizedData");
    let content = content_normalized_data(&vba_bin).expect("ContentNormalizedData");
    assert_eq!(content, b"Sub Foo()\r\nEnd Sub\r\n".to_vec());

    // Per MS-OVBA §2.4.2.5, the module transcript includes:
    // - (TypeRecord.Id || Reserved) when TypeRecord.Id == 0x0021
    // - LF-normalized module source with Attribute filtering
    // - a trailing module name + LF when `HashModuleNameFlag` becomes true
    let mut expected = Vec::new();
    expected.extend_from_slice(&0x0021u16.to_le_bytes());
    expected.extend_from_slice(&0u16.to_le_bytes());
    expected.extend_from_slice(b"Sub Foo()\nEnd Sub\n\nModule1\n");

    assert_ne!(
        v3, content,
        "v3 transcript includes module metadata and should differ from ContentNormalizedData"
    );
    assert_eq!(v3, expected);
}

#[test]
fn v3_content_normalized_data_accepts_tlv_projectversion_record_framing() {
    let fixed = build_project_with_project_info_records_only();
    let tlv = build_project_with_project_info_records_only_tlv_projectversion();

    let expected = v3_content_normalized_data(&fixed)
        .expect("V3ContentNormalizedData should parse fixed-length PROJECTVERSION");
    let actual = v3_content_normalized_data(&tlv)
        .expect("V3ContentNormalizedData should parse TLV PROJECTVERSION");

    assert_eq!(
        actual, expected,
        "PROJECTVERSION TLV framing should not change V3ContentNormalizedData output"
    );
}

#[test]
fn v3_content_normalized_data_uses_unicode_module_and_stream_names_when_present() {
    // Build an in-memory vbaProject.bin with a Unicode module stream name supplied via
    // MODULESTREAMNAME (0x001A) with Reserved=0x0032.
    //
    // The real OLE stream name is Unicode-only. If the implementation ignores StreamNameUnicode
    // (or mishandles Reserved=0x0032), it will fail to open the module stream.
    let module_stream_name_unicode = "МодульПоток"; // non-ASCII
    let module_stream_name_unicode_bytes = utf16le_bytes(module_stream_name_unicode);

    let module_name_ansi = "AnsiModuleName";
    let module_name_unicode = "ИмяМодуля"; // non-ASCII
    let module_name_unicode_bytes = utf16le_bytes(module_name_unicode);

    // Module source must set HashModuleNameFlag=true so the module name is appended.
    let module_source = concat!(
        "Attribute VB_Name = \"IgnoredByV3\"\r\n",
        "Sub Hello()\r\n",
        "End Sub\r\n",
    );
    let module_container = compress_container(module_source.as_bytes());

    let dir_decompressed = {
        let mut out = Vec::new();
        push_record(&mut out, 0x0019, module_name_ansi.as_bytes()); // MODULENAME
        push_record(&mut out, 0x0047, &module_name_unicode_bytes); // MODULENAMEUNICODE

        // MODULESTREAMNAME with both MBCS + Unicode names, Reserved=0x0032.
        // The MBCS name is deliberately wrong/nonexistent to ensure Unicode is used.
        push_module_stream_name_record(
            &mut out,
            b"WrongStreamName",
            &module_stream_name_unicode_bytes,
        );

        push_record(&mut out, 0x0021, &0u16.to_le_bytes()); // MODULETYPE (standard)
        push_record(&mut out, 0x0031, &0u32.to_le_bytes()); // MODULETEXTOFFSET (0)
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
        let stream_path = format!("VBA/{module_stream_name_unicode}");
        let mut s = ole.create_stream(&stream_path).expect("module stream");
        s.write_all(&module_container).expect("write module source");
    }

    let vba_project_bin = ole.into_inner().into_inner();

    let normalized = v3_content_normalized_data(&vba_project_bin)
        .expect("v3_content_normalized_data should succeed (module stream must be found)");

    let mut expected_unicode_suffix = module_name_unicode_bytes.clone();
    expected_unicode_suffix.push(b'\n');
    assert!(
        contains_subslice(&normalized, &expected_unicode_suffix),
        "expected V3ContentNormalizedData to contain UTF-16LE MODULENAMEUNICODE bytes + LF"
    );

    let mut unexpected_ansi_suffix = module_name_ansi.as_bytes().to_vec();
    unexpected_ansi_suffix.push(b'\n');
    assert!(
        !contains_subslice(&normalized, &unexpected_ansi_suffix),
        "expected V3ContentNormalizedData NOT to contain ANSI MODULENAME bytes + LF"
    );
}

#[test]
fn v3_content_normalized_data_handles_modulestreamname_unicode_with_len_prefix_and_nul() {
    // Some real-world producers include an internal u32 length prefix and/or trailing NUL
    // code units in `*_UNICODE` record payloads. Ensure we still decode the stream name correctly
    // when it is supplied via MODULESTREAMNAME (0x001A) with Reserved=0x0032.
    let module_stream_name_unicode = "МодульПоток"; // non-ASCII

    // StreamNameUnicode bytes: `u32 byte_len || utf16le_bytes || trailing_nul`.
    let mut stream_name_utf16 = utf16le_bytes(module_stream_name_unicode);
    stream_name_utf16.extend_from_slice(&0u16.to_le_bytes()); // NUL terminator (defensive)
    let mut module_stream_name_unicode_bytes =
        (stream_name_utf16.len() as u32).to_le_bytes().to_vec();
    module_stream_name_unicode_bytes.extend_from_slice(&stream_name_utf16);

    let module_name_ansi = "AnsiModuleName";
    let module_name_unicode = "ИмяМодуля"; // non-ASCII
    let module_name_unicode_bytes = utf16le_bytes(module_name_unicode);

    // Module source must set HashModuleNameFlag=true so the module name is appended.
    let module_source = concat!(
        "Attribute VB_Name = \"IgnoredByV3\"\r\n",
        "Sub Hello()\r\n",
        "End Sub\r\n",
    );
    let module_container = compress_container(module_source.as_bytes());

    let dir_decompressed = {
        let mut out = Vec::new();
        push_record(&mut out, 0x0019, module_name_ansi.as_bytes()); // MODULENAME
        push_record(&mut out, 0x0047, &module_name_unicode_bytes); // MODULENAMEUNICODE

        // MODULESTREAMNAME with both MBCS + Unicode names, Reserved=0x0032.
        // The MBCS name is deliberately wrong/nonexistent to ensure Unicode is used.
        push_module_stream_name_record(
            &mut out,
            b"WrongStreamName",
            &module_stream_name_unicode_bytes,
        );

        push_record(&mut out, 0x0021, &0u16.to_le_bytes()); // MODULETYPE (standard)
        push_record(&mut out, 0x0031, &0u32.to_le_bytes()); // MODULETEXTOFFSET (0)
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
        let stream_path = format!("VBA/{module_stream_name_unicode}");
        let mut s = ole.create_stream(&stream_path).expect("module stream");
        s.write_all(&module_container).expect("write module source");
    }

    let vba_project_bin = ole.into_inner().into_inner();
    let normalized = v3_content_normalized_data(&vba_project_bin).expect(
        "V3ContentNormalizedData should succeed when StreamNameUnicode is decoded robustly",
    );

    let mut expected_unicode_suffix = module_name_unicode_bytes.clone();
    expected_unicode_suffix.push(b'\n');
    assert!(
        contains_subslice(&normalized, &expected_unicode_suffix),
        "expected V3ContentNormalizedData to contain UTF-16LE MODULENAMEUNICODE bytes + LF"
    );

    let mut unexpected_ansi_suffix = module_name_ansi.as_bytes().to_vec();
    unexpected_ansi_suffix.push(b'\n');
    assert!(
        !contains_subslice(&normalized, &unexpected_ansi_suffix),
        "expected V3ContentNormalizedData NOT to contain ANSI MODULENAME bytes + LF"
    );
}

#[test]
fn v3_content_normalized_data_handles_modulestreamname_unicode_with_len_prefix_excluding_nul() {
    // Variant of the previous regression: some producers include a trailing UTF-16 NUL terminator
    // but do not count it in the internal u32 length prefix.
    let module_stream_name_unicode = "МодульПоток"; // non-ASCII

    // StreamNameUnicode bytes: `u32 byte_len_without_nul || utf16le_bytes || trailing_nul`.
    let mut stream_name_utf16 = utf16le_bytes(module_stream_name_unicode);
    stream_name_utf16.extend_from_slice(&0u16.to_le_bytes()); // NUL terminator
    let byte_len_without_nul = (stream_name_utf16.len() - 2) as u32;
    let mut module_stream_name_unicode_bytes = byte_len_without_nul.to_le_bytes().to_vec();
    module_stream_name_unicode_bytes.extend_from_slice(&stream_name_utf16);

    let module_name_ansi = "AnsiModuleName";
    let module_name_unicode = "ИмяМодуля"; // non-ASCII
    let module_name_unicode_bytes = utf16le_bytes(module_name_unicode);

    // Module source must set HashModuleNameFlag=true so the module name is appended.
    let module_source = concat!(
        "Attribute VB_Name = \"IgnoredByV3\"\r\n",
        "Sub Hello()\r\n",
        "End Sub\r\n",
    );
    let module_container = compress_container(module_source.as_bytes());

    let dir_decompressed = {
        let mut out = Vec::new();
        push_record(&mut out, 0x0019, module_name_ansi.as_bytes()); // MODULENAME
        push_record(&mut out, 0x0047, &module_name_unicode_bytes); // MODULENAMEUNICODE

        // MODULESTREAMNAME with both MBCS + Unicode names, Reserved=0x0032.
        // The MBCS name is deliberately wrong/nonexistent to ensure Unicode is used.
        push_module_stream_name_record(
            &mut out,
            b"WrongStreamName",
            &module_stream_name_unicode_bytes,
        );

        push_record(&mut out, 0x0021, &0u16.to_le_bytes()); // MODULETYPE (standard)
        push_record(&mut out, 0x0031, &0u32.to_le_bytes()); // MODULETEXTOFFSET (0)
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
        let stream_path = format!("VBA/{module_stream_name_unicode}");
        let mut s = ole.create_stream(&stream_path).expect("module stream");
        s.write_all(&module_container).expect("write module source");
    }

    let vba_project_bin = ole.into_inner().into_inner();
    let normalized = v3_content_normalized_data(&vba_project_bin)
        .expect("V3ContentNormalizedData should succeed with len prefix excluding trailing NUL");

    let mut expected_unicode_suffix = module_name_unicode_bytes.clone();
    expected_unicode_suffix.push(b'\n');
    assert!(
        contains_subslice(&normalized, &expected_unicode_suffix),
        "expected V3ContentNormalizedData to contain UTF-16LE MODULENAMEUNICODE bytes + LF"
    );

    let mut unexpected_ansi_suffix = module_name_ansi.as_bytes().to_vec();
    unexpected_ansi_suffix.push(b'\n');
    assert!(
        !contains_subslice(&normalized, &unexpected_ansi_suffix),
        "expected V3ContentNormalizedData NOT to contain ANSI MODULENAME bytes + LF"
    );
}

#[test]
fn v3_content_normalized_data_handles_modulestreamname_unicode_with_code_unit_len_prefix() {
    // Variant of the previous regression: the internal u32 prefix is sometimes a UTF-16 code unit
    // count, not a byte count.
    let module_stream_name_unicode = "МодульПоток"; // non-ASCII

    // StreamNameUnicode bytes: `u32 code_unit_len || utf16le_bytes`.
    let stream_name_utf16 = utf16le_bytes(module_stream_name_unicode);
    let code_unit_len = (stream_name_utf16.len() / 2) as u32;
    let mut module_stream_name_unicode_bytes = code_unit_len.to_le_bytes().to_vec();
    module_stream_name_unicode_bytes.extend_from_slice(&stream_name_utf16);

    let module_name_ansi = "AnsiModuleName";
    let module_name_unicode = "ИмяМодуля"; // non-ASCII
    let module_name_unicode_bytes = utf16le_bytes(module_name_unicode);

    // Module source must set HashModuleNameFlag=true so the module name is appended.
    let module_source = concat!(
        "Attribute VB_Name = \"IgnoredByV3\"\r\n",
        "Sub Hello()\r\n",
        "End Sub\r\n",
    );
    let module_container = compress_container(module_source.as_bytes());

    let dir_decompressed = {
        let mut out = Vec::new();
        push_record(&mut out, 0x0019, module_name_ansi.as_bytes()); // MODULENAME
        push_record(&mut out, 0x0047, &module_name_unicode_bytes); // MODULENAMEUNICODE

        // MODULESTREAMNAME with both MBCS + Unicode names, Reserved=0x0032.
        // The MBCS name is deliberately wrong/nonexistent to ensure Unicode is used.
        push_module_stream_name_record(
            &mut out,
            b"WrongStreamName",
            &module_stream_name_unicode_bytes,
        );

        push_record(&mut out, 0x0021, &0u16.to_le_bytes()); // MODULETYPE (standard)
        push_record(&mut out, 0x0031, &0u32.to_le_bytes()); // MODULETEXTOFFSET (0)
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
        let stream_path = format!("VBA/{module_stream_name_unicode}");
        let mut s = ole.create_stream(&stream_path).expect("module stream");
        s.write_all(&module_container).expect("write module source");
    }

    let vba_project_bin = ole.into_inner().into_inner();
    let normalized = v3_content_normalized_data(&vba_project_bin)
        .expect("V3ContentNormalizedData should succeed when code-unit length prefix is handled");

    let mut expected_unicode_suffix = module_name_unicode_bytes.clone();
    expected_unicode_suffix.push(b'\n');
    assert!(
        contains_subslice(&normalized, &expected_unicode_suffix),
        "expected V3ContentNormalizedData to contain UTF-16LE MODULENAMEUNICODE bytes + LF"
    );

    let mut unexpected_ansi_suffix = module_name_ansi.as_bytes().to_vec();
    unexpected_ansi_suffix.push(b'\n');
    assert!(
        !contains_subslice(&normalized, &unexpected_ansi_suffix),
        "expected V3ContentNormalizedData NOT to contain ANSI MODULENAME bytes + LF"
    );
}

#[test]
fn v3_content_normalized_data_errors_on_truncated_modulestreamname_unicode_tail() {
    // If MODULESTREAMNAME advertises a Unicode tail (Reserved=0x0032) but the stream is truncated,
    // parsing must fail cleanly with DirParseError::Truncated (and must not panic).
    let dir_decompressed = {
        let mut out = Vec::new();
        push_record(&mut out, 0x0019, b"Module1"); // MODULENAME

        // MODULESTREAMNAME record with Unicode tail, but truncated `StreamNameUnicode` bytes.
        out.extend_from_slice(&0x001Au16.to_le_bytes()); // Id
        out.extend_from_slice(&1u32.to_le_bytes()); // SizeOfStreamName
        out.extend_from_slice(b"X"); // StreamName (MBCS)
        out.extend_from_slice(&0x0032u16.to_le_bytes()); // Reserved marker
        out.extend_from_slice(&100u32.to_le_bytes()); // SizeOfStreamNameUnicode (too large)
        out.extend_from_slice(&[0x00, 0x00]); // truncated UTF-16LE bytes

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

    let vba_project_bin = ole.into_inner().into_inner();
    let err = v3_content_normalized_data(&vba_project_bin).expect_err("expected parse error");
    assert!(
        matches!(
            err,
            formula_vba::ParseError::Dir(formula_vba::DirParseError::Truncated)
        ),
        "unexpected error: {err:?}"
    );
}

#[test]
fn v3_content_normalized_data_project_information_includes_only_fields_listed_in_ms_ovba_pseudocode(
) {
    let vba_bin = build_project_with_project_info_records_only();
    let v3 = v3_content_normalized_data(&vba_bin).expect("V3ContentNormalizedData");

    // Project-info values (must match those used in the builder).
    let project_name = b"MyV3Project";

    let syskind = 0xA1B2C3D4u32;
    let lcid = 0x11223344u32;
    let lcid_invoke = 0x55667788u32;
    let codepage = 0xCAFEu16;

    let docstring = b"__DOCSTRING_BYTES__";
    let docstring_unicode = b"D\0O\0C\0U\0N\0I\0";

    let helpfile1 = b"__HELPFILE1_PATH__";
    let helpfile2 = b"__HELPFILE2_PATH__";

    let helpcontext = 0x0BADF00Du32;

    let project_lib_flags = 0x01020304u32;

    let version_major = 0xCAFEBABEu32;
    let version_minor = 0xBEEFu16;

    let constants = b"ABC=1";
    // UTF-16LE("ABC=1").
    let constants_unicode = b"A\x00B\x00C\x00=\x001\x00";

    // Expected prefix per MS-OVBA §2.4.2.5 `V3ContentNormalizedData` pseudocode:
    // - includes only specific fields for some project-info records (e.g. header bytes only)
    // - excludes record payload bytes for others (DocString, HelpFile path bytes, HelpContext value, etc.)
    let mut expected_prefix = Vec::new();

    // PROJECTSYSKIND: include Id + Size, exclude SysKind value.
    expected_prefix.extend_from_slice(&0x0001u16.to_le_bytes());
    expected_prefix.extend_from_slice(&4u32.to_le_bytes());

    // PROJECTLCID: include Id + Size + Lcid value.
    expected_prefix.extend_from_slice(&0x0002u16.to_le_bytes());
    expected_prefix.extend_from_slice(&4u32.to_le_bytes());
    expected_prefix.extend_from_slice(&lcid.to_le_bytes());

    // PROJECTLCIDINVOKE: include Id + Size + LcidInvoke value.
    expected_prefix.extend_from_slice(&0x0014u16.to_le_bytes());
    expected_prefix.extend_from_slice(&4u32.to_le_bytes());
    expected_prefix.extend_from_slice(&lcid_invoke.to_le_bytes());

    // PROJECTCODEPAGE: include Id + Size, exclude CodePage value.
    expected_prefix.extend_from_slice(&0x0003u16.to_le_bytes());
    expected_prefix.extend_from_slice(&2u32.to_le_bytes());

    // PROJECTNAME: include Id + SizeOfProjectName + ProjectName bytes.
    expected_prefix.extend_from_slice(&0x0004u16.to_le_bytes());
    expected_prefix.extend_from_slice(&(project_name.len() as u32).to_le_bytes());
    expected_prefix.extend_from_slice(project_name);

    // PROJECTDOCSTRING: include Id + SizeOfDocString + Reserved + SizeOfDocStringUnicode,
    // but NOT the DocString bytes or DocStringUnicode bytes.
    expected_prefix.extend_from_slice(&0x0005u16.to_le_bytes());
    expected_prefix.extend_from_slice(&(docstring.len() as u32).to_le_bytes());
    expected_prefix.extend_from_slice(&0x0040u16.to_le_bytes());
    expected_prefix.extend_from_slice(&(docstring_unicode.len() as u32).to_le_bytes());

    // PROJECTHELPFILEPATH: include Id + SizeOfHelpFile1 + Reserved + SizeOfHelpFile2,
    // but NOT the help file path bytes.
    expected_prefix.extend_from_slice(&0x0006u16.to_le_bytes());
    expected_prefix.extend_from_slice(&(helpfile1.len() as u32).to_le_bytes());
    expected_prefix.extend_from_slice(&0x003Du16.to_le_bytes());
    expected_prefix.extend_from_slice(&(helpfile2.len() as u32).to_le_bytes());

    // PROJECTHELPCONTEXT: include Id + Size, exclude HelpContext value.
    expected_prefix.extend_from_slice(&0x0007u16.to_le_bytes());
    expected_prefix.extend_from_slice(&4u32.to_le_bytes());

    // PROJECTLIBFLAGS: include Id + Size + ProjectLibFlags value.
    expected_prefix.extend_from_slice(&0x0008u16.to_le_bytes());
    expected_prefix.extend_from_slice(&4u32.to_le_bytes());
    expected_prefix.extend_from_slice(&project_lib_flags.to_le_bytes());

    // PROJECTVERSION: include all fields (fixed-length record).
    expected_prefix.extend_from_slice(&0x0009u16.to_le_bytes());
    expected_prefix.extend_from_slice(&0x00000004u32.to_le_bytes()); // Reserved
    expected_prefix.extend_from_slice(&version_major.to_le_bytes());
    expected_prefix.extend_from_slice(&version_minor.to_le_bytes());

    // PROJECTCONSTANTS: include Id + SizeOfConstants + Constants bytes + Reserved +
    // SizeOfConstantsUnicode + ConstantsUnicode bytes.
    expected_prefix.extend_from_slice(&0x000Cu16.to_le_bytes());
    expected_prefix.extend_from_slice(&(constants.len() as u32).to_le_bytes());
    expected_prefix.extend_from_slice(constants);
    expected_prefix.extend_from_slice(&0x003Cu16.to_le_bytes());
    expected_prefix.extend_from_slice(&(constants_unicode.len() as u32).to_le_bytes());
    expected_prefix.extend_from_slice(constants_unicode);

    assert!(
        v3.len() >= expected_prefix.len(),
        "expected V3ContentNormalizedData to be at least {} bytes, got {}",
        expected_prefix.len(),
        v3.len()
    );
    let prefix = &v3[..expected_prefix.len()];
    assert_eq!(prefix, expected_prefix);

    // Explicitly assert that omitted bytes do not appear in the project-info prefix.
    let syskind_record = {
        let mut buf = Vec::new();
        buf.extend_from_slice(&0x0001u16.to_le_bytes());
        buf.extend_from_slice(&4u32.to_le_bytes());
        buf.extend_from_slice(&syskind.to_le_bytes());
        buf
    };
    assert!(
        !contains_subslice(prefix, &syskind_record),
        "SysKind value bytes must not be present (only Id/Size are appended)"
    );
    let codepage_record = {
        let mut buf = Vec::new();
        buf.extend_from_slice(&0x0003u16.to_le_bytes());
        buf.extend_from_slice(&2u32.to_le_bytes());
        buf.extend_from_slice(&codepage.to_le_bytes());
        buf
    };
    assert!(
        !contains_subslice(prefix, &codepage_record),
        "CodePage value bytes must not be present (only Id/Size are appended)"
    );
    assert!(
        !contains_subslice(prefix, docstring),
        "DocString bytes must not be present (only lengths + reserved fields are appended)"
    );
    assert!(
        !contains_subslice(prefix, docstring_unicode),
        "DocStringUnicode bytes must not be present (only length + reserved fields are appended)"
    );
    assert!(
        !contains_subslice(prefix, helpfile1),
        "HelpFile1 bytes must not be present (only lengths + reserved fields are appended)"
    );
    assert!(
        !contains_subslice(prefix, helpfile2),
        "HelpFile2 bytes must not be present (only lengths + reserved fields are appended)"
    );
    let helpcontext_record = {
        let mut buf = Vec::new();
        buf.extend_from_slice(&0x0007u16.to_le_bytes());
        buf.extend_from_slice(&4u32.to_le_bytes());
        buf.extend_from_slice(&helpcontext.to_le_bytes());
        buf
    };
    assert!(
        !contains_subslice(prefix, &helpcontext_record),
        "HelpContext value bytes must not be present (only Id/Size are appended)"
    );
}

#[test]
fn v3_content_normalized_data_project_information_accepts_noncanonical_unicode_record_ids() {
    // Some producers use non-canonical record ids for project-info Unicode/alternate string variants
    // (e.g. 0x0041 instead of 0x0040). Ensure our v3 transcript builder accepts them.
    let vba_bin = build_project_with_project_info_records_only_noncanonical_unicode_ids();
    let v3 = v3_content_normalized_data(&vba_bin).expect("V3ContentNormalizedData");

    let project_name = b"MyV3Project";

    let syskind = 0xA1B2C3D4u32;
    let lcid = 0x11223344u32;
    let lcid_invoke = 0x55667788u32;
    let codepage = 0xCAFEu16;

    let docstring = b"__DOCSTRING_BYTES__";
    let docstring_unicode = b"D\0O\0C\0U\0N\0I\0";

    let helpfile1 = b"__HELPFILE1_PATH__";
    let helpfile2 = b"__HELPFILE2_PATH__";

    let helpcontext = 0x0BADF00Du32;

    let project_lib_flags = 0x01020304u32;

    let version_major = 0xCAFEBABEu32;
    let version_minor = 0xBEEFu16;

    let constants = b"ABC=1";
    let constants_unicode = b"A\x00B\x00C\x00=\x001\x00";

    let mut expected_prefix = Vec::new();

    // PROJECTSYSKIND: header only.
    expected_prefix.extend_from_slice(&0x0001u16.to_le_bytes());
    expected_prefix.extend_from_slice(&4u32.to_le_bytes());

    // PROJECTLCID: full record.
    expected_prefix.extend_from_slice(&0x0002u16.to_le_bytes());
    expected_prefix.extend_from_slice(&4u32.to_le_bytes());
    expected_prefix.extend_from_slice(&lcid.to_le_bytes());

    // PROJECTLCIDINVOKE: full record.
    expected_prefix.extend_from_slice(&0x0014u16.to_le_bytes());
    expected_prefix.extend_from_slice(&4u32.to_le_bytes());
    expected_prefix.extend_from_slice(&lcid_invoke.to_le_bytes());

    // PROJECTCODEPAGE: header only.
    expected_prefix.extend_from_slice(&0x0003u16.to_le_bytes());
    expected_prefix.extend_from_slice(&2u32.to_le_bytes());

    // PROJECTNAME: full record.
    expected_prefix.extend_from_slice(&0x0004u16.to_le_bytes());
    expected_prefix.extend_from_slice(&(project_name.len() as u32).to_le_bytes());
    expected_prefix.extend_from_slice(project_name);

    // PROJECTDOCSTRING: header only (with non-canonical Unicode marker id=0x0041).
    expected_prefix.extend_from_slice(&0x0005u16.to_le_bytes());
    expected_prefix.extend_from_slice(&(docstring.len() as u32).to_le_bytes());
    expected_prefix.extend_from_slice(&0x0041u16.to_le_bytes());
    expected_prefix.extend_from_slice(&(docstring_unicode.len() as u32).to_le_bytes());

    // PROJECTHELPFILEPATH: header only (with non-canonical HelpFile2 id=0x0042).
    expected_prefix.extend_from_slice(&0x0006u16.to_le_bytes());
    expected_prefix.extend_from_slice(&(helpfile1.len() as u32).to_le_bytes());
    expected_prefix.extend_from_slice(&0x0042u16.to_le_bytes());
    expected_prefix.extend_from_slice(&(helpfile2.len() as u32).to_le_bytes());

    // PROJECTHELPCONTEXT: header only.
    expected_prefix.extend_from_slice(&0x0007u16.to_le_bytes());
    expected_prefix.extend_from_slice(&4u32.to_le_bytes());

    // PROJECTLIBFLAGS: full record.
    expected_prefix.extend_from_slice(&0x0008u16.to_le_bytes());
    expected_prefix.extend_from_slice(&4u32.to_le_bytes());
    expected_prefix.extend_from_slice(&project_lib_flags.to_le_bytes());

    // PROJECTVERSION: full fixed-length record.
    expected_prefix.extend_from_slice(&0x0009u16.to_le_bytes());
    expected_prefix.extend_from_slice(&0x00000004u32.to_le_bytes()); // Reserved
    expected_prefix.extend_from_slice(&version_major.to_le_bytes());
    expected_prefix.extend_from_slice(&version_minor.to_le_bytes());

    // PROJECTCONSTANTS: full record including Unicode sub-record (non-canonical id=0x0043).
    expected_prefix.extend_from_slice(&0x000Cu16.to_le_bytes());
    expected_prefix.extend_from_slice(&(constants.len() as u32).to_le_bytes());
    expected_prefix.extend_from_slice(constants);
    expected_prefix.extend_from_slice(&0x0043u16.to_le_bytes());
    expected_prefix.extend_from_slice(&(constants_unicode.len() as u32).to_le_bytes());
    expected_prefix.extend_from_slice(constants_unicode);

    assert!(
        v3.len() >= expected_prefix.len(),
        "expected V3ContentNormalizedData to be at least {} bytes, got {}",
        expected_prefix.len(),
        v3.len()
    );
    let prefix = &v3[..expected_prefix.len()];
    assert_eq!(prefix, expected_prefix);

    // Omitted bytes must not appear.
    let syskind_record = {
        let mut buf = Vec::new();
        buf.extend_from_slice(&0x0001u16.to_le_bytes());
        buf.extend_from_slice(&4u32.to_le_bytes());
        buf.extend_from_slice(&syskind.to_le_bytes());
        buf
    };
    assert!(
        !contains_subslice(prefix, &syskind_record),
        "SysKind value bytes must not be present (only Id/Size are appended)"
    );
    let codepage_record = {
        let mut buf = Vec::new();
        buf.extend_from_slice(&0x0003u16.to_le_bytes());
        buf.extend_from_slice(&2u32.to_le_bytes());
        buf.extend_from_slice(&codepage.to_le_bytes());
        buf
    };
    assert!(
        !contains_subslice(prefix, &codepage_record),
        "CodePage value bytes must not be present (only Id/Size are appended)"
    );
    assert!(
        !contains_subslice(prefix, docstring),
        "DocString bytes must not be present (only lengths + reserved fields are appended)"
    );
    assert!(
        !contains_subslice(prefix, docstring_unicode),
        "DocStringUnicode bytes must not be present (only length + reserved fields are appended)"
    );
    assert!(
        !contains_subslice(prefix, helpfile1),
        "HelpFile1 bytes must not be present (only lengths + reserved fields are appended)"
    );
    assert!(
        !contains_subslice(prefix, helpfile2),
        "HelpFile2 bytes must not be present (only lengths + reserved fields are appended)"
    );
    let helpcontext_record = {
        let mut buf = Vec::new();
        buf.extend_from_slice(&0x0007u16.to_le_bytes());
        buf.extend_from_slice(&4u32.to_le_bytes());
        buf.extend_from_slice(&helpcontext.to_le_bytes());
        buf
    };
    assert!(
        !contains_subslice(prefix, &helpcontext_record),
        "HelpContext value bytes must not be present (only Id/Size are appended)"
    );
}

#[test]
fn project_normalized_data_v3_appends_padded_forms_normalized_data_when_designer_present() {
    let vba_bin = build_project_with_designer_storage();

    let content_v3 = v3_content_normalized_data(&vba_bin).expect("V3ContentNormalizedData");
    let forms = forms_normalized_data(&vba_bin).expect("FormsNormalizedData");
    let project =
        project_normalized_data_v3_transcript(&vba_bin).expect("ProjectNormalizedData v3");

    let mut expected_content_v3 = Vec::new();
    expected_content_v3.extend_from_slice(b"Sub FormHello()\nEnd Sub\n\nUserForm1\n");
    assert_eq!(content_v3, expected_content_v3);

    let mut expected_forms = Vec::new();
    expected_forms.extend_from_slice(b"ABC");
    expected_forms.extend(std::iter::repeat_n(0u8, 1020));
    assert_eq!(forms, expected_forms);

    // ProjectNormalizedData v3 includes filtered PROJECT stream properties before the v3 dir/module
    // transcript.
    let expected_project_prefix = b"Name=\"VBAProject\"\r\nBaseClass=\"UserForm1\"\r\n".to_vec();
    let expected_project = [
        expected_project_prefix.as_slice(),
        expected_content_v3.as_slice(),
        expected_forms.as_slice(),
    ]
    .concat();
    assert_eq!(project, expected_project);
}

#[test]
fn v3_content_normalized_data_skips_projectcompatversion_record() {
    // MS-OVBA §2.4.2.5: PROJECTCOMPATVERSION (0x004A) is not appended to the V3 transcript, but it
    // appears in many real-world `VBA/dir` streams. Ensure its presence does not perturb parsing or
    // hashing.
    let (vba_without, compat_version) = build_project_with_projectcompatversion(false);
    let (vba_with, _compat_version2) = build_project_with_projectcompatversion(true);

    let normalized_without =
        v3_content_normalized_data(&vba_without).expect("V3ContentNormalizedData without compat");
    let normalized_with =
        v3_content_normalized_data(&vba_with).expect("V3ContentNormalizedData with compat");

    assert_eq!(
        normalized_without, normalized_with,
        "PROJECTCOMPATVERSION (0x004A) must not affect V3ContentNormalizedData"
    );

    // Contents Hash v3 (the actual digest used by DigitalSignatureExt) should also be unaffected.
    let digest_without = contents_hash_v3(&vba_without).expect("ContentsHash v3 without compat");
    let digest_with = contents_hash_v3(&vba_with).expect("ContentsHash v3 with compat");
    assert_eq!(
        digest_without, digest_with,
        "PROJECTCOMPATVERSION (0x004A) must not affect ContentsHash v3"
    );

    // Sanity check: ensure module bytes are still present so this test is sensitive to parsing
    // misalignment (e.g. if skipping the record broke subsequent module parsing).
    assert!(
        contains_subslice(&normalized_without, b"Sub Foo()"),
        "expected V3ContentNormalizedData to include module source bytes"
    );

    // Regression assertion: the compat version payload bytes must not be present in the output.
    assert!(
        !normalized_without
            .windows(compat_version.len())
            .any(|w| w == compat_version),
        "V3ContentNormalizedData must skip PROJECTCOMPATVERSION payload bytes"
    );

    // Stronger assertion: the full record header+payload must not appear contiguously in the
    // transcript either.
    let compat_record = [
        0x004Au16.to_le_bytes().as_slice(),
        4u32.to_le_bytes().as_slice(),
        compat_version.as_slice(),
    ]
    .concat();
    assert!(
        !contains_subslice(&normalized_without, &compat_record),
        "V3ContentNormalizedData must not include PROJECTCOMPATVERSION record bytes"
    );
}

#[test]
fn v3_content_normalized_data_resolves_module_stream_name_from_unicode_record_variant() {
    let vba_bin = build_project_unicode_only_module_stream_name();
    let v3 = v3_content_normalized_data(&vba_bin).expect("V3ContentNormalizedData");

    let mut expected = Vec::new();
    // TypeRecord.Id (0x0021) + Reserved (0)
    expected.extend_from_slice(&0x0021u16.to_le_bytes());
    expected.extend_from_slice(&0u16.to_le_bytes());
    // LF-normalized source + trailing module name (Unicode bytes) + LF.
    expected.extend_from_slice(b"Sub Foo()\nEnd Sub\n\n");
    expected.extend_from_slice(&utf16le_bytes("模块名"));
    expected.push(b'\n');

    assert_eq!(v3, expected);
}

#[test]
fn project_normalized_data_v3_includes_project_properties_and_resolves_unicode_module_stream_name()
{
    let vba_bin = build_project_unicode_only_module_stream_name_with_project_stream();
    let normalized =
        project_normalized_data_v3_transcript(&vba_bin).expect("ProjectNormalizedData v3");

    let mut expected = Vec::new();
    // Filtered PROJECT stream properties (CRLF normalized).
    expected.extend_from_slice(b"Name=\"VBAProject\"\r\n");
    // Then V3ContentNormalizedData for the single module.
    expected.extend_from_slice(&0x0021u16.to_le_bytes());
    expected.extend_from_slice(&0u16.to_le_bytes());
    expected.extend_from_slice(b"Sub Foo()\nEnd Sub\n\n");
    expected.extend_from_slice(&utf16le_bytes("模块名"));
    expected.push(b'\n');

    assert_eq!(normalized, expected);
}

#[test]
fn v3_content_normalized_data_resolves_len_prefixed_module_stream_name_unicode_record() {
    let vba_bin = build_project_unicode_module_stream_name_with_internal_len_prefix();
    let v3 = v3_content_normalized_data(&vba_bin).expect("V3ContentNormalizedData");

    let mut expected = Vec::new();
    // TypeRecord.Id (0x0021) + Reserved (0)
    expected.extend_from_slice(&0x0021u16.to_le_bytes());
    expected.extend_from_slice(&0u16.to_le_bytes());
    // LF-normalized source + trailing module name (Unicode bytes) + LF.
    expected.extend_from_slice(b"Sub Foo()\nEnd Sub\n\n");
    expected.extend_from_slice(&utf16le_bytes("模块名"));
    expected.push(b'\n');

    assert_eq!(v3, expected);
}

#[test]
fn v3_content_normalized_data_prefers_unicode_module_stream_name_record_when_both_present() {
    let vba_bin = build_project_with_ansi_and_unicode_module_stream_name_records();
    let v3 = v3_content_normalized_data(&vba_bin).expect("V3ContentNormalizedData");

    let mut expected = Vec::new();
    expected.extend_from_slice(&0x0021u16.to_le_bytes());
    expected.extend_from_slice(&0u16.to_le_bytes());
    expected.extend_from_slice(b"Sub Foo()\nEnd Sub\n\nModule1\n");

    assert_eq!(v3, expected);
}

fn build_project_with_modulestreamname_unicode_marker_0048() -> Vec<u8> {
    // Keep module source already in normalized form to make expected bytes simple.
    let module_code = b"Sub Foo()\r\nEnd Sub\r\n";
    let module_container = compress_container(module_code);

    // Unicode stream name we will actually create in the OLE.
    let module_stream_name_unicode = "模块1";

    let dir_decompressed = {
        let mut out = Vec::new();

        // MODULENAME (ANSI)
        push_record(&mut out, 0x0019, b"Module1");

        // MODULESTREAMNAME using an alternate marker value (0x0048) for the Unicode stream name.
        //
        // This is not the most common encoding (0x0032 is typical), but it is observed in some
        // projects and v3 transcript logic must still resolve the Unicode stream name for OLE
        // lookup.
        let stream_name_mbcs = b"Wrong";
        let stream_name_unicode = utf16le_bytes(module_stream_name_unicode);
        out.extend_from_slice(&0x001Au16.to_le_bytes());
        out.extend_from_slice(&(stream_name_mbcs.len() as u32).to_le_bytes());
        out.extend_from_slice(stream_name_mbcs);
        out.extend_from_slice(&0x0048u16.to_le_bytes());
        out.extend_from_slice(&(stream_name_unicode.len() as u32).to_le_bytes());
        out.extend_from_slice(&stream_name_unicode);

        // MODULETYPE (procedural; TypeRecord.Id=0x0021)
        push_record(&mut out, 0x0021, &0u16.to_le_bytes());
        push_record(&mut out, 0x0031, &0u32.to_le_bytes()); // MODULETEXTOFFSET (0)
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
            .create_stream(&format!("VBA/{module_stream_name_unicode}"))
            .expect("module stream");
        s.write_all(&module_container).expect("write module");
    }

    ole.into_inner().into_inner()
}

#[test]
fn v3_content_normalized_data_prefers_modulenameunicode_payload_bytes_when_present() {
    let vba_bin = build_project_with_ansi_and_unicode_module_name_records();
    let v3 = v3_content_normalized_data(&vba_bin).expect("V3ContentNormalizedData");

    let mut expected = Vec::new();
    expected.extend_from_slice(&0x0021u16.to_le_bytes());
    expected.extend_from_slice(&0u16.to_le_bytes());
    expected.extend_from_slice(b"Sub Foo()\nEnd Sub\n\n");
    expected.extend_from_slice(&utf16le_bytes("Module1"));
    expected.push(b'\n');

    assert_eq!(v3, expected);
}

#[test]
fn v3_content_normalized_data_strips_internal_len_prefix_in_modulenameunicode_payload() {
    let vba_bin = build_project_with_modulenameunicode_internal_len_prefix();
    let v3 = v3_content_normalized_data(&vba_bin).expect("V3ContentNormalizedData");

    let mut expected = Vec::new();
    expected.extend_from_slice(&0x0021u16.to_le_bytes());
    expected.extend_from_slice(&0u16.to_le_bytes());
    expected.extend_from_slice(b"Sub Foo()\nEnd Sub\n\n");
    // The internal u32 length prefix must NOT be present in the transcript.
    expected.extend_from_slice(&utf16le_bytes("Module1"));
    expected.push(b'\n');

    assert_eq!(v3, expected);
}

#[test]
fn v3_content_normalized_data_strips_internal_len_prefix_excluding_trailing_nul_in_modulenameunicode_payload_byte_count(
) {
    let vba_bin = build_project_with_modulenameunicode_internal_len_prefix_excluding_trailing_nul_byte_count();
    let v3 = v3_content_normalized_data(&vba_bin).expect("V3ContentNormalizedData");

    let mut expected = Vec::new();
    expected.extend_from_slice(&0x0021u16.to_le_bytes());
    expected.extend_from_slice(&0u16.to_le_bytes());
    expected.extend_from_slice(b"Sub Foo()\nEnd Sub\n\n");
    // The internal u32 length prefix and the trailing UTF-16 NUL terminator must NOT be present in
    // the transcript.
    expected.extend_from_slice(&utf16le_bytes("Module1"));
    expected.push(b'\n');

    assert_eq!(v3, expected);
}

#[test]
fn v3_content_normalized_data_strips_internal_len_prefix_excluding_trailing_nul_in_modulenameunicode_payload_code_units(
) {
    let vba_bin = build_project_with_modulenameunicode_internal_len_prefix_excluding_trailing_nul_code_units();
    let v3 = v3_content_normalized_data(&vba_bin).expect("V3ContentNormalizedData");

    let mut expected = Vec::new();
    expected.extend_from_slice(&0x0021u16.to_le_bytes());
    expected.extend_from_slice(&0u16.to_le_bytes());
    expected.extend_from_slice(b"Sub Foo()\nEnd Sub\n\n");
    // The internal u32 length prefix and the trailing UTF-16 NUL terminator must NOT be present in
    // the transcript.
    expected.extend_from_slice(&utf16le_bytes("Module1"));
    expected.push(b'\n');

    assert_eq!(v3, expected);
}

#[test]
fn v3_content_normalized_data_resolves_modulestreamname_unicode_marker_0048() {
    let vba_bin = build_project_with_modulestreamname_unicode_marker_0048();
    let v3 = v3_content_normalized_data(&vba_bin).expect("V3ContentNormalizedData");

    let mut expected = Vec::new();
    expected.extend_from_slice(&0x0021u16.to_le_bytes());
    expected.extend_from_slice(&0u16.to_le_bytes());
    expected.extend_from_slice(b"Sub Foo()\nEnd Sub\n\nModule1\n");

    assert_eq!(v3, expected);
}
