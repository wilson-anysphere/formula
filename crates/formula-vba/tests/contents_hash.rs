use std::io::{Cursor, Write};

use encoding_rs::{WINDOWS_1251, WINDOWS_1252};
use formula_vba::{
    agile_content_hash_md5, compress_container, content_hash_md5, content_normalized_data,
};
use md5::{Digest as _, Md5};

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

fn build_two_module_project(module_order: [&str; 2]) -> Vec<u8> {
    // Distinct module source so we can assert byte ordering in ContentNormalizedData directly.
    let module_a_code = b"'MODULE-A\r\nSub A()\r\nEnd Sub\r\n";
    let module_b_code = b"'MODULE-B\r\nSub B()\r\nEnd Sub\r\n";

    let module_a_container = compress_container(module_a_code);
    let module_b_container = compress_container(module_b_code);

    // Build a minimal decompressed `VBA/dir` stream that lists the modules in the
    // desired order. The critical part for this test is the order of module records:
    // MODULENAME (0x0019) starts each module record group.
    let dir_decompressed = {
        let mut out = Vec::new();
        // PROJECTNAME (harmless here).
        push_record(&mut out, 0x0004, b"VBAProject");

        for name in module_order {
            push_record(&mut out, 0x0019, name.as_bytes()); // MODULENAME

            // MODULESTREAMNAME + reserved u16.
            let mut stream_name = Vec::new();
            stream_name.extend_from_slice(name.as_bytes());
            stream_name.extend_from_slice(&0u16.to_le_bytes());
            push_record(&mut out, 0x001A, &stream_name);

            // MODULETYPE (standard)
            push_record(&mut out, 0x0021, &0u16.to_le_bytes());
            // MODULETEXTOFFSET: our module stream is just the compressed container.
            push_record(&mut out, 0x0031, &0u32.to_le_bytes());
        }

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

    // Minimal OLE layout: `VBA/dir` + the module streams. Create the module streams in alphabetical
    // order (A then B) to ensure the tested ordering comes from `VBA/dir`, not OLE insertion order.
    {
        let mut s = ole.create_stream("VBA/ModuleA").expect("module A stream");
        s.write_all(&module_a_container).expect("write module A");
    }
    {
        let mut s = ole.create_stream("VBA/ModuleB").expect("module B stream");
        s.write_all(&module_b_container).expect("write module B");
    }

    ole.into_inner().into_inner()
}

fn build_single_module_project_with_dir_prelude(dir_prelude: &[(u16, &[u8])], module_code: &[u8]) -> Vec<u8> {
    let module_container = compress_container(module_code);

    let dir_decompressed = {
        let mut out = Vec::new();
        for (id, data) in dir_prelude {
            push_record(&mut out, *id, data);
        }

        // Minimal module record group.
        push_record(&mut out, 0x0019, b"Module1"); // MODULENAME

        // MODULESTREAMNAME + reserved u16.
        let mut stream_name = Vec::new();
        stream_name.extend_from_slice(b"Module1");
        stream_name.extend_from_slice(&0u16.to_le_bytes());
        push_record(&mut out, 0x001A, &stream_name);

        // MODULETYPE (standard) + MODULETEXTOFFSET (0: stream starts with a compressed container).
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
        let mut s = ole.create_stream("VBA/Module1").expect("module stream");
        s.write_all(&module_container).expect("write module");
    }

    ole.into_inner().into_inner()
}

#[test]
fn content_normalized_data_uses_module_record_order_from_dir_stream() {
    // Deliberately non-alphabetical order: B then A.
    let vba_bin = build_two_module_project(["ModuleB", "ModuleA"]);
    let normalized = content_normalized_data(&vba_bin).expect("content normalized data");

    let module_a_code = b"'MODULE-A\r\nSub A()\r\nEnd Sub\r\n";
    let module_b_code = b"'MODULE-B\r\nSub B()\r\nEnd Sub\r\n";

    let pos_b = find_subslice(&normalized, module_b_code).expect("ModuleB code should be present");
    let pos_a = find_subslice(&normalized, module_a_code).expect("ModuleA code should be present");
    assert!(
        pos_b < pos_a,
        "expected ModuleB bytes to appear before ModuleA bytes in ContentNormalizedData"
    );

    // Swapping module order in the `dir` stream should swap the order in the normalized data too.
    let vba_bin_swapped = build_two_module_project(["ModuleA", "ModuleB"]);
    let normalized_swapped =
        content_normalized_data(&vba_bin_swapped).expect("content normalized data");

    assert_ne!(
        normalized, normalized_swapped,
        "changing module stored order should change ContentNormalizedData"
    );

    let pos_a2 =
        find_subslice(&normalized_swapped, module_a_code).expect("ModuleA code should be present");
    let pos_b2 =
        find_subslice(&normalized_swapped, module_b_code).expect("ModuleB code should be present");
    assert!(
        pos_a2 < pos_b2,
        "expected ModuleA bytes to appear before ModuleB bytes when dir order is A then B"
    );
}

#[test]
fn content_normalized_data_decodes_cyrillic_module_stream_name_using_windows_1251() {
    // A Cyrillic module stream name encoded with Windows-1251. This is not valid UTF-8, so
    // `String::from_utf8_lossy` would corrupt it and fail to locate the matching OLE stream.
    let stream_name = "Привет"; // "hello" in Russian
    let (stream_name_bytes, _, _) = WINDOWS_1251.encode(stream_name);

    let module_code = "Sub Hello()\r\n'привет\r\nEnd Sub\r\n";
    let (module_code_bytes, _, _) = WINDOWS_1251.encode(module_code);
    let module_container = compress_container(module_code_bytes.as_ref());

    let dir_decompressed = {
        let mut out = Vec::new();

        // PROJECTCODEPAGE (u16 LE)
        push_record(&mut out, 0x0003, &1251u16.to_le_bytes());

        // Module records.
        push_record(&mut out, 0x0019, stream_name_bytes.as_ref()); // MODULENAME

        // MODULESTREAMNAME + reserved u16.
        let mut stream_name_record = Vec::new();
        stream_name_record.extend_from_slice(stream_name_bytes.as_ref());
        stream_name_record.extend_from_slice(&0u16.to_le_bytes());
        push_record(&mut out, 0x001A, &stream_name_record);

        // MODULETYPE (standard)
        push_record(&mut out, 0x0021, &0u16.to_le_bytes());
        // MODULETEXTOFFSET (0: stream starts with a compressed container)
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
        let path = format!("VBA/{stream_name}");
        let mut s = ole.create_stream(&path).expect("module stream");
        s.write_all(&module_container).expect("write module");
    }

    let vba_bin = ole.into_inner().into_inner();
    let normalized = content_normalized_data(&vba_bin).expect("ContentNormalizedData");

    // No reference records, so ContentNormalizedData should consist of the normalized module bytes.
    assert_eq!(normalized, module_code_bytes.as_ref());
}

#[test]
fn content_normalized_data_reference_records_registered_and_project() {
    // Build a decompressed `VBA/dir` stream that contains two REFERENCE records:
    // - 0x000D (REFERENCEREGISTERED)
    // - 0x000E (REFERENCEPROJECT)
    let dir_decompressed = {
        let mut out = Vec::new();

        // 0x000D (REFERENCEREGISTERED): use a libid that begins with '{' (0x7B).
        push_record(&mut out, 0x000D, b"{REG}");

        // 0x000E (REFERENCEPROJECT): two u32-len-prefixed strings + major(u32) + minor(u16).
        //
        // Choose version numbers so the little-endian representation contains a NUL byte early:
        // major=1 => 0x01 0x00 0x00 0x00
        // The MS-OVBA pseudocode copies bytes from a TempBuffer until the first NUL byte, so this
        // should stop immediately after copying the low byte (0x01) of `major`.
        let libid_absolute = b"ProjLib";
        let libid_relative = b"";
        let major: u32 = 1;
        let minor: u16 = 0;

        let mut reference_project = Vec::new();
        reference_project.extend_from_slice(&(libid_absolute.len() as u32).to_le_bytes());
        reference_project.extend_from_slice(libid_absolute);
        reference_project.extend_from_slice(&(libid_relative.len() as u32).to_le_bytes());
        reference_project.extend_from_slice(libid_relative);
        reference_project.extend_from_slice(&major.to_le_bytes());
        reference_project.extend_from_slice(&minor.to_le_bytes());
        push_record(&mut out, 0x000E, &reference_project);

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
    let vba_bin = ole.into_inner().into_inner();

    let normalized = content_normalized_data(&vba_bin).expect("ContentNormalizedData");

    // 0x000D case: ensure the libid bytes are included (starts with '{' = 0x7B).
    assert!(
        normalized.contains(&0x7B),
        "expected ContentNormalizedData to contain 0x7B ('{{') from REFERENCEREGISTERED"
    );

    // 0x000E case: manually-constructed expected byte vector based on MS-OVBA pseudocode:
    // TempBuffer = LibidAbsolute || LibidRelative || MajorVersion(u32le) || MinorVersion(u16le)
    // then copy bytes until the first 0x00 byte.
    //
    // With major=1 (01 00 00 00) and minor=0 (00 00), this yields:
    // LibidAbsolute + 0x01
    let expected_project = b"ProjLib\x01".to_vec();
    let expected_full = [b"{REG}".as_slice(), expected_project.as_slice()].concat();

    assert_eq!(normalized, expected_full);
}

#[test]
fn content_normalized_data_decodes_utf16le_module_stream_name_without_bom() {
    // Some producers store module stream-name strings in UTF-16LE form even when the project
    // codepage is an 8-bit encoding. Ensure the `ContentNormalizedData` builder uses the same
    // UTF-16LE heuristic as the main `DirStream` parser so module streams can still be located.
    let module_name = "Module1";
    let module_name_utf16le: Vec<u8> = module_name
        .encode_utf16()
        .flat_map(|u| u.to_le_bytes())
        .collect();

    let module_code = b"Option Explicit\r\n";
    let module_container = compress_container(module_code);

    let dir_decompressed = {
        let mut out = Vec::new();
        // PROJECTCODEPAGE (Windows-1252)
        push_record(&mut out, 0x0003, &1252u16.to_le_bytes());

        // Module record group: MODULENAME is ASCII, MODULESTREAMNAME is UTF-16LE.
        push_record(&mut out, 0x0019, module_name.as_bytes());
        let mut stream_name = Vec::new();
        stream_name.extend_from_slice(&module_name_utf16le);
        stream_name.extend_from_slice(&0u16.to_le_bytes());
        push_record(&mut out, 0x001A, &stream_name);
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
            .create_stream(format!("VBA/{module_name}"))
            .expect("module stream");
        s.write_all(&module_container).expect("write module");
    }
    let vba_bin = ole.into_inner().into_inner();

    let normalized = content_normalized_data(&vba_bin).expect("ContentNormalizedData");
    assert_eq!(normalized, module_code);
}

#[test]
fn content_normalized_data_includes_projectname_and_projectconstants_bytes_in_dir_order() {
    let module_code = b"Option Explicit\r\n";

    let vba_bin = build_single_module_project_with_dir_prelude(
        &[(0x0004, b"Proj"), (0x000C, b"Const")],
        module_code,
    );

    let normalized = content_normalized_data(&vba_bin).expect("ContentNormalizedData");
    let expected = [b"ProjConst".as_slice(), module_code].concat();
    assert_eq!(normalized, expected);
}

#[test]
fn content_normalized_data_preserves_project_record_order_for_projectname_and_constants() {
    let module_code = b"Option Explicit\r\n";

    // Same bytes as the previous test, but with PROJECTCONSTANTS appearing before PROJECTNAME in
    // the decompressed `VBA/dir` stream.
    let vba_bin = build_single_module_project_with_dir_prelude(
        &[(0x000C, b"Const"), (0x0004, b"Proj")],
        module_code,
    );

    let normalized = content_normalized_data(&vba_bin).expect("ContentNormalizedData");
    let expected = [b"ConstProj".as_slice(), module_code].concat();
    assert_eq!(normalized, expected);
}

#[test]
fn content_normalized_data_reference_project_includes_libid_relative_before_version_byte() {
    // Ensure `REFERENCEPROJECT` (0x000E) normalization concatenates:
    //   LibidAbsolute || LibidRelative || Major(u32le) || Minor(u16le)
    // and then copies bytes until the first 0x00.
    //
    // In particular, this test ensures LibidRelative is *not* skipped.
    let dir_decompressed = {
        let mut out = Vec::new();

        let libid_absolute = b"Abs";
        let libid_relative = b"Rel";
        let major: u32 = 1; // 01 00 00 00 => stops after copying 0x01
        let minor: u16 = 0;

        let mut reference_project = Vec::new();
        reference_project.extend_from_slice(&(libid_absolute.len() as u32).to_le_bytes());
        reference_project.extend_from_slice(libid_absolute);
        reference_project.extend_from_slice(&(libid_relative.len() as u32).to_le_bytes());
        reference_project.extend_from_slice(libid_relative);
        reference_project.extend_from_slice(&major.to_le_bytes());
        reference_project.extend_from_slice(&minor.to_le_bytes());
        push_record(&mut out, 0x000E, &reference_project);

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
    let vba_bin = ole.into_inner().into_inner();

    let normalized = content_normalized_data(&vba_bin).expect("ContentNormalizedData");
    assert_eq!(normalized, b"AbsRel\x01".to_vec());
}

#[test]
fn content_normalized_data_reference_project_copies_entire_tempbuffer_when_no_nul_bytes() {
    // REFERENCEPROJECT normalization copies bytes from:
    //   LibidAbsolute || LibidRelative || Major(u32le) || Minor(u16le)
    // until it encounters the first 0x00.
    //
    // When none of these fields contain a NUL byte, the entire TempBuffer should be copied.
    let dir_decompressed = {
        let mut out = Vec::new();

        let libid_absolute = b"Abs";
        let libid_relative = b"Rel";
        let major: u32 = 0x11223344; // bytes: 44 33 22 11 (no NUL)
        let minor: u16 = 0x5566; // bytes: 66 55 (no NUL)

        let mut reference_project = Vec::new();
        reference_project.extend_from_slice(&(libid_absolute.len() as u32).to_le_bytes());
        reference_project.extend_from_slice(libid_absolute);
        reference_project.extend_from_slice(&(libid_relative.len() as u32).to_le_bytes());
        reference_project.extend_from_slice(libid_relative);
        reference_project.extend_from_slice(&major.to_le_bytes());
        reference_project.extend_from_slice(&minor.to_le_bytes());
        push_record(&mut out, 0x000E, &reference_project);

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
    let vba_bin = ole.into_inner().into_inner();

    let normalized = content_normalized_data(&vba_bin).expect("ContentNormalizedData");

    let expected = [
        b"Abs".as_slice(),
        b"Rel".as_slice(),
        &[0x44, 0x33, 0x22, 0x11],
        &[0x66, 0x55],
    ]
    .concat();
    assert_eq!(normalized, expected);
}

#[test]
fn content_normalized_data_reference_project_ignores_trailing_bytes_after_version_fields() {
    // Ensure we only incorporate the fields described by the MS-OVBA pseudocode:
    //   LibidAbsolute || LibidRelative || MajorVersion || MinorVersion
    // and ignore any extra bytes that may appear in the record payload.
    let dir_decompressed = {
        let mut out = Vec::new();

        let libid_absolute = b"Abs";
        let libid_relative = b"Rel";
        let major: u32 = 0x11223344; // bytes: 44 33 22 11 (no NUL)
        let minor: u16 = 0x5566; // bytes: 66 55 (no NUL)
        let trailing = b"TRAILING-BYTES";

        let mut reference_project = Vec::new();
        reference_project.extend_from_slice(&(libid_absolute.len() as u32).to_le_bytes());
        reference_project.extend_from_slice(libid_absolute);
        reference_project.extend_from_slice(&(libid_relative.len() as u32).to_le_bytes());
        reference_project.extend_from_slice(libid_relative);
        reference_project.extend_from_slice(&major.to_le_bytes());
        reference_project.extend_from_slice(&minor.to_le_bytes());
        reference_project.extend_from_slice(trailing);
        push_record(&mut out, 0x000E, &reference_project);

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
    let vba_bin = ole.into_inner().into_inner();

    let normalized = content_normalized_data(&vba_bin).expect("ContentNormalizedData");

    let expected = [
        b"Abs".as_slice(),
        b"Rel".as_slice(),
        &[0x44, 0x33, 0x22, 0x11],
        &[0x66, 0x55],
    ]
    .concat();
    assert_eq!(normalized, expected);
}

#[test]
fn content_normalized_data_module_newlines_and_attribute_stripping() {
    // Module source includes:
    // - Attribute lines (mixed case) that must be stripped (case-insensitive match)
    // - Attribute lines where `Attribute` is followed by a **tab** (whitespace) that must also be stripped
    // - CRLF, CR-only, and lone-LF line endings
    // - A non-attribute line containing the word "Attribute" (must be preserved)
    let module_code = concat!(
        "aTtRiBuTe\tVB_Name = \"Module1\"\r\n",
        "Option Explicit\r",
        "Print \"Attribute\"\n",
        "AtTrIbUtE VB_Base = \"0{00000000-0000-0000-0000-000000000000}\"\r\n",
        "Sub Foo()\r\n",
        "End Sub\r\n",
    );

    let module_container = compress_container(module_code.as_bytes());

    // Minimal `dir` stream describing a single module at offset 0.
    let dir_decompressed = {
        let mut out = Vec::new();
        push_record(&mut out, 0x0019, b"Module1");
        let mut stream_name = Vec::new();
        stream_name.extend_from_slice(b"Module1");
        stream_name.extend_from_slice(&0u16.to_le_bytes());
        push_record(&mut out, 0x001A, &stream_name);
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
        let mut s = ole.create_stream("VBA/Module1").expect("module stream");
        // MODULETEXTOFFSET is 0, so the stream starts with a compressed container.
        s.write_all(&module_container).expect("write module");
    }
    let vba_bin = ole.into_inner().into_inner();

    let normalized = content_normalized_data(&vba_bin).expect("ContentNormalizedData");

    // Only the non-Attribute lines should remain, each terminated with CRLF.
    let expected = concat!(
        "Option Explicit\r\n",
        "Print \"Attribute\"\r\n",
        "Sub Foo()\r\n",
        "End Sub\r\n",
    )
    .as_bytes()
    .to_vec();

    assert_eq!(normalized, expected);
}

#[test]
fn content_normalized_data_does_not_strip_attribute_with_leading_whitespace() {
    // Per MS-OVBA, only lines that begin with `Attribute` (start-of-line match) are stripped.
    // If the word is preceded by whitespace, it should be treated as a normal code line.
    let module_code = concat!(
        " Attribute VB_Name = \"Module1\"\r\n",
        "\tAttribute VB_Base = \"0{00000000-0000-0000-0000-000000000000}\"\r\n",
        "Attribute VB_Description = \"stripped\"\r\n",
        "Option Explicit\r\n",
    );

    let module_container = compress_container(module_code.as_bytes());

    let dir_decompressed = {
        let mut out = Vec::new();
        push_record(&mut out, 0x0019, b"Module1");
        let mut stream_name = Vec::new();
        stream_name.extend_from_slice(b"Module1");
        stream_name.extend_from_slice(&0u16.to_le_bytes());
        push_record(&mut out, 0x001A, &stream_name);
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
        let mut s = ole.create_stream("VBA/Module1").expect("module stream");
        s.write_all(&module_container).expect("write module");
    }
    let vba_bin = ole.into_inner().into_inner();

    let normalized = content_normalized_data(&vba_bin).expect("ContentNormalizedData");

    // The two leading-whitespace lines should remain; the true Attribute line should be stripped.
    let expected = concat!(
        " Attribute VB_Name = \"Module1\"\r\n",
        "\tAttribute VB_Base = \"0{00000000-0000-0000-0000-000000000000}\"\r\n",
        "Option Explicit\r\n",
    )
    .as_bytes()
    .to_vec();
    assert_eq!(normalized, expected);
}

#[test]
fn content_normalized_data_preserves_blank_lines() {
    // Blank (empty) lines are not Attribute lines, and should be preserved as CRLF in the
    // normalized output.
    let module_code = concat!(
        "Option Explicit\r\n",
        "\r\n",
        "Sub Foo()\r\n",
        "End Sub\r\n",
    );
    let module_container = compress_container(module_code.as_bytes());

    let dir_decompressed = {
        let mut out = Vec::new();
        push_record(&mut out, 0x0019, b"Module1");
        let mut stream_name = Vec::new();
        stream_name.extend_from_slice(b"Module1");
        stream_name.extend_from_slice(&0u16.to_le_bytes());
        push_record(&mut out, 0x001A, &stream_name);
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
        let mut s = ole.create_stream("VBA/Module1").expect("module stream");
        s.write_all(&module_container).expect("write module");
    }
    let vba_bin = ole.into_inner().into_inner();

    let normalized = content_normalized_data(&vba_bin).expect("ContentNormalizedData");
    assert_eq!(normalized, module_code.as_bytes());
}

#[test]
fn content_normalized_data_attribute_keyword_matching_is_exact() {
    // Ensure we only strip `Attribute` lines when the keyword is followed by whitespace (or end-of-line),
    // not when it's merely a prefix like `AttributeX`.
    //
    // Also ensure a line that is exactly `Attribute` is treated as an Attribute line and stripped.
    let module_code = concat!(
        "AttributeX VB_Name = \"Module1\"\r\n",
        "ATTRIBUTE\r\n",
        "Option Explicit\r\n",
    );

    let module_container = compress_container(module_code.as_bytes());

    // Minimal `dir` stream describing a single module at offset 0.
    let dir_decompressed = {
        let mut out = Vec::new();
        push_record(&mut out, 0x0019, b"Module1");
        let mut stream_name = Vec::new();
        stream_name.extend_from_slice(b"Module1");
        stream_name.extend_from_slice(&0u16.to_le_bytes());
        push_record(&mut out, 0x001A, &stream_name);
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
        let mut s = ole.create_stream("VBA/Module1").expect("module stream");
        // MODULETEXTOFFSET is 0, so the stream starts with a compressed container.
        s.write_all(&module_container).expect("write module");
    }
    let vba_bin = ole.into_inner().into_inner();

    let normalized = content_normalized_data(&vba_bin).expect("ContentNormalizedData");

    let expected = concat!(
        "AttributeX VB_Name = \"Module1\"\r\n",
        "Option Explicit\r\n",
    )
    .as_bytes()
    .to_vec();

    assert_eq!(normalized, expected);
}

#[test]
fn content_normalized_data_uses_module_name_when_module_stream_name_record_is_missing() {
    // Some broken/handcrafted `VBA/dir` streams may omit MODULESTREAMNAME (0x001A).
    // In that case, we should fall back to MODULENAME for stream lookup.
    let module_code = b"Option Explicit\r\n";
    let module_container = compress_container(module_code);

    let dir_decompressed = {
        let mut out = Vec::new();
        push_record(&mut out, 0x0019, b"Module1"); // MODULENAME only (no 0x001A)
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
        let mut s = ole.create_stream("VBA/Module1").expect("module stream");
        s.write_all(&module_container).expect("write module");
    }
    let vba_bin = ole.into_inner().into_inner();

    let normalized = content_normalized_data(&vba_bin).expect("ContentNormalizedData");
    assert_eq!(normalized, module_code);
}

#[test]
fn content_normalized_data_uses_module_stream_name_record_for_stream_lookup() {
    // In MS-OVBA, MODULENAME (0x0019) and MODULESTREAMNAME (0x001A) are distinct records.
    // ContentNormalizedData must read module source bytes from the module's *stream name*
    // (0x001A), not the display/module name (0x0019).
    let module_code = "Option Explicit\r\n";
    let module_container = compress_container(module_code.as_bytes());

    let dir_decompressed = {
        let mut out = Vec::new();

        // MODULENAME: not the actual stream name.
        push_record(&mut out, 0x0019, b"NiceName");

        // MODULESTREAMNAME + reserved u16.
        let mut stream_name = Vec::new();
        stream_name.extend_from_slice(b"Stream1");
        stream_name.extend_from_slice(&0u16.to_le_bytes());
        push_record(&mut out, 0x001A, &stream_name);

        // MODULETYPE + MODULETEXTOFFSET (0: stream begins with compressed container)
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
        // Create only the MODULESTREAMNAME stream (not "VBA/NiceName").
        let mut s = ole.create_stream("VBA/Stream1").expect("module stream");
        s.write_all(&module_container).expect("write module");
    }
    let vba_bin = ole.into_inner().into_inner();

    let normalized = content_normalized_data(&vba_bin).expect("ContentNormalizedData");
    assert_eq!(normalized, module_code.as_bytes());
}
#[test]
fn content_normalized_data_includes_project_name_and_constants() {
    // Attribute line should be stripped, and the remaining module lines normalized to CRLF.
    let module_code = b"Attribute VB_Name = \"Module1\"\nSub A()\nEnd Sub\n";
    let module_container = compress_container(module_code);

    let dir_decompressed = {
        let mut out = Vec::new();
        push_record(&mut out, 0x0004, b"P"); // PROJECTNAME
        push_record(&mut out, 0x000C, b"C"); // PROJECTCONSTANTS

        push_record(&mut out, 0x0019, b"Module1"); // MODULENAME
        let mut stream_name = Vec::new();
        stream_name.extend_from_slice(b"Module1");
        stream_name.extend_from_slice(&0u16.to_le_bytes());
        push_record(&mut out, 0x001A, &stream_name); // MODULESTREAMNAME
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
        let mut s = ole.create_stream("VBA/Module1").expect("module stream");
        s.write_all(&module_container).expect("write module");
    }
    let vba_bin = ole.into_inner().into_inner();

    let normalized = content_normalized_data(&vba_bin).expect("content normalized data");
    assert_eq!(&normalized[..2], b"PC", "expected project name + constants prefix");
    assert!(
        normalized.ends_with(b"Sub A()\r\nEnd Sub\r\n"),
        "expected module newlines to be normalized to CRLF: {normalized:?}"
    );
}

#[test]
fn content_normalized_data_decodes_module_stream_name_using_dir_codepage() {
    let module_name = "Módülé1";
    let module_code = "Option Explicit\r\n";
    let module_container = compress_container(module_code.as_bytes());

    let (module_name_bytes, _, _) = WINDOWS_1252.encode(module_name);

    let dir_decompressed = {
        let mut out = Vec::new();
        // PROJECTCODEPAGE (u16 LE)
        push_record(&mut out, 0x0003, &1252u16.to_le_bytes());
        // MODULENAME
        push_record(&mut out, 0x0019, module_name_bytes.as_ref());
        // MODULESTREAMNAME + reserved u16
        let mut stream_name = Vec::new();
        stream_name.extend_from_slice(module_name_bytes.as_ref());
        stream_name.extend_from_slice(&0u16.to_le_bytes());
        push_record(&mut out, 0x001A, &stream_name);
        // MODULETYPE + MODULETEXTOFFSET
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
        let stream_path = format!("VBA/{module_name}");
        let mut s = ole.create_stream(&stream_path).expect("module stream");
        s.write_all(&module_container).expect("write module");
    }
    let vba_bin = ole.into_inner().into_inner();

    let normalized = content_normalized_data(&vba_bin).expect("ContentNormalizedData");
    assert_eq!(normalized, module_code.as_bytes());
}

#[test]
fn content_normalized_data_does_not_mistake_two_byte_ascii_module_name_for_utf16le() {
    // Regression test for our UTF-16LE detection heuristic: a 2-byte MBCS/ASCII name like "AB"
    // must not be treated as UTF-16LE (which would decode as a single unrelated code unit and
    // break module stream lookup).
    let module_name = "AB";
    let module_code = "Option Explicit\r\n";
    let module_container = compress_container(module_code.as_bytes());

    let dir_decompressed = {
        let mut out = Vec::new();
        // PROJECTCODEPAGE (u16 LE)
        push_record(&mut out, 0x0003, &1252u16.to_le_bytes());

        // MODULENAME / MODULESTREAMNAME
        push_record(&mut out, 0x0019, module_name.as_bytes());
        let mut stream_name = Vec::new();
        stream_name.extend_from_slice(module_name.as_bytes());
        stream_name.extend_from_slice(&0u16.to_le_bytes());
        push_record(&mut out, 0x001A, &stream_name);

        // MODULETYPE + MODULETEXTOFFSET
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
        let stream_path = format!("VBA/{module_name}");
        let mut s = ole.create_stream(&stream_path).expect("module stream");
        s.write_all(&module_container).expect("write module");
    }
    let vba_bin = ole.into_inner().into_inner();

    let normalized = content_normalized_data(&vba_bin).expect("ContentNormalizedData");
    assert_eq!(normalized, module_code.as_bytes());
}

#[test]
fn content_normalized_data_finds_module_source_without_text_offset_using_signature_scan() {
    // Ensure we exercise the same "scan for compressed container signature" fallback that the
    // module parser uses when `MODULETEXTOFFSET` (0x0031) is absent.
    let module_code = "Option Explicit\r\n";
    let module_container = compress_container(module_code.as_bytes());

    // Prefix the module stream with some header bytes that should not be mistaken for a compressed
    // container signature.
    let mut module_stream = vec![0x01, 0x00, 0x00, 0x99, 0x99, 0x88, 0x77];
    module_stream.extend_from_slice(&module_container);

    let dir_decompressed = {
        let mut out = Vec::new();
        push_record(&mut out, 0x0019, b"Module1");
        let mut stream_name = Vec::new();
        stream_name.extend_from_slice(b"Module1");
        stream_name.extend_from_slice(&0u16.to_le_bytes());
        push_record(&mut out, 0x001A, &stream_name);
        push_record(&mut out, 0x0021, &0u16.to_le_bytes());
        // Intentionally omit MODULETEXTOFFSET (0x0031).
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
        s.write_all(&module_stream).expect("write module");
    }
    let vba_bin = ole.into_inner().into_inner();

    let normalized = content_normalized_data(&vba_bin).expect("ContentNormalizedData");
    assert_eq!(normalized, module_code.as_bytes());
}

#[test]
fn content_normalized_data_signature_scan_skips_invalid_container_candidates() {
    // When MODULETEXTOFFSET is missing, we fall back to a signature scan to locate the compressed
    // source container. That scan should not stop at the first “looks like a container” match if
    // decompression fails; it should keep searching for a later valid container.
    let module_code = "Option Explicit\r\n";
    let module_container = compress_container(module_code.as_bytes());

    // False CompressedContainer signature:
    // - 0x01 signature
    // - u16 header 0x3FFF (uncompressed, size_field=0x0FFF => expects 4096 bytes of chunk data)
    // This will fail to decompress because the buffer is much shorter than that.
    let mut module_stream = vec![0x01, 0xFF, 0x3F];
    module_stream.extend_from_slice(&module_container);

    let dir_decompressed = {
        let mut out = Vec::new();
        push_record(&mut out, 0x0019, b"Module1");
        let mut stream_name = Vec::new();
        stream_name.extend_from_slice(b"Module1");
        stream_name.extend_from_slice(&0u16.to_le_bytes());
        push_record(&mut out, 0x001A, &stream_name);
        push_record(&mut out, 0x0021, &0u16.to_le_bytes());
        // Intentionally omit MODULETEXTOFFSET (0x0031) so the signature scan is used.
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
        s.write_all(&module_stream).expect("write module");
    }
    let vba_bin = ole.into_inner().into_inner();

    let normalized = content_normalized_data(&vba_bin).expect("ContentNormalizedData");
    assert_eq!(normalized, module_code.as_bytes());
}

#[test]
fn content_hash_md5_matches_md5_of_content_normalized_data() {
    let vba_bin = build_two_module_project(["ModuleA", "ModuleB"]);
    let normalized = content_normalized_data(&vba_bin).expect("ContentNormalizedData");
    let expected: [u8; 16] = Md5::digest(&normalized).into();

    let got = content_hash_md5(&vba_bin).expect("Content Hash (MD5)");
    assert_eq!(got, expected);
}

#[test]
fn content_normalized_data_respects_module_text_offset_when_stream_has_prefix() {
    // Real module streams can contain a binary header prefix before the MS-OVBA CompressedContainer.
    // The module's MODULETEXTOFFSET / TextOffset record (0x0031) indicates where the compressed
    // source begins.
    let prefix_len = 20usize;
    let prefix = vec![0x00u8; prefix_len]; // avoid 0x01 so decompressing from offset 0 fails

    // Include an Attribute line so we also exercise source normalization (Attribute stripping).
    let module_code = concat!(
        "Attribute VB_Name = \"Module1\"\r\n",
        "Sub Hello()\r\n",
        "End Sub\r\n",
    );
    let module_container = compress_container(module_code.as_bytes());

    let mut module_stream = prefix.clone();
    module_stream.extend_from_slice(&module_container);

    let dir_decompressed = {
        let mut out = Vec::new();
        push_record(&mut out, 0x0019, b"Module1"); // MODULENAME
        let mut stream_name = Vec::new();
        stream_name.extend_from_slice(b"Module1");
        stream_name.extend_from_slice(&0u16.to_le_bytes());
        push_record(&mut out, 0x001A, &stream_name); // MODULESTREAMNAME
        push_record(&mut out, 0x0021, &0u16.to_le_bytes()); // MODULETYPE
        push_record(
            &mut out,
            0x0031, // MODULETEXTOFFSET / TextOffset
            &(prefix_len as u32).to_le_bytes(),
        );
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
        s.write_all(&module_stream).expect("write module");
    }
    let vba_bin = ole.into_inner().into_inner();

    let normalized = content_normalized_data(&vba_bin).expect("ContentNormalizedData");
    assert_eq!(normalized, b"Sub Hello()\r\nEnd Sub\r\n".to_vec());

    // If TextOffset is wrong, decompression should fail (proves we actually respect the 0x0031
    // record and don't just scan the stream for a compressed container signature).
    let dir_decompressed_wrong_offset = {
        let mut out = Vec::new();
        push_record(&mut out, 0x0019, b"Module1");
        let mut stream_name = Vec::new();
        stream_name.extend_from_slice(b"Module1");
        stream_name.extend_from_slice(&0u16.to_le_bytes());
        push_record(&mut out, 0x001A, &stream_name);
        push_record(&mut out, 0x0021, &0u16.to_le_bytes());
        push_record(&mut out, 0x0031, &0u32.to_le_bytes()); // wrong: should be prefix_len
        out
    };
    let dir_container_wrong_offset = compress_container(&dir_decompressed_wrong_offset);

    let cursor = Cursor::new(Vec::new());
    let mut ole = cfb::CompoundFile::create(cursor).expect("create cfb");
    ole.create_storage("VBA").expect("VBA storage");
    {
        let mut s = ole.create_stream("VBA/dir").expect("dir stream");
        s.write_all(&dir_container_wrong_offset)
            .expect("write dir");
    }
    {
        let mut s = ole.create_stream("VBA/Module1").expect("module stream");
        s.write_all(&module_stream).expect("write module");
    }
    let vba_bin_wrong_offset = ole.into_inner().into_inner();

    let err = content_normalized_data(&vba_bin_wrong_offset).expect_err("expected error");
    assert!(
        matches!(
            err,
            formula_vba::ParseError::Compression(formula_vba::CompressionError::InvalidSignature(
                0x00
            ))
        ),
        "unexpected error for wrong TextOffset: {err:?}"
    );
}

#[test]
fn content_normalized_data_respects_module_text_offset_even_if_header_looks_like_container() {
    // When MODULETEXTOFFSET is present, we should use it directly (rather than scanning for a
    // compressed-container signature).
    //
    // This test prefixes the module stream with bytes that *look like* a CompressedContainer start
    // (0x01 + chunk header with signature bits 0b011), but are intentionally incomplete/invalid so
    // that attempting to decompress from offset 0 would fail.
    let module_code = "Option Explicit\r\n";
    let module_container = compress_container(module_code.as_bytes());

    // False CompressedContainer signature (truncated chunk data):
    // - 0x01 signature
    // - u16 header 0x3000 (uncompressed, size_field=0 => expects 1 data byte after header)
    // - missing that data byte => decompress_container would error if started here
    let mut module_stream = vec![0x01, 0x00, 0x30];

    // Pad up to the declared text offset.
    let text_offset: u32 = 12;
    while module_stream.len() < text_offset as usize {
        module_stream.push(0xAA);
    }
    module_stream.extend_from_slice(&module_container);

    let dir_decompressed = {
        let mut out = Vec::new();
        push_record(&mut out, 0x0019, b"Module1");
        let mut stream_name = Vec::new();
        stream_name.extend_from_slice(b"Module1");
        stream_name.extend_from_slice(&0u16.to_le_bytes());
        push_record(&mut out, 0x001A, &stream_name);
        push_record(&mut out, 0x0021, &0u16.to_le_bytes());
        push_record(&mut out, 0x0031, &text_offset.to_le_bytes());
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
        s.write_all(&module_stream).expect("write module");
    }
    let vba_bin = ole.into_inner().into_inner();

    let normalized = content_normalized_data(&vba_bin).expect("ContentNormalizedData");
    assert_eq!(normalized, module_code.as_bytes());
}

#[test]
fn agile_content_hash_md5_matches_content_hash_when_no_designers_present() {
    // Build a minimal project that includes a PROJECT stream (required for FormsNormalizedData) but
    // no `BaseClass=` lines, so `FormsNormalizedData` is the empty byte sequence.
    let module_code = "Option Explicit\r\n";
    let module_container = compress_container(module_code.as_bytes());

    let dir_decompressed = {
        let mut out = Vec::new();
        push_record(&mut out, 0x0019, b"Module1");
        let mut stream_name = Vec::new();
        stream_name.extend_from_slice(b"Module1");
        stream_name.extend_from_slice(&0u16.to_le_bytes());
        push_record(&mut out, 0x001A, &stream_name);
        push_record(&mut out, 0x0021, &0u16.to_le_bytes());
        push_record(&mut out, 0x0031, &0u32.to_le_bytes());
        out
    };
    let dir_container = compress_container(&dir_decompressed);

    let cursor = Cursor::new(Vec::new());
    let mut ole = cfb::CompoundFile::create(cursor).expect("create cfb");
    ole.create_storage("VBA").expect("VBA storage");
    {
        let mut s = ole.create_stream("PROJECT").expect("PROJECT stream");
        s.write_all(b"Name=\"VBAProject\"\r\nModule=Module1\r\n")
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
    let vba_bin = ole.into_inner().into_inner();

    let content_hash = content_hash_md5(&vba_bin).expect("Content Hash");
    let agile_hash = agile_content_hash_md5(&vba_bin)
        .expect("Agile Content Hash computation")
        .expect("FormsNormalizedData should be available");
    assert_eq!(agile_hash, content_hash);
}
