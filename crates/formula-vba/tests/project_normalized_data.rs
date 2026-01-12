use std::io::{Cursor, Write};

use formula_vba::{
    compress_container, project_normalized_data, project_normalized_data_v3,
    project_normalized_data_v3_dir_records, DirParseError, ParseError,
};

fn push_record(out: &mut Vec<u8>, id: u16, data: &[u8]) {
    out.extend_from_slice(&id.to_le_bytes());
    out.extend_from_slice(&(data.len() as u32).to_le_bytes());
    out.extend_from_slice(data);
}

fn build_vba_bin_with_dir_decompressed(dir_decompressed: &[u8]) -> Vec<u8> {
    let dir_container = compress_container(dir_decompressed);

    let cursor = Cursor::new(Vec::new());
    let mut ole = cfb::CompoundFile::create(cursor).expect("create cfb");
    ole.create_storage("VBA").expect("VBA storage");
    {
        let mut s = ole.create_stream("VBA/dir").expect("dir stream");
        s.write_all(&dir_container).expect("write dir");
    }
    ole.into_inner().into_inner()
}

fn utf16le_bytes(s: &str) -> Vec<u8> {
    let mut out = Vec::new();
    for unit in s.encode_utf16() {
        out.extend_from_slice(&unit.to_le_bytes());
    }
    out
}

fn unicode_record_data(s: &str) -> Vec<u8> {
    let units: Vec<u16> = s.encode_utf16().collect();
    let mut out = Vec::with_capacity(4 + units.len() * 2);
    out.extend_from_slice(&(units.len() as u32).to_le_bytes());
    for u in units {
        out.extend_from_slice(&u.to_le_bytes());
    }
    out
}

#[test]
fn project_normalized_data_includes_expected_dir_records_and_prefers_unicode_variants() {
    // Build a synthetic decompressed `VBA/dir` stream with:
    // - multiple included project-info records
    // - one excluded record
    // - ANSI + UNICODE pairs where the algorithm must prefer the UNICODE record.
    let dir_decompressed = {
        let mut out = Vec::new();

        // Included: PROJECTSYSKIND
        push_record(&mut out, 0x0001, &1u32.to_le_bytes());
        // Included (optional): PROJECTCOMPATVERSION
        push_record(&mut out, 0x004A, &0xDEADBEEFu32.to_le_bytes());
        // Included: PROJECTLCID
        push_record(&mut out, 0x0002, &0x0409u32.to_le_bytes());
        // Included: PROJECTCODEPAGE
        push_record(&mut out, 0x0003, &1252u16.to_le_bytes());
        // Included: PROJECTNAME
        push_record(&mut out, 0x0004, b"MyProject");

        // Included (ANSI), but followed by UNICODE -> should be skipped in favor of UNICODE.
        push_record(&mut out, 0x0005, b"Doc");
        // Included: PROJECTDOCSTRINGUNICODE (paired with 0x0005 above).
        push_record(&mut out, 0x0040, &utf16le_bytes("Doc"));

        // Excluded: REFERENCEREGISTERED (0x000D)
        push_record(&mut out, 0x000D, b"{EXCLUDED}");

        // Included (ANSI), but followed by UNICODE -> should be skipped in favor of UNICODE.
        push_record(&mut out, 0x000C, b"Const=1");
        // Included: PROJECTCONSTANTSUNICODE (paired with 0x000C above).
        push_record(&mut out, 0x003C, &utf16le_bytes("Const=1"));

        out
    };

    let vba_bin = build_vba_bin_with_dir_decompressed(&dir_decompressed);
    let normalized = project_normalized_data(&vba_bin).expect("ProjectNormalizedData");

    let expected = [
        1u32.to_le_bytes().as_slice(),
        0xDEADBEEFu32.to_le_bytes().as_slice(),
        0x0409u32.to_le_bytes().as_slice(),
        1252u16.to_le_bytes().as_slice(),
        b"MyProject".as_slice(),
        utf16le_bytes("Doc").as_slice(),
        utf16le_bytes("Const=1").as_slice(),
    ]
    .concat();

    assert_eq!(normalized, expected);
}

#[test]
fn project_normalized_data_skips_projectcompatversion_record() {
    // Real-world `VBA/dir` streams often include PROJECTCOMPATVERSION (0x004A) in the
    // ProjectInformation record list. MS-OVBA does not include this record in the
    // ProjectNormalizedData transcript, so it must be safely skipped.
    let compat_version = 0xDEADBEEFu32.to_le_bytes();

    fn build_dir(include_compat: bool, compat_version: &[u8; 4]) -> Vec<u8> {
        let mut out = Vec::new();

        push_record(&mut out, 0x0001, &1u32.to_le_bytes()); // PROJECTSYSKIND
        push_record(&mut out, 0x0002, &0x0409u32.to_le_bytes()); // PROJECTLCID
        push_record(&mut out, 0x0014, &0x0409u32.to_le_bytes()); // PROJECTLCIDINVOKE
        push_record(&mut out, 0x0003, &1252u16.to_le_bytes()); // PROJECTCODEPAGE

        push_record(&mut out, 0x0004, b"VBAProject"); // PROJECTNAME
        push_record(&mut out, 0x0005, b"DocString"); // PROJECTDOCSTRING
        push_record(&mut out, 0x0006, b"C:\\help.chm"); // PROJECTHELPFILEPATH
        push_record(&mut out, 0x0007, &0u32.to_le_bytes()); // PROJECTHELPCONTEXT
        push_record(&mut out, 0x0008, &0u32.to_le_bytes()); // PROJECTLIBFLAGS

        let mut version = Vec::new();
        version.extend_from_slice(&1u16.to_le_bytes());
        version.extend_from_slice(&0u16.to_le_bytes());
        push_record(&mut out, 0x0009, &version); // PROJECTVERSION

        if include_compat {
            push_record(&mut out, 0x004A, compat_version); // PROJECTCOMPATVERSION
        }

        push_record(&mut out, 0x000C, b"Constants"); // PROJECTCONSTANTS

        out
    }

    let dir_without = build_dir(false, &compat_version);
    let dir_with = build_dir(true, &compat_version);

    let vba_without = build_vba_bin_with_dir_decompressed(&dir_without);
    let vba_with = build_vba_bin_with_dir_decompressed(&dir_with);

    let normalized_without =
        project_normalized_data(&vba_without).expect("ProjectNormalizedData without compat");
    let normalized_with =
        project_normalized_data(&vba_with).expect("ProjectNormalizedData with compat");

    assert_eq!(
        normalized_without, normalized_with,
        "PROJECTCOMPATVERSION (0x004A) must not affect ProjectNormalizedData"
    );
    assert!(
        find_subslice(&normalized_without, b"VBAProject").is_some(),
        "expected ProjectNormalizedData to include PROJECTNAME bytes"
    );
    assert!(
        !normalized_without
            .windows(compat_version.len())
            .any(|w| w == compat_version),
        "ProjectNormalizedData must skip PROJECTCOMPATVERSION payload bytes"
    );
}

#[test]
fn project_normalized_data_v3_missing_vba_dir_stream() {
    let cursor = Cursor::new(Vec::new());
    let mut ole = cfb::CompoundFile::create(cursor).expect("create cfb");
    ole.create_storage("VBA").expect("VBA storage");

    let vba_bin = ole.into_inner().into_inner();
    let err = project_normalized_data_v3(&vba_bin).expect_err("expected MissingStream");
    match err {
        ParseError::MissingStream("VBA/dir") => {}
        other => panic!("expected MissingStream(\"VBA/dir\"), got {other:?}"),
    }
}

#[test]
fn project_normalized_data_v3_dir_truncated_record_header() {
    // One valid record followed by <6 leftover bytes so the next record header is truncated.
    let dir_decompressed = {
        let mut out = Vec::new();
        // REFERENCEREGISTERED
        push_record(&mut out, 0x000D, b"X");
        out.extend_from_slice(&[0xAA, 0xBB, 0xCC, 0xDD, 0xEE]); // 5 bytes (truncated header)
        out
    };

    let vba_bin = build_vba_bin_with_dir_decompressed(&dir_decompressed);
    let err = project_normalized_data_v3(&vba_bin).expect_err("expected dir parse error");
    match err {
        ParseError::Dir(DirParseError::Truncated) => {}
        other => panic!("expected Dir(Truncated), got {other:?}"),
    }
}

#[test]
fn project_normalized_data_v3_dir_bad_record_length_beyond_buffer() {
    // Header claims `len=10`, but only 1 payload byte is present.
    let dir_decompressed = {
        let mut out = Vec::new();
        out.extend_from_slice(&0x000Du16.to_le_bytes()); // REFERENCEREGISTERED
        out.extend_from_slice(&10u32.to_le_bytes());
        out.extend_from_slice(b"X"); // insufficient payload
        out
    };

    let vba_bin = build_vba_bin_with_dir_decompressed(&dir_decompressed);
    let err = project_normalized_data_v3(&vba_bin).expect_err("expected dir parse error");
    match err {
        ParseError::Dir(DirParseError::BadRecordLength { id, len }) => {
            assert_eq!(id, 0x000D);
            assert_eq!(len, 10);
        }
        other => panic!("expected Dir(BadRecordLength), got {other:?}"),
    }
}

#[test]
fn project_normalized_data_v3_minimal_project_concatenates_selected_records() {
    let dir_decompressed = {
        let mut out = Vec::new();

        // PROJECTNAME
        push_record(&mut out, 0x0004, b"VBAProject");
        // PROJECTCONSTANTS
        push_record(&mut out, 0x000C, b"Const=1");

        // Module group
        push_record(&mut out, 0x0019, b"Module1"); // MODULENAME

        // MODULESTREAMNAME + reserved u16 at the end.
        let mut stream_name = Vec::new();
        stream_name.extend_from_slice(b"Module1");
        stream_name.extend_from_slice(&0u16.to_le_bytes());
        push_record(&mut out, 0x001A, &stream_name);

        push_record(&mut out, 0x0021, &0u16.to_le_bytes()); // MODULETYPE

        out
    };

    let vba_bin = build_vba_bin_with_dir_decompressed(&dir_decompressed);
    let normalized =
        project_normalized_data_v3_dir_records(&vba_bin).expect("ProjectNormalizedDataV3");

    let expected = [
        b"VBAProject".as_slice(),
        b"Const=1".as_slice(),
        b"Module1".as_slice(),
        b"Module1".as_slice(),
        &0u16.to_le_bytes(),
    ]
    .concat();

    assert_eq!(normalized, expected);
}

#[test]
fn project_normalized_data_v3_prefers_unicode_over_ansi_for_strings() {
    let dir_decompressed = {
        let mut out = Vec::new();

        // Both ANSI and Unicode project name records; v3 should emit only Unicode payload bytes.
        push_record(&mut out, 0x0004, b"AnsiProj");
        push_record(&mut out, 0x0040, &unicode_record_data("UniProj"));

        // Module group with both MODULENAME and MODULENAMEUNICODE.
        push_record(&mut out, 0x0019, b"AnsiMod");
        push_record(&mut out, 0x0047, &unicode_record_data("UniMod"));
        push_record(&mut out, 0x0021, &0u16.to_le_bytes()); // MODULETYPE

        out
    };

    let vba_bin = build_vba_bin_with_dir_decompressed(&dir_decompressed);
    let normalized =
        project_normalized_data_v3_dir_records(&vba_bin).expect("ProjectNormalizedDataV3");

    let expected = [utf16le_bytes("UniProj"), utf16le_bytes("UniMod"), 0u16.to_le_bytes().to_vec()]
        .concat();

    assert_eq!(normalized, expected);
    assert!(
        !normalized.windows(b"AnsiProj".len()).any(|w| w == b"AnsiProj"),
        "expected ANSI PROJECTNAME bytes to be omitted when PROJECTNAMEUNICODE is present"
    );
    assert!(
        !normalized.windows(b"AnsiMod".len()).any(|w| w == b"AnsiMod"),
        "expected ANSI MODULENAME bytes to be omitted when MODULENAMEUNICODE is present"
    );
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
fn project_normalized_data_preserves_designer_storage_element_traversal_order() {
    // Regression test for MS-OVBA `NormalizeDesignerStorage` / `NormalizeStorage` traversal order as
    // used by `ProjectNormalizedData` (MS-OVBA §2.4.2.2 + §2.4.2.6).
    //
    // The spec pseudocode iterates:
    //   FOR EACH StorageElement (stream or storage) IN Storage
    // without defining a sort order. Our implementation intentionally follows the deterministic
    // compound-file enumeration order exposed by the `cfb` crate (MS-CFB red-black tree ordering),
    // rather than sorting by full OLE path.
    //
    // This test constructs a designer storage with:
    // - stream `Y` with bytes `b"Y"`
    // - nested storage `Child` containing stream `X` with bytes `b"X"`
    //
    // Lexicographic full-path sorting would yield `Child/X` before `Y`. The storage-element order
    // used by `cfb` yields `Y` before `Child` because `Y` has a shorter name (MS-CFB compares name
    // length first).
    let cursor = Cursor::new(Vec::new());
    let mut ole = cfb::CompoundFile::create(cursor).expect("create cfb");

    // Minimal PROJECT stream identifying one designer module (ProjectDesignerModule.BaseClass).
    {
        let mut s = ole.create_stream("PROJECT").expect("PROJECT stream");
        s.write_all(b"BaseClass=UserForm1\r\n")
            .expect("write PROJECT");
    }

    // Minimal VBA/dir describing the designer module → designer storage mapping.
    //
    // FormsNormalizedData (which contributes to ProjectNormalizedData) resolves `BaseClass=` values
    // via `VBA/dir` MODULENAME → MODULESTREAMNAME.
    ole.create_storage("VBA").expect("VBA storage");
    {
        let dir_decompressed = {
            let mut out = Vec::new();
            // PROJECTCODEPAGE (u16 LE): Windows-1252.
            push_record(&mut out, 0x0003, &1252u16.to_le_bytes());
            // MODULENAME (module identifier)
            push_record(&mut out, 0x0019, b"UserForm1");
            // MODULESTREAMNAME (designer storage name) + reserved u16
            let mut stream_name = Vec::new();
            stream_name.extend_from_slice(b"UserForm1");
            stream_name.extend_from_slice(&0u16.to_le_bytes());
            push_record(&mut out, 0x001A, &stream_name);
            // MODULETYPE (UserForm)
            push_record(&mut out, 0x0021, &3u16.to_le_bytes());
            out
        };
        let dir_container = compress_container(&dir_decompressed);
        let mut s = ole.create_stream("VBA/dir").expect("dir stream");
        s.write_all(&dir_container).expect("write dir");
    }

    // Designer storage: `UserForm1` with `Y` and `Child/X`.
    ole.create_storage("UserForm1").expect("designer storage");
    ole.create_storage("UserForm1/Child")
        .expect("nested storage");
    {
        let mut s = ole.create_stream("UserForm1/Y").expect("Y stream");
        s.write_all(b"Y").expect("write Y");
    }
    {
        let mut s = ole
            .create_stream("UserForm1/Child/X")
            .expect("X stream");
        s.write_all(b"X").expect("write X");
    }

    let vba_project_bin = ole.into_inner().into_inner();
    let normalized =
        project_normalized_data(&vba_project_bin).expect("compute ProjectNormalizedData");

    // Extract the FormsNormalizedData suffix (two 1023-byte padded blocks).
    let forms_len = 1023 * 2;
    assert!(
        normalized.len() >= forms_len,
        "expected output to include FormsNormalizedData suffix"
    );
    let forms = &normalized[normalized.len() - forms_len..];

    let mut expected = Vec::new();
    expected.extend_from_slice(b"Y");
    expected.extend(std::iter::repeat(0u8).take(1022));
    expected.extend_from_slice(b"X");
    expected.extend(std::iter::repeat(0u8).take(1022));

    assert_eq!(
        forms, expected,
        "expected designer stream `Y` to be normalized before nested `Child/X`"
    );
}

#[test]
fn project_normalized_data_v3_is_sensitive_to_module_record_group_order() {
    fn build_dir_with_modules(order: [&'static str; 2]) -> Vec<u8> {
        let mut out = Vec::new();
        push_record(&mut out, 0x0004, b"VBAProject");

        for name in order {
            push_record(&mut out, 0x0019, name.as_bytes());
            let mut stream_name = Vec::new();
            stream_name.extend_from_slice(name.as_bytes());
            stream_name.extend_from_slice(&0u16.to_le_bytes());
            push_record(&mut out, 0x001A, &stream_name);
            push_record(&mut out, 0x0021, &0u16.to_le_bytes());
        }

        out
    }

    // Non-alphabetical: ModuleB then ModuleA.
    let vba_bin = build_vba_bin_with_dir_decompressed(&build_dir_with_modules(["ModuleB", "ModuleA"]));
    let normalized =
        project_normalized_data_v3_dir_records(&vba_bin).expect("ProjectNormalizedDataV3");

    let pos_b = find_subslice(&normalized, b"ModuleB").expect("ModuleB bytes present");
    let pos_a = find_subslice(&normalized, b"ModuleA").expect("ModuleA bytes present");
    assert!(pos_b < pos_a, "expected ModuleB group to precede ModuleA group");

    // Swapping module order should swap the normalized output.
    let vba_bin_swapped =
        build_vba_bin_with_dir_decompressed(&build_dir_with_modules(["ModuleA", "ModuleB"]));
    let normalized_swapped =
        project_normalized_data_v3_dir_records(&vba_bin_swapped).expect("ProjectNormalizedDataV3");
    assert_ne!(
        normalized, normalized_swapped,
        "changing module stored order should change ProjectNormalizedDataV3"
    );

    let pos_a2 = find_subslice(&normalized_swapped, b"ModuleA").expect("ModuleA bytes present");
    let pos_b2 = find_subslice(&normalized_swapped, b"ModuleB").expect("ModuleB bytes present");
    assert!(
        pos_a2 < pos_b2,
        "expected ModuleA group to precede ModuleB group when dir order is A then B"
    );
}
#[test]
fn project_normalized_data_handles_lfcr_nwln_and_strips_host_extender_ref_newlines() {
    // Regression test:
    // - MS-OVBA defines `NWLN` as CRLF *or* LFCR (not only CRLF).
    // - `String::lines()` does not treat LFCR as a single newline and can leave a leading `\r`
    //   on the next line, breaking section parsing and hashing.
    //
    // Construct a PROJECT stream using LFCR for the key lines we care about, and include both LFCR
    // and CRLF variants in HostExtenderRef lines so we assert both are stripped.
    let project_stream_bytes = concat!(
        "BaseClass=UserForm1\n\r",
        "Name=\"VBAProject\"\n\r",
        "[Host Extender Info]\n\r",
        "HostExtenderRef=RefLFCR-0123456789\n\r",
        "HostExtenderRef=RefCRLF-ABCDEFGHIJ\r\n",
        "[Workspace]\n\r",
    )
    .as_bytes();

    let designer_bytes = b"DESIGNER-STORAGE-BYTES";

    // Minimal decompressed `VBA/dir` describing a single UserForm module so FormsNormalizedData can
    // resolve the `BaseClass=` identifier to a designer storage name.
    let dir_decompressed = {
        let mut out = Vec::new();
        // PROJECTCODEPAGE (optional, but makes the encoding choice explicit).
        push_record(&mut out, 0x0003, &1252u16.to_le_bytes());

        // Module group: UserForm1
        push_record(&mut out, 0x0019, b"UserForm1"); // MODULENAME
        let mut stream_name = Vec::new();
        stream_name.extend_from_slice(b"UserForm1");
        stream_name.extend_from_slice(&0u16.to_le_bytes());
        push_record(&mut out, 0x001A, &stream_name); // MODULESTREAMNAME (+ reserved u16)
        push_record(&mut out, 0x0021, &0x0003u16.to_le_bytes()); // MODULETYPE (UserForm)
        out
    };
    let dir_container = compress_container(&dir_decompressed);

    let cursor = Cursor::new(Vec::new());
    let mut ole = cfb::CompoundFile::create(cursor).expect("create compound file");

    ole.create_storage("VBA").expect("create VBA storage");
    {
        let mut s = ole.create_stream("VBA/dir").expect("dir stream");
        s.write_all(&dir_container).expect("write dir");
    }

    {
        let mut s = ole.create_stream("PROJECT").expect("PROJECT stream");
        s.write_all(project_stream_bytes)
            .expect("write PROJECT bytes");
    }

    ole.create_storage("UserForm1")
        .expect("create designer storage");
    {
        let mut s = ole
            .create_stream("UserForm1/Payload")
            .expect("create designer stream");
        s.write_all(designer_bytes)
            .expect("write designer stream bytes");
    }

    let vba_project_bin = ole.into_inner().into_inner();
    let normalized =
        project_normalized_data(&vba_project_bin).expect("compute ProjectNormalizedData");

    // Ensure LFCR line endings do not prevent `[Host Extender Info]` section detection.
    assert!(
        find_subslice(&normalized, b"Host Extender Info").is_some(),
        "expected ProjectNormalizedData to include `Host Extender Info` section marker"
    );

    // Ensure HostExtenderRef values are present...
    assert!(
        find_subslice(&normalized, b"RefLFCR-0123456789").is_some(),
        "expected ProjectNormalizedData to include HostExtenderRef value (LFCR)"
    );
    assert!(
        find_subslice(&normalized, b"RefCRLF-ABCDEFGHIJ").is_some(),
        "expected ProjectNormalizedData to include HostExtenderRef value (CRLF)"
    );

    // ...but with all newline forms removed (both LFCR and CRLF).
    assert!(
        find_subslice(&normalized, b"RefLFCR-0123456789\n\r").is_none(),
        "expected HostExtenderRef (LFCR) to have NWLN removed"
    );
    assert!(
        find_subslice(&normalized, b"RefLFCR-0123456789\r\n").is_none(),
        "expected HostExtenderRef (LFCR) to have NWLN removed"
    );
    assert!(
        find_subslice(&normalized, b"RefCRLF-ABCDEFGHIJ\r\n").is_none(),
        "expected HostExtenderRef (CRLF) to have NWLN removed"
    );
    assert!(
        find_subslice(&normalized, b"RefCRLF-ABCDEFGHIJ\n\r").is_none(),
        "expected HostExtenderRef (CRLF) to have NWLN removed"
    );

    // Ensure designer storage bytes referenced by `BaseClass=` are included.
    assert!(
        find_subslice(&normalized, designer_bytes).is_some(),
        "expected ProjectNormalizedData to include designer storage stream bytes for BaseClass"
    );
}
