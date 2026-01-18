use std::io::{Cursor, Write};

use encoding_rs::WINDOWS_1251;
use formula_vba::{
    compress_container, contents_hash_v3, project_normalized_data, project_normalized_data_v3,
    project_normalized_data_v3_dir_records, project_normalized_data_v3_transcript, DirParseError,
    ParseError,
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
    {
        let mut s = ole.create_stream("PROJECT").expect("PROJECT stream");
        s.write_all(b"Name=\"VBAProject\"\r\n")
            .expect("write PROJECT");
    }
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
    let mut out = Vec::new();
    let _ = out.try_reserve_exact(4usize.saturating_add(units.len().saturating_mul(2)));
    out.extend_from_slice(&(units.len() as u32).to_le_bytes());
    for u in units {
        out.extend_from_slice(&u.to_le_bytes());
    }
    out
}

fn unicode_record_data_bytes_len(s: &str) -> Vec<u8> {
    let payload = utf16le_bytes(s);
    let mut out = Vec::new();
    let _ = out.try_reserve_exact(4usize.saturating_add(payload.len()));
    out.extend_from_slice(&(payload.len() as u32).to_le_bytes());
    out.extend_from_slice(&payload);
    out
}

fn utf16le_bytes_with_trailing_nul(s: &str) -> Vec<u8> {
    let mut out = utf16le_bytes(s);
    out.extend_from_slice(&0u16.to_le_bytes());
    out
}

fn unicode_record_data_bytes_len_excluding_trailing_nul(s: &str) -> Vec<u8> {
    let payload_without_nul = utf16le_bytes(s);
    let mut payload = payload_without_nul.clone();
    payload.extend_from_slice(&0u16.to_le_bytes()); // UTF-16 NUL terminator
    let mut out = Vec::new();
    let _ = out.try_reserve_exact(4usize.saturating_add(payload.len()));
    // Prefix is the byte count excluding the trailing terminator.
    out.extend_from_slice(&(payload_without_nul.len() as u32).to_le_bytes());
    out.extend_from_slice(&payload);
    out
}

fn unicode_record_data_code_units_excluding_trailing_nul(s: &str) -> Vec<u8> {
    let payload_without_nul = utf16le_bytes(s);
    let mut payload = payload_without_nul.clone();
    payload.extend_from_slice(&0u16.to_le_bytes()); // UTF-16 NUL terminator
    let mut out = Vec::new();
    let _ = out.try_reserve_exact(4usize.saturating_add(payload.len()));
    // Prefix is the UTF-16 code unit count excluding the trailing terminator.
    out.extend_from_slice(&((payload_without_nul.len() / 2) as u32).to_le_bytes());
    out.extend_from_slice(&payload);
    out
}

fn unicode_record_data_bytes_len_including_trailing_nul(s: &str) -> Vec<u8> {
    let payload = utf16le_bytes_with_trailing_nul(s);
    let mut out = Vec::new();
    let _ = out.try_reserve_exact(4usize.saturating_add(payload.len()));
    out.extend_from_slice(&(payload.len() as u32).to_le_bytes());
    out.extend_from_slice(&payload);
    out
}

fn unicode_record_data_code_units_including_trailing_nul(s: &str) -> Vec<u8> {
    let payload = utf16le_bytes_with_trailing_nul(s);
    let mut out = Vec::new();
    let _ = out.try_reserve_exact(4usize.saturating_add(payload.len()));
    out.extend_from_slice(&((payload.len() / 2) as u32).to_le_bytes());
    out.extend_from_slice(&payload);
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
        // PROJECT stream ProjectProperties contribution: key bytes + value bytes (no separators).
        b"NameVBAProject".as_slice(),
    ]
    .concat();

    assert_eq!(normalized, expected);
}

#[test]
fn project_normalized_data_prefers_alternate_unicode_dir_record_ids() {
    // Some real-world `VBA/dir` streams use non-canonical record IDs for Unicode/alternate string
    // variants:
    // - PROJECTDOCSTRINGUNICODE: 0x0041
    // - PROJECTHELPFILEPATH2:    0x0042
    // - PROJECTCONSTANTSUNICODE: 0x0043
    //
    // Ensure ProjectNormalizedData prefers these variants and skips the ANSI record payload when
    // the Unicode form immediately follows.
    let dir_decompressed = {
        let mut out = Vec::new();

        // Included: PROJECTNAME
        push_record(&mut out, 0x0004, b"MyProject");

        // PROJECTDOCSTRING (ANSI) followed by alternate Unicode record id 0x0041.
        push_record(&mut out, 0x0005, b"DocAnsi");
        push_record(&mut out, 0x0041, &utf16le_bytes("DocUni"));

        // PROJECTHELPFILEPATH (ANSI) followed by alternate second-path/Unicode record id 0x0042.
        push_record(&mut out, 0x0006, b"HelpAnsi");
        push_record(&mut out, 0x0042, &utf16le_bytes("HelpUni"));

        // PROJECTCONSTANTS (ANSI) followed by alternate Unicode record id 0x0043.
        push_record(&mut out, 0x000C, b"ConstAnsi");
        push_record(&mut out, 0x0043, &utf16le_bytes("ConstUni"));

        out
    };

    let vba_bin = build_vba_bin_with_dir_decompressed(&dir_decompressed);
    let normalized = project_normalized_data(&vba_bin).expect("ProjectNormalizedData");

    let expected = [
        b"MyProject".as_slice(),
        utf16le_bytes("DocUni").as_slice(),
        utf16le_bytes("HelpUni").as_slice(),
        utf16le_bytes("ConstUni").as_slice(),
        // PROJECT stream ProjectProperties contribution: key bytes + value bytes (no separators).
        b"NameVBAProject".as_slice(),
    ]
    .concat();

    assert_eq!(normalized, expected);
    assert!(
        find_subslice(&normalized, b"DocAnsi").is_none(),
        "expected ANSI PROJECTDOCSTRING bytes to be omitted when Unicode variant is present"
    );
    assert!(
        find_subslice(&normalized, b"HelpAnsi").is_none(),
        "expected ANSI PROJECTHELPFILEPATH bytes to be omitted when alternate Unicode variant is present"
    );
    assert!(
        find_subslice(&normalized, b"ConstAnsi").is_none(),
        "expected ANSI PROJECTCONSTANTS bytes to be omitted when Unicode variant is present"
    );
}

#[test]
fn project_normalized_data_strips_internal_unicode_length_prefix_for_project_records() {
    // Some producers embed an *internal* u32 length prefix inside the payload of Unicode records
    // (in addition to the normal `Id || Size || Data` framing). Ensure we strip this prefix when it
    // is consistent with the remaining bytes.
    let dir_decompressed = {
        let mut out = Vec::new();

        // Included: PROJECTNAME
        push_record(&mut out, 0x0004, b"MyProject");

        // PROJECTDOCSTRING (ANSI) followed by alternate Unicode record id 0x0041, but the Unicode
        // payload itself has an internal u32 length prefix.
        push_record(&mut out, 0x0005, b"DocAnsi");
        push_record(&mut out, 0x0041, &unicode_record_data("DocUni"));

        // PROJECTHELPFILEPATH (ANSI) followed by alternate second-path/Unicode record id 0x0042,
        // with an internal *byte-count* length prefix.
        push_record(&mut out, 0x0006, b"HelpAnsi");
        push_record(&mut out, 0x0042, &unicode_record_data_bytes_len("HelpUni"));

        // PROJECTCONSTANTS (ANSI) followed by alternate Unicode record id 0x0043 with internal
        // length prefix.
        push_record(&mut out, 0x000C, b"ConstAnsi");
        push_record(&mut out, 0x0043, &unicode_record_data("ConstUni"));

        out
    };

    let vba_bin = build_vba_bin_with_dir_decompressed(&dir_decompressed);
    let normalized = project_normalized_data(&vba_bin).expect("ProjectNormalizedData");

    let expected = [
        b"MyProject".as_slice(),
        utf16le_bytes("DocUni").as_slice(),
        utf16le_bytes("HelpUni").as_slice(),
        utf16le_bytes("ConstUni").as_slice(),
        b"NameVBAProject".as_slice(),
    ]
    .concat();
    assert_eq!(normalized, expected);

    // Ensure the internal u32 length prefixes were removed.
    assert!(
        find_subslice(&normalized, &("DocUni".encode_utf16().count() as u32).to_le_bytes()).is_none(),
        "did not expect internal code-unit length prefix bytes to appear in output"
    );
    assert!(
        find_subslice(&normalized, &(utf16le_bytes("HelpUni").len() as u32).to_le_bytes()).is_none(),
        "did not expect internal byte-length prefix bytes to appear in output"
    );
}

#[test]
fn project_normalized_data_includes_projectcompatversion_record() {
    // Real-world `VBA/dir` streams often include PROJECTCOMPATVERSION (0x004A) in the
    // ProjectInformation record list. Ensure it contributes to ProjectNormalizedData.
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

    assert_ne!(
        normalized_without, normalized_with,
        "PROJECTCOMPATVERSION (0x004A) must affect ProjectNormalizedData when present"
    );
    assert!(
        find_subslice(&normalized_without, b"VBAProject").is_some(),
        "expected ProjectNormalizedData to include PROJECTNAME bytes"
    );
    assert!(
        !normalized_without
            .windows(compat_version.len())
            .any(|w| w == compat_version),
        "ProjectNormalizedData without compat must not contain PROJECTCOMPATVERSION payload bytes"
    );
    assert!(
        normalized_with
            .windows(compat_version.len())
            .any(|w| w == compat_version),
        "ProjectNormalizedData must include PROJECTCOMPATVERSION payload bytes when present"
    );
}

#[test]
fn project_normalized_data_ignores_workspace_section_from_project_stream() {
    // MS-OVBA `PROJECT` stream structure:
    //   VBAPROJECTText = ProjectProperties NWLN HostExtenders [NWLN ProjectWorkspace]
    //
    // Regression: ProjectNormalizedData MUST ignore ProjectWorkspace / [Workspace] section.
    //
    // This is important for V3 signature binding because the Workspace section is machine-local.
    let project_stream = concat!(
        "Name=\"VBAProject\"\r\n",
        "\r\n",
        "[Host Extender Info]\r\n",
        "HostExtenderRef=MyHostExtender\r\n",
        "\r\n",
        "[Workspace]\r\n",
        "ThisWorkbook=SHOULD_NOT_APPEAR_IN_HASH\r\n",
    );

    // `project_normalized_data()` always incorporates data from `VBA/dir`, so include a minimal
    // stream with a *different* project name to ensure we're asserting the Name property is taken
    // from the PROJECT stream.
    let dir_decompressed = {
        let mut out = Vec::new();
        // PROJECTSYSKIND
        push_record(&mut out, 0x0001, &1u32.to_le_bytes());
        // PROJECTNAME (distinct from the PROJECT stream's `Name="VBAProject"` line)
        push_record(&mut out, 0x0004, b"DirProject");
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
        let mut s = ole.create_stream("PROJECT").expect("PROJECT stream");
        s.write_all(project_stream.as_bytes())
            .expect("write PROJECT");
    }

    let vba_bin = ole.into_inner().into_inner();
    let normalized = project_normalized_data(&vba_bin).expect("ProjectNormalizedData");

    // ProjectProperties contribution: property name + value bytes (no separators).
    assert!(
        find_subslice(&normalized, b"Name").is_some(),
        "expected ProjectNormalizedData to include PROJECT stream property name bytes"
    );
    assert!(
        find_subslice(&normalized, b"VBAProject").is_some(),
        "expected ProjectNormalizedData to include PROJECT stream property value bytes"
    );

    // HostExtenders contribution: include section name and HostExtenderRef line bytes (no NWLN).
    assert!(
        find_subslice(&normalized, b"Host Extender Info").is_some(),
        "expected ProjectNormalizedData to include Host Extender Info section contribution"
    );
    assert!(
        find_subslice(&normalized, b"HostExtenderRef=MyHostExtender").is_some(),
        "expected ProjectNormalizedData to include HostExtenderRef line bytes"
    );
    assert!(
        find_subslice(&normalized, b"HostExtenderRef=MyHostExtender\r\n").is_none()
            && find_subslice(&normalized, b"HostExtenderRef=MyHostExtender\n\r").is_none(),
        "expected HostExtenderRef line bytes to be appended without NWLN"
    );

    // ProjectWorkspace must be ignored (neither the header nor its distinctive lines should appear).
    assert!(
        find_subslice(&normalized, b"Workspace").is_none(),
        "ProjectNormalizedData must ignore the [Workspace] section"
    );
    assert!(
        find_subslice(&normalized, b"ThisWorkbook=SHOULD_NOT_APPEAR_IN_HASH").is_none(),
        "ProjectNormalizedData must ignore ProjectWorkspace lines"
    );
}

#[test]
fn project_normalized_data_missing_vba_dir_stream() {
    let cursor = Cursor::new(Vec::new());
    let mut ole = cfb::CompoundFile::create(cursor).expect("create cfb");
    ole.create_storage("VBA").expect("VBA storage");
    {
        let mut s = ole.create_stream("PROJECT").expect("PROJECT stream");
        s.write_all(b"Name=\"VBAProject\"\r\n")
            .expect("write PROJECT");
    }

    let vba_bin = ole.into_inner().into_inner();
    let err = project_normalized_data(&vba_bin).expect_err("expected MissingStream");
    match err {
        ParseError::MissingStream("VBA/dir") => {}
        other => panic!("expected MissingStream(\"VBA/dir\"), got {other:?}"),
    }
}

#[test]
fn project_normalized_data_dir_truncated_record_header() {
    // One valid record followed by <6 leftover bytes so the next record header is truncated.
    let dir_decompressed = {
        let mut out = Vec::new();
        // PROJECTSYSKIND
        push_record(&mut out, 0x0001, &1u32.to_le_bytes());
        out.extend_from_slice(&[0xAA, 0xBB, 0xCC, 0xDD, 0xEE]); // 5 bytes (truncated header)
        out
    };

    let vba_bin = build_vba_bin_with_dir_decompressed(&dir_decompressed);
    let err = project_normalized_data(&vba_bin).expect_err("expected dir parse error");
    match err {
        ParseError::Dir(DirParseError::Truncated) => {}
        other => panic!("expected Dir(Truncated), got {other:?}"),
    }
}

#[test]
fn project_normalized_data_dir_bad_record_length_beyond_buffer() {
    // Header claims `len=10`, but only 1 payload byte is present.
    let dir_decompressed = {
        let mut out = Vec::new();
        out.extend_from_slice(&0x0001u16.to_le_bytes()); // PROJECTSYSKIND
        out.extend_from_slice(&10u32.to_le_bytes());
        out.extend_from_slice(b"X"); // insufficient payload
        out
    };

    let vba_bin = build_vba_bin_with_dir_decompressed(&dir_decompressed);
    let err = project_normalized_data(&vba_bin).expect_err("expected dir parse error");
    match err {
        ParseError::Dir(DirParseError::BadRecordLength { id, len }) => {
            assert_eq!(id, 0x0001);
            assert_eq!(len, 10);
        }
        other => panic!("expected Dir(BadRecordLength), got {other:?}"),
    }
}

#[test]
fn project_normalized_data_dir_truncated_record_header_after_module_records() {
    // Regression: `project_normalized_data()` must still validate record framing *after* the first
    // module record group begins. Previously it stopped parsing at MODULENAME (0x0019), which could
    // hide truncation errors later in the stream.
    let dir_decompressed = {
        let mut out = Vec::new();
        push_record(&mut out, 0x0003, &1252u16.to_le_bytes()); // PROJECTCODEPAGE
        push_record(&mut out, 0x0019, b"Module1"); // MODULENAME (module record group start)
        out.extend_from_slice(&[0xAA, 0xBB, 0xCC, 0xDD, 0xEE]); // 5 bytes (truncated header)
        out
    };

    let vba_bin = build_vba_bin_with_dir_decompressed(&dir_decompressed);
    let err = project_normalized_data(&vba_bin).expect_err("expected dir parse error");
    match err {
        ParseError::Dir(DirParseError::Truncated) => {}
        other => panic!("expected Dir(Truncated), got {other:?}"),
    }
}

#[test]
fn project_normalized_data_dir_bad_record_length_beyond_buffer_after_module_records() {
    // Regression: ensure length validation is applied to records after module records too.
    let dir_decompressed = {
        let mut out = Vec::new();
        push_record(&mut out, 0x0003, &1252u16.to_le_bytes()); // PROJECTCODEPAGE
        push_record(&mut out, 0x0019, b"Module1"); // MODULENAME

        // Next record header claims `len=10`, but only 1 payload byte is present.
        out.extend_from_slice(&0x9999u16.to_le_bytes());
        out.extend_from_slice(&10u32.to_le_bytes());
        out.extend_from_slice(b"X"); // insufficient payload
        out
    };

    let vba_bin = build_vba_bin_with_dir_decompressed(&dir_decompressed);
    let err = project_normalized_data(&vba_bin).expect_err("expected dir parse error");
    match err {
        ParseError::Dir(DirParseError::BadRecordLength { id, len }) => {
            assert_eq!(id, 0x9999);
            assert_eq!(len, 10);
        }
        other => panic!("expected Dir(BadRecordLength), got {other:?}"),
    }
}

#[test]
fn project_normalized_data_stops_project_information_parsing_on_modulenameunicode() {
    // Some projects/fixtures may omit MODULENAME (0x0019) and begin the module list with
    // MODULENAMEUNICODE (0x0047). We must treat that as "modules have started" and stop
    // incorporating ProjectInformation record IDs that can be ambiguous by context (notably 0x004A).
    let dir_decompressed = {
        let mut out = Vec::new();

        // Start module records with MODULENAMEUNICODE (no prior MODULENAME).
        push_record(&mut out, 0x0047, &utf16le_bytes("Module1"));

        // This record ID is meaningful as PROJECTCOMPATVERSION only in ProjectInformation (before
        // modules). Ensure it is ignored here.
        push_record(&mut out, 0x004A, b"__COMPAT_SHOULD_NOT_APPEAR__");

        out
    };

    let vba_bin = build_vba_bin_with_dir_decompressed(&dir_decompressed);
    let normalized = project_normalized_data(&vba_bin).expect("ProjectNormalizedData");

    assert_eq!(
        normalized,
        b"NameVBAProject",
        "expected only PROJECT stream ProjectProperties tokens when dir records are module-only"
    );
    assert!(
        find_subslice(&normalized, b"__COMPAT_SHOULD_NOT_APPEAR__").is_none(),
        "expected 0x004A payload bytes to be ignored once module records start"
    );
}

#[test]
fn project_normalized_data_ignores_fixed_length_projectversion_record_after_module_records() {
    // Regression: PROJECTVERSION has special-case fixed-length parsing. Ensure this special-case
    // does not accidentally include bytes when the record appears after the module section begins.
    let dir_decompressed = {
        let mut out = Vec::new();

        // Begin module records.
        push_record(&mut out, 0x0019, b"Module1"); // MODULENAME

        // Append a fixed-length PROJECTVERSION record (0x0009). This should be skipped since the
        // project info section is already over.
        out.extend_from_slice(&0x0009u16.to_le_bytes());
        out.extend_from_slice(&0u32.to_le_bytes()); // Reserved
        out.extend_from_slice(&1u32.to_le_bytes()); // VersionMajor
        out.extend_from_slice(&0u16.to_le_bytes()); // VersionMinor

        out
    };

    let vba_bin = build_vba_bin_with_dir_decompressed(&dir_decompressed);
    let normalized = project_normalized_data(&vba_bin).expect("ProjectNormalizedData");

    assert_eq!(
        normalized,
        b"NameVBAProject",
        "expected PROJECTVERSION bytes to be ignored after module records begin"
    );
}

#[test]
fn project_normalized_data_v3_missing_vba_dir_stream() {
    let cursor = Cursor::new(Vec::new());
    let mut ole = cfb::CompoundFile::create(cursor).expect("create cfb");
    {
        let mut s = ole.create_stream("PROJECT").expect("PROJECT stream");
        s.write_all(b"Name=\"VBAProject\"\r\n")
            .expect("write PROJECT");
    }
    ole.create_storage("VBA").expect("VBA storage");

    let vba_bin = ole.into_inner().into_inner();
    let err = project_normalized_data_v3(&vba_bin).expect_err("expected MissingStream");
    match err {
        ParseError::MissingStream("VBA/dir") => {}
        other => panic!("expected MissingStream(\"VBA/dir\"), got {other:?}"),
    }
}

#[test]
fn project_normalized_data_v3_filters_project_stream_properties_and_includes_designer_bytes() {
    // Build a minimal vbaProject.bin with a `PROJECT` stream containing both excluded and included
    // properties per MS-OVBA §2.4.2.6. Also include a `BaseClass=` line so FormsNormalizedData is
    // incorporated.
    let project_stream = concat!(
        "ID=\"{00000000-0000-0000-0000-000000000000}\"\r\n",
        "Document=ThisWorkbook/&H00000000\r\n",
        "CMG=\"CMGSECRET\"\r\n",
        "DPB=\"DPBSECRET\"\r\n",
        "GC=\"GCSECRET\"\r\n",
        "Name=\"VBAProject\"\r\n",
        "Package={11111111-2222-3333-4444-555555555555}\r\n",
        "BaseClass=UserForm1\r\n",
    )
    .as_bytes();

    let userform_source = b"Sub FormHello()\r\nEnd Sub\r\n";
    let userform_container = compress_container(userform_source);

    let dir_decompressed = {
        let mut out = Vec::new();
        // MODULENAME
        push_record(&mut out, 0x0019, b"UserForm1");
        // MODULESTREAMNAME + reserved u16
        let mut stream_name = Vec::new();
        stream_name.extend_from_slice(b"UserForm1");
        stream_name.extend_from_slice(&0u16.to_le_bytes());
        push_record(&mut out, 0x001A, &stream_name);
        // MODULETYPE = UserForm (0x0003 per MS-OVBA)
        push_record(&mut out, 0x0021, &0x0003u16.to_le_bytes());
        // MODULETEXTOFFSET (0)
        push_record(&mut out, 0x0031, &0u32.to_le_bytes());
        out
    };
    let dir_container = compress_container(&dir_decompressed);

    let designer_bytes = b"DESIGNER-STORAGE-BYTES";

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
            .expect("write userform module bytes");
    }
    {
        let mut s = ole
            .create_stream("UserForm1/Payload")
            .expect("designer stream");
        s.write_all(designer_bytes).expect("write designer bytes");
    }

    let vba_bin = ole.into_inner().into_inner();
    let normalized =
        project_normalized_data_v3_transcript(&vba_bin).expect("ProjectNormalizedData v3");

    // Excluded PROJECT properties must not contribute.
    for needle in [
        b"ID=" as &[u8],
        b"Document=" as &[u8],
        b"CMG=" as &[u8],
        b"DPB=" as &[u8],
        b"GC=" as &[u8],
        b"CMGSECRET" as &[u8],
        b"DPBSECRET" as &[u8],
        b"GCSECRET" as &[u8],
    ] {
        assert!(
            find_subslice(&normalized, needle).is_none(),
            "did not expect ProjectNormalizedData to contain excluded property bytes: {:?}",
            std::str::from_utf8(needle).unwrap_or("<non-utf8>"),
        );
    }

    // At least one included property should be present.
    assert!(
        find_subslice(&normalized, b"Name=\"VBAProject\"").is_some(),
        "expected included Name property to be present"
    );
    assert!(
        find_subslice(&normalized, b"Package={11111111-2222-3333-4444-555555555555}").is_some(),
        "expected included Package property to be present"
    );

    // Designer storage bytes must be included when BaseClass= is present.
    assert!(
        find_subslice(&normalized, designer_bytes).is_some(),
        "expected ProjectNormalizedData to include designer stream bytes when BaseClass= is present"
    );

    // Regression guard: changing excluded PROJECT properties must not affect ContentsHashV3.
    let project_stream_changed_excluded = std::str::from_utf8(project_stream)
        .expect("PROJECT stream is valid UTF-8 for this test")
        .replace("CMGSECRET", "CMGCHANGED")
        .into_bytes();
    let cursor = Cursor::new(Vec::new());
    let mut ole2 = cfb::CompoundFile::create(cursor).expect("create cfb");
    ole2.create_storage("VBA").expect("VBA storage");
    ole2.create_storage("UserForm1").expect("designer storage");
    {
        let mut s = ole2.create_stream("PROJECT").expect("PROJECT stream");
        s.write_all(&project_stream_changed_excluded)
            .expect("write PROJECT");
    }
    {
        let mut s = ole2.create_stream("VBA/dir").expect("dir stream");
        s.write_all(&dir_container).expect("write dir");
    }
    {
        let mut s = ole2
            .create_stream("VBA/UserForm1")
            .expect("userform module stream");
        s.write_all(&userform_container)
            .expect("write userform module bytes");
    }
    {
        let mut s = ole2
            .create_stream("UserForm1/Payload")
            .expect("designer stream");
        s.write_all(designer_bytes).expect("write designer bytes");
    }
    let vba_bin2 = ole2.into_inner().into_inner();

    let digest1 = contents_hash_v3(&vba_bin).expect("ContentsHashV3");
    let digest2 = contents_hash_v3(&vba_bin2).expect("ContentsHashV3 (excluded props changed)");
    assert_eq!(
        digest2, digest1,
        "excluded PROJECT properties must not influence ContentsHashV3"
    );

    // Changing an included property should affect the hash.
    let project_stream_changed_included = std::str::from_utf8(project_stream)
        .expect("PROJECT stream is valid UTF-8 for this test")
        .replace("VBAProject", "VBAProject2")
        .into_bytes();
    let cursor = Cursor::new(Vec::new());
    let mut ole3 = cfb::CompoundFile::create(cursor).expect("create cfb");
    ole3.create_storage("VBA").expect("VBA storage");
    ole3.create_storage("UserForm1").expect("designer storage");
    {
        let mut s = ole3.create_stream("PROJECT").expect("PROJECT stream");
        s.write_all(&project_stream_changed_included)
            .expect("write PROJECT");
    }
    {
        let mut s = ole3.create_stream("VBA/dir").expect("dir stream");
        s.write_all(&dir_container).expect("write dir");
    }
    {
        let mut s = ole3
            .create_stream("VBA/UserForm1")
            .expect("userform module stream");
        s.write_all(&userform_container)
            .expect("write userform module bytes");
    }
    {
        let mut s = ole3
            .create_stream("UserForm1/Payload")
            .expect("designer stream");
        s.write_all(designer_bytes).expect("write designer bytes");
    }
    let vba_bin3 = ole3.into_inner().into_inner();
    let digest3 = contents_hash_v3(&vba_bin3).expect("ContentsHashV3 (included props changed)");
    assert_ne!(
        digest3, digest1,
        "included PROJECT properties must influence ContentsHashV3"
    );
}

#[test]
fn project_normalized_data_v3_ignores_workspace_section_and_excludes_additional_security_properties() {
    // Regression: the `[Workspace]` section is machine/user-local and MUST NOT affect v3 signature
    // binding. Also ensure we exclude additional protection-related properties commonly seen in the
    // wild.
    let project_stream = concat!(
        "Name=\"VBAProject\"\r\n",
        "Password=SHOULD_SKIP\r\n",
        "VisibilityState=SHOULD_SKIP\r\n",
        "DocModule=SHOULD_SKIP\r\n",
        "ProtectionState=SHOULD_SKIP\r\n",
        "[Host Extender Info]\r\n",
        "HostExtenderRef=MyHostExtender\r\n",
        "[Workspace]\r\n",
        "ThisWorkbook=SHOULD_NOT_APPEAR_IN_HASH\r\n",
    )
    .as_bytes();

    // Minimal `VBA/dir` stream: only the dir Terminator record.
    let dir_decompressed = {
        let mut out = Vec::new();
        push_record(&mut out, 0x0010, &[]);
        out
    };
    let dir_container = compress_container(&dir_decompressed);

    let cursor = Cursor::new(Vec::new());
    let mut ole = cfb::CompoundFile::create(cursor).expect("create cfb");
    ole.create_storage("VBA").expect("VBA storage");
    {
        let mut s = ole.create_stream("PROJECT").expect("PROJECT stream");
        s.write_all(project_stream).expect("write PROJECT");
    }
    {
        let mut s = ole.create_stream("VBA/dir").expect("dir stream");
        s.write_all(&dir_container).expect("write dir");
    }

    let vba_bin = ole.into_inner().into_inner();
    let normalized =
        project_normalized_data_v3_transcript(&vba_bin).expect("ProjectNormalizedData v3");

    // Included properties should appear.
    assert!(
        find_subslice(&normalized, b"Name=\"VBAProject\"").is_some(),
        "expected included Name property bytes to be present"
    );
    assert!(
        find_subslice(&normalized, b"HostExtenderRef=MyHostExtender").is_some(),
        "expected Host Extender Info line bytes to be present"
    );

    // Excluded properties should not appear anywhere in the transcript.
    for needle in [
        b"Password=" as &[u8],
        b"VisibilityState=" as &[u8],
        b"DocModule=" as &[u8],
        b"ProtectionState=" as &[u8],
    ] {
        assert!(
            find_subslice(&normalized, needle).is_none(),
            "did not expect ProjectNormalizedData v3 to contain excluded property bytes: {:?}",
            std::str::from_utf8(needle).unwrap_or("<non-utf8>"),
        );
    }

    // Workspace section should be ignored entirely (including its distinctive key/value lines).
    assert!(
        find_subslice(&normalized, b"ThisWorkbook=SHOULD_NOT_APPEAR_IN_HASH").is_none(),
        "expected [Workspace] section lines to be ignored for v3 ProjectNormalizedData"
    );
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
fn project_normalized_data_v3_dir_records_skips_projectcompatversion_record() {
    // `project_normalized_data_v3_dir_records` is used as a v3 building block and should be robust
    // to PROJECTCOMPATVERSION (0x004A) appearing in the project information record list.
    let compat_version = 0xDEADBEEFu32.to_le_bytes();

    fn build_dir(include_compat: bool, compat_version: &[u8; 4]) -> Vec<u8> {
        let mut out = Vec::new();

        push_record(&mut out, 0x0001, &1u32.to_le_bytes()); // PROJECTSYSKIND
        if include_compat {
            push_record(&mut out, 0x004A, compat_version); // PROJECTCOMPATVERSION
        }
        push_record(&mut out, 0x0002, &0x0409u32.to_le_bytes()); // PROJECTLCID
        push_record(&mut out, 0x0003, &1252u16.to_le_bytes()); // PROJECTCODEPAGE

        push_record(&mut out, 0x0004, b"VBAProject"); // PROJECTNAME
        push_record(&mut out, 0x000C, b"Constants"); // PROJECTCONSTANTS

        // Minimal module group.
        push_record(&mut out, 0x0019, b"Module1"); // MODULENAME
        let mut stream_name = Vec::new();
        stream_name.extend_from_slice(b"Module1");
        stream_name.extend_from_slice(&0u16.to_le_bytes());
        push_record(&mut out, 0x001A, &stream_name); // MODULESTREAMNAME (+ reserved u16)
        push_record(&mut out, 0x0021, &0u16.to_le_bytes()); // MODULETYPE

        out
    }

    let dir_without = build_dir(false, &compat_version);
    let dir_with = build_dir(true, &compat_version);

    let vba_without = build_vba_bin_with_dir_decompressed(&dir_without);
    let vba_with = build_vba_bin_with_dir_decompressed(&dir_with);

    let normalized_without = project_normalized_data_v3_dir_records(&vba_without)
        .expect("ProjectNormalizedDataV3 dir-records without compat");
    let normalized_with = project_normalized_data_v3_dir_records(&vba_with)
        .expect("ProjectNormalizedDataV3 dir-records with compat");

    assert_eq!(
        normalized_without, normalized_with,
        "PROJECTCOMPATVERSION (0x004A) must not affect ProjectNormalizedDataV3 dir-record transcript"
    );
    assert!(
        find_subslice(&normalized_without, b"VBAProject").is_some(),
        "expected ProjectNormalizedDataV3 to include PROJECTNAME bytes"
    );
    assert!(
        !normalized_without
            .windows(compat_version.len())
            .any(|w| w == compat_version),
        "ProjectNormalizedDataV3 must skip PROJECTCOMPATVERSION payload bytes"
    );
    assert!(
        !normalized_with
            .windows(compat_version.len())
            .any(|w| w == compat_version),
        "ProjectNormalizedDataV3 must skip PROJECTCOMPATVERSION payload bytes"
    );
}

#[test]
fn project_normalized_data_v3_prefers_unicode_over_ansi_for_strings() {
    let dir_decompressed = {
        let mut out = Vec::new();

        // Both ANSI and Unicode project docstring records; v3 should emit only Unicode payload bytes.
        push_record(&mut out, 0x0005, b"AnsiDoc");
        push_record(&mut out, 0x0040, &unicode_record_data("UniDoc"));

        // Module group with both MODULENAME and MODULENAMEUNICODE.
        push_record(&mut out, 0x0019, b"AnsiMod");
        push_record(&mut out, 0x0047, &unicode_record_data("UniMod"));

        // Both ANSI and Unicode MODULESTREAMNAME records.
        let mut stream_name = Vec::new();
        stream_name.extend_from_slice(b"AnsiStream");
        stream_name.extend_from_slice(&0u16.to_le_bytes());
        push_record(&mut out, 0x001A, &stream_name);
        push_record(&mut out, 0x0032, &unicode_record_data("UniStream"));

        push_record(&mut out, 0x0021, &0u16.to_le_bytes()); // MODULETYPE

        out
    };

    let vba_bin = build_vba_bin_with_dir_decompressed(&dir_decompressed);
    let normalized =
        project_normalized_data_v3_dir_records(&vba_bin).expect("ProjectNormalizedDataV3");

    let expected = [
        utf16le_bytes("UniDoc"),
        utf16le_bytes("UniMod"),
        utf16le_bytes("UniStream"),
        0u16.to_le_bytes().to_vec(),
    ]
    .concat();

    assert_eq!(normalized, expected);
    assert!(
        !normalized.windows(b"AnsiDoc".len()).any(|w| w == b"AnsiDoc"),
        "expected ANSI PROJECTDOCSTRING bytes to be omitted when PROJECTDOCSTRINGUNICODE is present"
    );
    assert!(
        !normalized.windows(b"AnsiMod".len()).any(|w| w == b"AnsiMod"),
        "expected ANSI MODULENAME bytes to be omitted when MODULENAMEUNICODE is present"
    );
    assert!(
        !normalized
            .windows(b"AnsiStream".len())
            .any(|w| w == b"AnsiStream"),
        "expected ANSI MODULESTREAMNAME bytes to be omitted when MODULESTREAMNAMEUNICODE is present"
    );
}

#[test]
fn project_normalized_data_v3_treats_0048_as_modulestreamnameunicode_when_following_modulestreamname() {
    // Some real-world `VBA/dir` encodings store MODULESTREAMNAMEUNICODE as a separate record with id
    // 0x0048 immediately following MODULESTREAMNAME (0x001A). (0x0048 is normally used for
    // MODULEDOCSTRINGUNICODE in TLV-ish layouts.)
    //
    // Ensure `project_normalized_data_v3_dir_records` treats it as a Unicode stream-name variant so
    // the ANSI MODULESTREAMNAME bytes are omitted.
    let dir_decompressed = {
        let mut out = Vec::new();

        push_record(&mut out, 0x0019, b"Module1");

        // ANSI MODULESTREAMNAME should be omitted when the Unicode variant is present.
        let mut stream_name = Vec::new();
        stream_name.extend_from_slice(b"AnsiStream");
        stream_name.extend_from_slice(&0u16.to_le_bytes());
        push_record(&mut out, 0x001A, &stream_name);

        // Unicode stream-name bytes stored in a separate record id (0x0048).
        push_record(&mut out, 0x0048, &unicode_record_data("UniStream"));

        push_record(&mut out, 0x0021, &0u16.to_le_bytes()); // MODULETYPE

        out
    };

    let vba_bin = build_vba_bin_with_dir_decompressed(&dir_decompressed);
    let normalized =
        project_normalized_data_v3_dir_records(&vba_bin).expect("ProjectNormalizedDataV3");

    let expected = [
        b"Module1".as_slice(),
        utf16le_bytes("UniStream").as_slice(),
        0u16.to_le_bytes().as_slice(),
    ]
    .concat();

    assert_eq!(normalized, expected);
    assert!(
        !normalized
            .windows(b"AnsiStream".len())
            .any(|w| w == b"AnsiStream"),
        "expected ANSI MODULESTREAMNAME bytes to be omitted when Unicode (0x0048) variant is present"
    );
}

#[test]
fn project_normalized_data_v3_accepts_moduledocstring_record_id_001b() {
    // Some producers use 0x001B instead of 0x001C for MODULEDOCSTRING (ANSI).
    let dir_decompressed = {
        let mut out = Vec::new();

        push_record(&mut out, 0x0019, b"Module1");
        push_record(&mut out, 0x001B, b"AnsiDoc");
        push_record(&mut out, 0x0021, &0u16.to_le_bytes()); // MODULETYPE

        out
    };

    let vba_bin = build_vba_bin_with_dir_decompressed(&dir_decompressed);
    let normalized =
        project_normalized_data_v3_dir_records(&vba_bin).expect("ProjectNormalizedDataV3");

    let expected = [b"Module1".to_vec(), b"AnsiDoc".to_vec(), 0u16.to_le_bytes().to_vec()].concat();

    assert_eq!(normalized, expected);
}

#[test]
fn project_normalized_data_v3_omits_moduledocstring_001b_when_unicode_record_present() {
    // When the Unicode docstring variant is present, v3 should omit the ANSI docstring bytes (even
    // when the ANSI record id is 0x001B).
    let dir_decompressed = {
        let mut out = Vec::new();

        push_record(&mut out, 0x0019, b"Module1");
        push_record(&mut out, 0x001B, b"AnsiDoc");
        push_record(&mut out, 0x0048, &unicode_record_data("UniDoc"));
        push_record(&mut out, 0x0021, &0u16.to_le_bytes()); // MODULETYPE

        out
    };

    let vba_bin = build_vba_bin_with_dir_decompressed(&dir_decompressed);
    let normalized =
        project_normalized_data_v3_dir_records(&vba_bin).expect("ProjectNormalizedDataV3");

    let expected = [b"Module1".to_vec(), utf16le_bytes("UniDoc"), 0u16.to_le_bytes().to_vec()].concat();

    assert_eq!(normalized, expected);
    assert!(
        !normalized.windows(b"AnsiDoc".len()).any(|w| w == b"AnsiDoc"),
        "expected ANSI MODULEDOCSTRING bytes to be omitted when Unicode variant is present"
    );
}

#[test]
fn project_normalized_data_v3_treats_0049_as_moduledocstringunicode_when_following_moduledocstring() {
    // Some producers appear to use 0x0049 as the Unicode marker/record for MODULEDOCSTRING rather
    // than the canonical 0x0048. Ensure we only treat it as a docstring Unicode variant when it
    // follows MODULEDOCSTRING, and that it causes the ANSI docstring bytes to be omitted.
    let dir_decompressed = {
        let mut out = Vec::new();

        push_record(&mut out, 0x0019, b"Module1");
        push_record(&mut out, 0x001C, b"AnsiDoc");
        push_record(&mut out, 0x0049, &unicode_record_data("UniDoc"));
        push_record(&mut out, 0x0021, &0u16.to_le_bytes()); // MODULETYPE

        out
    };

    let vba_bin = build_vba_bin_with_dir_decompressed(&dir_decompressed);
    let normalized =
        project_normalized_data_v3_dir_records(&vba_bin).expect("ProjectNormalizedDataV3");

    let expected = [b"Module1".to_vec(), utf16le_bytes("UniDoc"), 0u16.to_le_bytes().to_vec()].concat();

    assert_eq!(normalized, expected);
    assert!(
        !normalized.windows(b"AnsiDoc".len()).any(|w| w == b"AnsiDoc"),
        "expected ANSI MODULEDOCSTRING bytes to be omitted when 0x0049 Unicode variant is present"
    );
}

#[test]
fn project_normalized_data_v3_ignores_modulehelpfilepathunicode_0049() {
    // Ensure we do not accidentally treat MODULEHELPFILEPATHUNICODE (0x0049) as a docstring record.
    // `project_normalized_data_v3_dir_records` intentionally excludes helpfile path records.
    let dir_decompressed = {
        let mut out = Vec::new();

        push_record(&mut out, 0x0019, b"Module1");
        push_record(&mut out, 0x001D, b"HelpPath"); // MODULEHELPFILEPATH
        push_record(&mut out, 0x0049, &unicode_record_data("HelpPathUni"));
        push_record(&mut out, 0x0021, &0u16.to_le_bytes()); // MODULETYPE

        out
    };

    let vba_bin = build_vba_bin_with_dir_decompressed(&dir_decompressed);
    let normalized =
        project_normalized_data_v3_dir_records(&vba_bin).expect("ProjectNormalizedDataV3");

    let expected = [b"Module1".to_vec(), 0u16.to_le_bytes().to_vec()].concat();
    assert_eq!(normalized, expected);
}

#[test]
fn project_normalized_data_v3_includes_nonprocedural_moduletype_record_id_0022() {
    // Some module groups use TypeRecord.Id=0x0022 (non-procedural modules). This helper includes
    // module type record data bytes, so it should include the 0x0022 payload too.
    let dir_decompressed = {
        let mut out = Vec::new();

        push_record(&mut out, 0x0019, b"Module1");
        push_record(&mut out, 0x0022, &0u16.to_le_bytes()); // MODULETYPE (non-procedural)

        out
    };

    let vba_bin = build_vba_bin_with_dir_decompressed(&dir_decompressed);
    let normalized =
        project_normalized_data_v3_dir_records(&vba_bin).expect("ProjectNormalizedDataV3");

    let expected = [b"Module1".to_vec(), 0u16.to_le_bytes().to_vec()].concat();
    assert_eq!(normalized, expected);
}

#[test]
fn project_normalized_data_v3_strips_unicode_length_prefix_when_prefix_is_byte_count() {
    // Some producers embed an internal u32 length prefix in Unicode dir record payloads where the
    // length is the UTF-16LE byte count (not code units). Ensure we strip the prefix in this case.
    let dir_decompressed = {
        let mut out = Vec::new();

        // Both ANSI and Unicode project docstring records; v3 should emit only Unicode payload bytes.
        push_record(&mut out, 0x0005, b"AnsiDoc");
        push_record(&mut out, 0x0040, &unicode_record_data_bytes_len("UniDoc"));

        // Module group with both MODULENAME and MODULENAMEUNICODE.
        push_record(&mut out, 0x0019, b"AnsiMod");
        push_record(&mut out, 0x0047, &unicode_record_data_bytes_len("UniMod"));

        // Both ANSI and Unicode MODULESTREAMNAME records.
        let mut stream_name = Vec::new();
        stream_name.extend_from_slice(b"AnsiStream");
        stream_name.extend_from_slice(&0u16.to_le_bytes());
        push_record(&mut out, 0x001A, &stream_name);
        push_record(&mut out, 0x0032, &unicode_record_data_bytes_len("UniStream"));

        push_record(&mut out, 0x0021, &0u16.to_le_bytes()); // MODULETYPE

        out
    };

    let vba_bin = build_vba_bin_with_dir_decompressed(&dir_decompressed);
    let normalized =
        project_normalized_data_v3_dir_records(&vba_bin).expect("ProjectNormalizedDataV3");

    let expected = [
        utf16le_bytes("UniDoc"),
        utf16le_bytes("UniMod"),
        utf16le_bytes("UniStream"),
        0u16.to_le_bytes().to_vec(),
    ]
    .concat();

    assert_eq!(normalized, expected);
}

#[test]
fn project_normalized_data_v3_strips_unicode_length_prefix_when_prefix_excludes_trailing_nul() {
    // Some producers embed an internal u32 length prefix in Unicode dir record payloads where the
    // payload also includes a trailing UTF-16 NUL terminator that is *not* counted by the prefix.
    // Ensure we strip both the prefix and the uncounted terminator bytes.
    let dir_decompressed = {
        let mut out = Vec::new();

        // Both ANSI and Unicode project docstring records; v3 should emit only Unicode payload bytes.
        push_record(&mut out, 0x0005, b"AnsiDoc");
        // Prefix is a byte count, excluding trailing UTF-16 NUL.
        push_record(
            &mut out,
            0x0040,
            &unicode_record_data_bytes_len_excluding_trailing_nul("UniDoc"),
        );

        // Module group with both MODULENAME and MODULENAMEUNICODE.
        push_record(&mut out, 0x0019, b"AnsiMod");
        // Prefix is a UTF-16 code unit count, excluding trailing UTF-16 NUL.
        push_record(
            &mut out,
            0x0047,
            &unicode_record_data_code_units_excluding_trailing_nul("UniMod"),
        );

        // Both ANSI and Unicode MODULESTREAMNAME records.
        let mut stream_name = Vec::new();
        stream_name.extend_from_slice(b"AnsiStream");
        stream_name.extend_from_slice(&0u16.to_le_bytes());
        push_record(&mut out, 0x001A, &stream_name);
        // Prefix is a byte count, excluding trailing UTF-16 NUL.
        push_record(
            &mut out,
            0x0032,
            &unicode_record_data_bytes_len_excluding_trailing_nul("UniStream"),
        );

        push_record(&mut out, 0x0021, &0u16.to_le_bytes()); // MODULETYPE

        out
    };

    let vba_bin = build_vba_bin_with_dir_decompressed(&dir_decompressed);
    let normalized =
        project_normalized_data_v3_dir_records(&vba_bin).expect("ProjectNormalizedDataV3");

    let expected = [
        utf16le_bytes("UniDoc"),
        utf16le_bytes("UniMod"),
        utf16le_bytes("UniStream"),
        0u16.to_le_bytes().to_vec(),
    ]
    .concat();

    assert_eq!(normalized, expected);
}

#[test]
fn project_normalized_data_v3_strips_trailing_utf16_nul_terminator_from_unicode_records() {
    // Some producers include a UTF-16 NUL terminator at the end of Unicode dir record payloads,
    // regardless of whether an internal length prefix is present or whether it counts the
    // terminator. Ensure we strip a single trailing terminator so the transcript is stable.
    let dir_decompressed = {
        let mut out = Vec::new();

        // PROJECTDOCSTRINGUNICODE with an internal *byte-count* length prefix that *includes* the
        // trailing terminator.
        push_record(&mut out, 0x0005, b"AnsiDoc");
        push_record(
            &mut out,
            0x0040,
            &unicode_record_data_bytes_len_including_trailing_nul("UniDoc"),
        );

        // MODULENAMEUNICODE with an internal *code-unit-count* length prefix that *includes* the
        // trailing terminator.
        push_record(&mut out, 0x0019, b"AnsiMod");
        push_record(
            &mut out,
            0x0047,
            &unicode_record_data_code_units_including_trailing_nul("UniMod"),
        );

        // MODULESTREAMNAMEUNICODE with *no* internal prefix, but a trailing terminator.
        let mut stream_name = Vec::new();
        stream_name.extend_from_slice(b"AnsiStream");
        stream_name.extend_from_slice(&0u16.to_le_bytes());
        push_record(&mut out, 0x001A, &stream_name);
        push_record(&mut out, 0x0032, &utf16le_bytes_with_trailing_nul("UniStream"));

        push_record(&mut out, 0x0021, &0u16.to_le_bytes()); // MODULETYPE

        out
    };

    let vba_bin = build_vba_bin_with_dir_decompressed(&dir_decompressed);
    let normalized =
        project_normalized_data_v3_dir_records(&vba_bin).expect("ProjectNormalizedDataV3");

    let expected = [
        utf16le_bytes("UniDoc"),
        utf16le_bytes("UniMod"),
        utf16le_bytes("UniStream"),
        0u16.to_le_bytes().to_vec(),
    ]
    .concat();

    assert_eq!(normalized, expected);
}

#[test]
fn project_normalized_data_v3_accepts_unicode_records_without_length_prefix() {
    // Some `VBA/dir` Unicode record variants are observed as raw UTF-16LE bytes without an internal
    // u32 length prefix. Ensure we incorporate the payload bytes unchanged.
    let dir_decompressed = {
        let mut out = Vec::new();

        // Both ANSI and Unicode project docstring records; v3 should emit only Unicode payload bytes.
        push_record(&mut out, 0x0005, b"AnsiDoc");
        push_record(&mut out, 0x0040, &utf16le_bytes("UniDoc"));

        // Module group with both MODULENAME and MODULENAMEUNICODE.
        push_record(&mut out, 0x0019, b"AnsiMod");
        push_record(&mut out, 0x0047, &utf16le_bytes("UniMod"));

        // Both ANSI and Unicode MODULESTREAMNAME records.
        let mut stream_name = Vec::new();
        stream_name.extend_from_slice(b"AnsiStream");
        stream_name.extend_from_slice(&0u16.to_le_bytes());
        push_record(&mut out, 0x001A, &stream_name);
        push_record(&mut out, 0x0032, &utf16le_bytes("UniStream"));

        push_record(&mut out, 0x0021, &0u16.to_le_bytes()); // MODULETYPE

        out
    };

    let vba_bin = build_vba_bin_with_dir_decompressed(&dir_decompressed);
    let normalized =
        project_normalized_data_v3_dir_records(&vba_bin).expect("ProjectNormalizedDataV3");

    let expected = [
        utf16le_bytes("UniDoc"),
        utf16le_bytes("UniMod"),
        utf16le_bytes("UniStream"),
        0u16.to_le_bytes().to_vec(),
    ]
    .concat();

    assert_eq!(normalized, expected);
}

#[test]
fn project_normalized_data_v3_prefers_unicode_for_noncanonical_project_record_ids() {
    // Some producers use non-canonical record ids for the Unicode project string variants. Our v3
    // dir-record helper should still prefer them over the ANSI records.
    let dir_decompressed = {
        let mut out = Vec::new();

        // PROJECTNAME
        push_record(&mut out, 0x0004, b"MyProject");

        // PROJECTDOCSTRING (ANSI) + non-canonical Unicode variant.
        push_record(&mut out, 0x0005, b"AnsiDoc");
        push_record(&mut out, 0x0041, &unicode_record_data("UniDoc"));

        // PROJECTCONSTANTS (ANSI) + non-canonical Unicode variant.
        push_record(&mut out, 0x000C, b"AnsiConst");
        push_record(&mut out, 0x0043, &unicode_record_data("UniConst"));

        out
    };

    let vba_bin = build_vba_bin_with_dir_decompressed(&dir_decompressed);
    let normalized =
        project_normalized_data_v3_dir_records(&vba_bin).expect("ProjectNormalizedDataV3");

    let expected = [
        b"MyProject".as_slice(),
        utf16le_bytes("UniDoc").as_slice(),
        utf16le_bytes("UniConst").as_slice(),
    ]
    .concat();

    assert_eq!(normalized, expected);
    assert!(
        !normalized.windows(b"AnsiDoc".len()).any(|w| w == b"AnsiDoc"),
        "expected ANSI PROJECTDOCSTRING bytes to be omitted when Unicode variant is present"
    );
    assert!(
        !normalized
            .windows(b"AnsiConst".len())
            .any(|w| w == b"AnsiConst"),
        "expected ANSI PROJECTCONSTANTS bytes to be omitted when Unicode variant is present"
    );
}

#[test]
fn project_normalized_data_v3_handles_unicode_only_modulename_group_start() {
    // Some real-world (non-spec) dir encodings omit MODULENAME (0x0019) entirely and emit only
    // MODULENAMEUNICODE (0x0047). Ensure the metadata transcript still treats this as the start of a
    // module record group so that Unicode-vs-ANSI preference works for subsequent module records.
    let dir_decompressed = {
        let mut out = Vec::new();

        // Module group described only by MODULENAMEUNICODE.
        push_record(&mut out, 0x0047, &unicode_record_data("UniMod"));

        // Provide both ANSI and Unicode stream name records; Unicode should win, and ANSI must be
        // omitted from the output.
        let mut stream_name = Vec::new();
        stream_name.extend_from_slice(b"AnsiStream");
        stream_name.extend_from_slice(&0u16.to_le_bytes());
        push_record(&mut out, 0x001A, &stream_name);
        push_record(&mut out, 0x0032, &unicode_record_data("UniStream"));

        push_record(&mut out, 0x0021, &0u16.to_le_bytes()); // MODULETYPE

        out
    };

    let vba_bin = build_vba_bin_with_dir_decompressed(&dir_decompressed);
    let normalized = project_normalized_data_v3_dir_records(&vba_bin).expect("ProjectNormalizedDataV3");

    let expected = [
        utf16le_bytes("UniMod"),
        utf16le_bytes("UniStream"),
        0u16.to_le_bytes().to_vec(),
    ]
    .concat();

    assert_eq!(normalized, expected);
    assert!(
        !normalized
            .windows(b"AnsiStream".len())
            .any(|w| w == b"AnsiStream"),
        "expected ANSI MODULESTREAMNAME bytes to be omitted when MODULESTREAMNAMEUNICODE is present"
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
fn project_normalized_data_multiple_baseclass_entries_preserve_order_and_precede_tokens() {
    // Regression test for MS-OVBA §2.4.2.6 `NormalizeProjectStream` ordering:
    // - Multiple `BaseClass=` (`ProjectDesignerModule`) properties MUST be processed in PROJECT
    //   stream order (not sorted, not deduped, not "first only").
    // - For each BaseClass property, `NormalizeDesignerStorage()` bytes MUST appear before the
    //   property name/value token bytes for that property.

    let cursor = Cursor::new(Vec::new());
    let mut ole = cfb::CompoundFile::create(cursor).expect("create cfb");

    // PROJECT stream: deliberate non-lexicographic order.
    {
        let mut s = ole.create_stream("PROJECT").expect("PROJECT stream");
        s.write_all(b"BaseClass=FormB\r\nBaseClass=FormA\r\n")
            .expect("write PROJECT");
    }

    // Minimal VBA/dir mapping so FormsNormalizedData can resolve the designer module identifiers
    // to storage names (MODULENAME -> MODULESTREAMNAME).
    ole.create_storage("VBA").expect("VBA storage");
    {
        let dir_decompressed = {
            let mut out = Vec::new();
            // PROJECTCODEPAGE (u16 LE): Windows-1252.
            push_record(&mut out, 0x0003, &1252u16.to_le_bytes());

            // Intentionally *not* in PROJECT stream order to ensure BaseClass ordering is derived
            // from PROJECT, not from VBA/dir module record group order.
            for module in [b"FormA".as_slice(), b"FormB".as_slice()] {
                push_record(&mut out, 0x0019, module); // MODULENAME
                let mut stream_name = Vec::new();
                stream_name.extend_from_slice(module);
                stream_name.extend_from_slice(&0u16.to_le_bytes());
                push_record(&mut out, 0x001A, &stream_name); // MODULESTREAMNAME (+ reserved u16)
                push_record(&mut out, 0x0021, &3u16.to_le_bytes()); // MODULETYPE (UserForm)
            }
            out
        };
        let dir_container = compress_container(&dir_decompressed);
        let mut s = ole.create_stream("VBA/dir").expect("dir stream");
        s.write_all(&dir_container).expect("write dir");
    }

    // Root-level designer storages and distinct stream bytes.
    ole.create_storage("FormA").expect("FormA storage");
    {
        let mut s = ole.create_stream("FormA/Payload").expect("FormA stream");
        s.write_all(b"A").expect("write FormA bytes");
    }
    ole.create_storage("FormB").expect("FormB storage");
    {
        let mut s = ole.create_stream("FormB/Payload").expect("FormB stream");
        s.write_all(b"B").expect("write FormB bytes");
    }

    let vba_project_bin = ole.into_inner().into_inner();
    let normalized = project_normalized_data(&vba_project_bin).expect("ProjectNormalizedData");

    let mut formb_padded = Vec::new();
    formb_padded.extend_from_slice(b"B");
    formb_padded.extend(std::iter::repeat_n(0u8, 1022));

    let mut forma_padded = Vec::new();
    forma_padded.extend_from_slice(b"A");
    forma_padded.extend(std::iter::repeat_n(0u8, 1022));

    let idx_formb = find_subslice(&normalized, &formb_padded).expect("FormB designer bytes");
    let idx_forma = find_subslice(&normalized, &forma_padded).expect("FormA designer bytes");
    assert!(idx_formb < idx_forma, "expected FormB before FormA (PROJECT order)");
    assert_eq!(
        normalized
            .windows(formb_padded.len())
            .filter(|w| *w == formb_padded.as_slice())
            .count(),
        1,
        "expected FormB normalized designer bytes to appear exactly once"
    );
    assert_eq!(
        normalized
            .windows(forma_padded.len())
            .filter(|w| *w == forma_padded.as_slice())
            .count(),
        1,
        "expected FormA normalized designer bytes to appear exactly once"
    );

    // Assert that the BaseClass property *tokens* preserve PROJECT stream order too.
    //
    // This is stricter than just counting `BaseClass` occurrences: it ensures we don't (for example)
    // sort BaseClass properties or incorrectly associate designer bytes with the wrong token group.
    let idx_base_formb =
        find_subslice(&normalized, b"BaseClassFormB").expect("BaseClass tokens for FormB");
    let idx_base_forma =
        find_subslice(&normalized, b"BaseClassFormA").expect("BaseClass tokens for FormA");
    assert!(
        idx_base_formb < idx_base_forma,
        "expected BaseClass tokens for FormB to precede FormA (PROJECT order)"
    );
    assert!(
        idx_base_formb < idx_forma,
        "expected BaseClass tokens for FormB to be emitted before FormA designer bytes (per-property ordering)"
    );
    let baseclass_occurrences = normalized
        .windows(b"BaseClass".len())
        .filter(|w| *w == b"BaseClass")
        .count();
    assert_eq!(
        baseclass_occurrences, 2,
        "expected exactly two BaseClass property token occurrences"
    );

    assert!(
        idx_formb < idx_base_formb,
        "expected FormB designer bytes before BaseClass tokens for FormB"
    );
    assert!(
        idx_forma < idx_base_forma,
        "expected FormA designer bytes before BaseClass tokens for FormA"
    );
}

#[test]
fn project_normalized_data_interleaves_baseclass_designer_bytes_with_other_properties() {
    // MS-OVBA §2.4.2.6 `NormalizeProjectStream` processes ProjectProperties line-by-line. For a
    // `BaseClass=` property it appends `NormalizeDesignerStorage()` output *before* the property's
    // name/value token bytes, and then continues with the next property.
    //
    // Regression: ensure we interleave designer bytes and tokens per-property, rather than emitting
    // all designer bytes first and all property tokens later.

    let cursor = Cursor::new(Vec::new());
    let mut ole = cfb::CompoundFile::create(cursor).expect("create cfb");

    // PROJECT stream order:
    // - Name (non-designer property)
    // - BaseClass=FormB (designer property)
    // - HelpFile (non-designer property)
    // - BaseClass=FormA (designer property)
    {
        let mut s = ole.create_stream("PROJECT").expect("PROJECT stream");
        s.write_all(
            b"Name=\"VBAProject\"\r\nBaseClass=FormB\r\nHelpFile=\"c:\\foo\"\r\nBaseClass=FormA\r\n",
        )
        .expect("write PROJECT");
    }

    // Minimal VBA/dir mapping for BaseClass module identifiers -> storage names.
    ole.create_storage("VBA").expect("VBA storage");
    {
        let dir_decompressed = {
            let mut out = Vec::new();
            push_record(&mut out, 0x0003, &1252u16.to_le_bytes()); // PROJECTCODEPAGE
            for module in [b"FormA".as_slice(), b"FormB".as_slice()] {
                push_record(&mut out, 0x0019, module); // MODULENAME
                let mut stream_name = Vec::new();
                stream_name.extend_from_slice(module);
                stream_name.extend_from_slice(&0u16.to_le_bytes());
                push_record(&mut out, 0x001A, &stream_name); // MODULESTREAMNAME (+ reserved u16)
                push_record(&mut out, 0x0021, &3u16.to_le_bytes()); // MODULETYPE (UserForm)
            }
            out
        };
        let dir_container = compress_container(&dir_decompressed);
        let mut s = ole.create_stream("VBA/dir").expect("dir stream");
        s.write_all(&dir_container).expect("write dir");
    }

    // Designer storages referenced by BaseClass lines.
    ole.create_storage("FormA").expect("FormA storage");
    {
        let mut s = ole.create_stream("FormA/Payload").expect("FormA stream");
        s.write_all(b"A").expect("write FormA bytes");
    }
    ole.create_storage("FormB").expect("FormB storage");
    {
        let mut s = ole.create_stream("FormB/Payload").expect("FormB stream");
        s.write_all(b"B").expect("write FormB bytes");
    }

    let vba_project_bin = ole.into_inner().into_inner();
    let normalized = project_normalized_data(&vba_project_bin).expect("ProjectNormalizedData");

    let mut formb_padded = Vec::new();
    formb_padded.extend_from_slice(b"B");
    formb_padded.extend(std::iter::repeat_n(0u8, 1022));

    let mut forma_padded = Vec::new();
    forma_padded.extend_from_slice(b"A");
    forma_padded.extend(std::iter::repeat_n(0u8, 1022));

    let idx_name = find_subslice(&normalized, b"NameVBAProject").expect("Name tokens");
    let idx_formb = find_subslice(&normalized, &formb_padded).expect("FormB designer bytes");
    let idx_base_formb =
        find_subslice(&normalized, b"BaseClassFormB").expect("BaseClassFormB tokens");
    let idx_help =
        find_subslice(&normalized, b"HelpFilec:\\foo").expect("HelpFile tokens (quotes stripped)");
    let idx_forma = find_subslice(&normalized, &forma_padded).expect("FormA designer bytes");
    let idx_base_forma =
        find_subslice(&normalized, b"BaseClassFormA").expect("BaseClassFormA tokens");

    assert!(
        idx_name < idx_formb,
        "expected Name tokens before any BaseClass designer bytes"
    );
    assert!(
        idx_formb < idx_base_formb,
        "expected FormB designer bytes before BaseClassFormB tokens"
    );
    assert!(
        idx_base_formb < idx_help,
        "expected BaseClassFormB tokens before subsequent non-designer property tokens"
    );
    assert!(
        idx_help < idx_forma,
        "expected interleaved properties: HelpFile tokens before FormA designer bytes"
    );
    assert!(
        idx_forma < idx_base_forma,
        "expected FormA designer bytes before BaseClassFormA tokens"
    );
}

#[test]
fn project_normalized_data_baseclass_strips_quotes_and_whitespace_for_designer_lookup() {
    // Regression: BaseClass values may be quoted and/or surrounded by whitespace. We must:
    // - strip surrounding quotes and whitespace for the BaseClass value when producing tokens, and
    // - still resolve the designer storage and include its normalized bytes.

    let cursor = Cursor::new(Vec::new());
    let mut ole = cfb::CompoundFile::create(cursor).expect("create cfb");

    // PROJECT stream: whitespace around '=' and quoted value.
    {
        let mut s = ole.create_stream("PROJECT").expect("PROJECT stream");
        s.write_all(b"  BaseClass = \"FormB\"  \r\n")
            .expect("write PROJECT");
    }

    // Minimal VBA/dir mapping (MODULENAME -> MODULESTREAMNAME).
    ole.create_storage("VBA").expect("VBA storage");
    {
        let dir_decompressed = {
            let mut out = Vec::new();
            push_record(&mut out, 0x0003, &1252u16.to_le_bytes()); // PROJECTCODEPAGE

            push_record(&mut out, 0x0019, b"FormB"); // MODULENAME
            let mut stream_name = Vec::new();
            stream_name.extend_from_slice(b"FormB");
            stream_name.extend_from_slice(&0u16.to_le_bytes());
            push_record(&mut out, 0x001A, &stream_name); // MODULESTREAMNAME (+ reserved u16)
            push_record(&mut out, 0x0021, &3u16.to_le_bytes()); // MODULETYPE (UserForm)
            out
        };
        let dir_container = compress_container(&dir_decompressed);
        let mut s = ole.create_stream("VBA/dir").expect("dir stream");
        s.write_all(&dir_container).expect("write dir");
    }

    // Designer storage referenced by BaseClass=FormB.
    ole.create_storage("FormB").expect("FormB storage");
    {
        let mut s = ole.create_stream("FormB/Payload").expect("FormB stream");
        s.write_all(b"B").expect("write FormB bytes");
    }

    let vba_project_bin = ole.into_inner().into_inner();
    let normalized = project_normalized_data(&vba_project_bin).expect("ProjectNormalizedData");

    let mut formb_padded = Vec::new();
    formb_padded.extend_from_slice(b"B");
    formb_padded.extend(std::iter::repeat_n(0u8, 1022));

    let idx_formb = find_subslice(&normalized, &formb_padded).expect("FormB designer bytes");
    let idx_base_formb =
        find_subslice(&normalized, b"BaseClassFormB").expect("BaseClassFormB tokens");

    assert!(
        idx_formb < idx_base_formb,
        "expected normalized designer bytes before BaseClassFormB tokens"
    );
    assert!(
        !normalized.contains(&b'"'),
        "expected ProjectNormalizedData to strip quotes from BaseClass value"
    );
}

#[test]
fn project_normalized_data_matches_baseclass_key_case_insensitively_but_preserves_token_case() {
    // Regression: some writers emit `baseclass=` (lowercase key). We should:
    // - still treat it as a BaseClass designer module declaration (case-insensitive match), but
    // - preserve the original key bytes in the emitted name token (no case normalization).

    let cursor = Cursor::new(Vec::new());
    let mut ole = cfb::CompoundFile::create(cursor).expect("create cfb");

    {
        let mut s = ole.create_stream("PROJECT").expect("PROJECT stream");
        s.write_all(b"baseclass=FormB\r\n")
            .expect("write PROJECT");
    }

    ole.create_storage("VBA").expect("VBA storage");
    {
        let dir_decompressed = {
            let mut out = Vec::new();
            push_record(&mut out, 0x0003, &1252u16.to_le_bytes()); // PROJECTCODEPAGE

            push_record(&mut out, 0x0019, b"FormB"); // MODULENAME
            let mut stream_name = Vec::new();
            stream_name.extend_from_slice(b"FormB");
            stream_name.extend_from_slice(&0u16.to_le_bytes());
            push_record(&mut out, 0x001A, &stream_name); // MODULESTREAMNAME (+ reserved u16)
            push_record(&mut out, 0x0021, &3u16.to_le_bytes()); // MODULETYPE (UserForm)
            out
        };
        let dir_container = compress_container(&dir_decompressed);
        let mut s = ole.create_stream("VBA/dir").expect("dir stream");
        s.write_all(&dir_container).expect("write dir");
    }

    ole.create_storage("FormB").expect("FormB storage");
    {
        let mut s = ole.create_stream("FormB/Payload").expect("FormB stream");
        s.write_all(b"B").expect("write FormB bytes");
    }

    let vba_project_bin = ole.into_inner().into_inner();
    let normalized = project_normalized_data(&vba_project_bin).expect("ProjectNormalizedData");

    let mut formb_padded = Vec::new();
    formb_padded.extend_from_slice(b"B");
    formb_padded.extend(std::iter::repeat_n(0u8, 1022));

    let idx_formb = find_subslice(&normalized, &formb_padded).expect("FormB designer bytes");
    let idx_tokens =
        find_subslice(&normalized, b"baseclassFormB").expect("baseclassFormB tokens");

    assert!(
        idx_formb < idx_tokens,
        "expected designer bytes to appear before the BaseClass property tokens"
    );
    assert!(
        find_subslice(&normalized, b"BaseClassFormB").is_none(),
        "expected name token case to be preserved (no case normalization to `BaseClass`)"
    );
}

#[test]
fn project_normalized_data_matches_baseclass_value_case_insensitively_but_preserves_token_value_case()
{
    // Regression: module identifiers used in `BaseClass=` can differ in ASCII case from the
    // corresponding `VBA/dir` MODULENAME. We should resolve the designer storage in a
    // case-insensitive way, but preserve the original value bytes in the emitted tokens.

    let cursor = Cursor::new(Vec::new());
    let mut ole = cfb::CompoundFile::create(cursor).expect("create cfb");

    // PROJECT stream uses lowercase module identifier.
    {
        let mut s = ole.create_stream("PROJECT").expect("PROJECT stream");
        s.write_all(b"BaseClass=formb\r\n")
            .expect("write PROJECT");
    }

    ole.create_storage("VBA").expect("VBA storage");
    {
        let dir_decompressed = {
            let mut out = Vec::new();
            push_record(&mut out, 0x0003, &1252u16.to_le_bytes()); // PROJECTCODEPAGE

            push_record(&mut out, 0x0019, b"FormB"); // MODULENAME
            let mut stream_name = Vec::new();
            stream_name.extend_from_slice(b"FormB");
            stream_name.extend_from_slice(&0u16.to_le_bytes());
            push_record(&mut out, 0x001A, &stream_name); // MODULESTREAMNAME (+ reserved u16)
            push_record(&mut out, 0x0021, &3u16.to_le_bytes()); // MODULETYPE (UserForm)
            out
        };
        let dir_container = compress_container(&dir_decompressed);
        let mut s = ole.create_stream("VBA/dir").expect("dir stream");
        s.write_all(&dir_container).expect("write dir");
    }

    ole.create_storage("FormB").expect("FormB storage");
    {
        let mut s = ole.create_stream("FormB/Payload").expect("FormB stream");
        s.write_all(b"B").expect("write FormB bytes");
    }

    let vba_project_bin = ole.into_inner().into_inner();
    let normalized = project_normalized_data(&vba_project_bin).expect("ProjectNormalizedData");

    let mut formb_padded = Vec::new();
    formb_padded.extend_from_slice(b"B");
    formb_padded.extend(std::iter::repeat_n(0u8, 1022));

    let idx_formb = find_subslice(&normalized, &formb_padded).expect("FormB designer bytes");
    let idx_tokens =
        find_subslice(&normalized, b"BaseClassformb").expect("BaseClassformb tokens");

    assert!(
        idx_formb < idx_tokens,
        "expected designer bytes to appear before the BaseClass property tokens"
    );
    assert!(
        find_subslice(&normalized, b"BaseClassFormB").is_none(),
        "expected BaseClass value bytes to be preserved (no case normalization)"
    );
}

#[test]
fn project_normalized_data_decodes_baseclass_value_using_project_stream_codepage() {
    // Regression: BaseClass values can contain non-ASCII module identifiers encoded in the VBA
    // project's codepage. `project_normalized_data()` should:
    // - detect `CodePage=` from the PROJECT stream,
    // - decode the BaseClass value using that codepage,
    // - match it to the corresponding `VBA/dir` module record, and
    // - include the referenced designer storage bytes in the transcript.

    let module_name = "Форма1";
    let (module_name_bytes, _, _) = WINDOWS_1251.encode(module_name);

    // Encode the PROJECT stream as Windows-1251, including a non-ASCII BaseClass value.
    let project_stream_text = format!("CodePage=1251\r\nBaseClass={module_name}\r\n");
    let (project_stream_bytes, _, _) = WINDOWS_1251.encode(&project_stream_text);

    let cursor = Cursor::new(Vec::new());
    let mut ole = cfb::CompoundFile::create(cursor).expect("create cfb");

    {
        let mut s = ole.create_stream("PROJECT").expect("PROJECT stream");
        s.write_all(project_stream_bytes.as_ref())
            .expect("write PROJECT");
    }

    // Dir stream includes a conflicting PROJECTCODEPAGE to ensure we prefer the PROJECT stream's
    // CodePage= line.
    ole.create_storage("VBA").expect("VBA storage");
    {
        let dir_decompressed = {
            let mut out = Vec::new();
            push_record(&mut out, 0x0003, &1252u16.to_le_bytes()); // PROJECTCODEPAGE (conflicts)

            // Module record group for the designer module referenced by BaseClass=...
            push_record(&mut out, 0x0019, module_name_bytes.as_ref()); // MODULENAME
            let mut stream_name = Vec::new();
            stream_name.extend_from_slice(module_name_bytes.as_ref());
            stream_name.extend_from_slice(&0u16.to_le_bytes()); // reserved u16
            push_record(&mut out, 0x001A, &stream_name); // MODULESTREAMNAME
            push_record(&mut out, 0x0021, &3u16.to_le_bytes()); // MODULETYPE (UserForm)

            out
        };
        let dir_container = compress_container(&dir_decompressed);
        let mut s = ole.create_stream("VBA/dir").expect("dir stream");
        s.write_all(&dir_container).expect("write dir");
    }

    // Root-level designer storage referenced by BaseClass=... (storage name is MODULESTREAMNAME).
    ole.create_storage(module_name).expect("designer storage");
    {
        let mut s = ole
            .create_stream(format!("{module_name}/Payload"))
            .expect("designer stream");
        s.write_all(b"Q").expect("write designer bytes");
    }

    let vba_project_bin = ole.into_inner().into_inner();
    let normalized = project_normalized_data(&vba_project_bin).expect("ProjectNormalizedData");

    let mut expected_padded = Vec::new();
    expected_padded.extend_from_slice(b"Q");
    expected_padded.extend(std::iter::repeat_n(0u8, 1022));

    let mut expected_tokens = Vec::new();
    expected_tokens.extend_from_slice(b"BaseClass");
    expected_tokens.extend_from_slice(module_name_bytes.as_ref());

    let idx_designer =
        find_subslice(&normalized, &expected_padded).expect("expected designer bytes");
    let idx_tokens =
        find_subslice(&normalized, &expected_tokens).expect("expected BaseClass tokens");
    assert!(
        idx_designer < idx_tokens,
        "expected designer bytes to appear before BaseClass property tokens"
    );
}

#[test]
fn project_normalized_data_falls_back_to_dir_projectcodepage_for_baseclass_value_decoding() {
    // Regression: If the PROJECT stream lacks a `CodePage=` line, we should fall back to
    // PROJECTCODEPAGE (0x0003) from `VBA/dir` when decoding non-ASCII BaseClass values for designer
    // lookup.

    let module_name = "Форма1";
    let (module_name_bytes, _, _) = WINDOWS_1251.encode(module_name);

    // Encode the PROJECT stream as Windows-1251 but *do not* include CodePage=.
    let project_stream_text = format!("BaseClass={module_name}\r\n");
    let (project_stream_bytes, _, _) = WINDOWS_1251.encode(&project_stream_text);

    let cursor = Cursor::new(Vec::new());
    let mut ole = cfb::CompoundFile::create(cursor).expect("create cfb");

    {
        let mut s = ole.create_stream("PROJECT").expect("PROJECT stream");
        s.write_all(project_stream_bytes.as_ref())
            .expect("write PROJECT");
    }

    ole.create_storage("VBA").expect("VBA storage");
    {
        let dir_decompressed = {
            let mut out = Vec::new();
            push_record(&mut out, 0x0003, &1251u16.to_le_bytes()); // PROJECTCODEPAGE

            // Module record group for the designer module referenced by BaseClass=...
            push_record(&mut out, 0x0019, module_name_bytes.as_ref()); // MODULENAME
            let mut stream_name = Vec::new();
            stream_name.extend_from_slice(module_name_bytes.as_ref());
            stream_name.extend_from_slice(&0u16.to_le_bytes()); // reserved u16
            push_record(&mut out, 0x001A, &stream_name); // MODULESTREAMNAME
            push_record(&mut out, 0x0021, &3u16.to_le_bytes()); // MODULETYPE (UserForm)

            out
        };
        let dir_container = compress_container(&dir_decompressed);
        let mut s = ole.create_stream("VBA/dir").expect("dir stream");
        s.write_all(&dir_container).expect("write dir");
    }

    ole.create_storage(module_name).expect("designer storage");
    {
        let mut s = ole
            .create_stream(format!("{module_name}/Payload"))
            .expect("designer stream");
        s.write_all(b"R").expect("write designer bytes");
    }

    let vba_project_bin = ole.into_inner().into_inner();
    let normalized = project_normalized_data(&vba_project_bin).expect("ProjectNormalizedData");

    let mut expected_padded = Vec::new();
    expected_padded.extend_from_slice(b"R");
    expected_padded.extend(std::iter::repeat_n(0u8, 1022));

    let mut expected_tokens = Vec::new();
    expected_tokens.extend_from_slice(b"BaseClass");
    expected_tokens.extend_from_slice(module_name_bytes.as_ref());

    let idx_designer =
        find_subslice(&normalized, &expected_padded).expect("expected designer bytes");
    let idx_tokens =
        find_subslice(&normalized, &expected_tokens).expect("expected BaseClass tokens");
    assert!(
        idx_designer < idx_tokens,
        "expected designer bytes to appear before BaseClass property tokens"
    );
}

#[test]
fn project_normalized_data_resolves_designer_storage_using_modulestreamnameunicode_record() {
    // Regression: BaseClass=... identifies a designer module by MODULENAME, but the corresponding
    // designer *storage* name can come from MODULESTREAMNAMEUNICODE (0x0032) rather than the ANSI
    // MODULESTREAMNAME (0x001A). Ensure we use the Unicode record for storage resolution when
    // present.

    let storage_name = "Форма1";

    let cursor = Cursor::new(Vec::new());
    let mut ole = cfb::CompoundFile::create(cursor).expect("create cfb");

    // PROJECT stream refers to the designer module identifier `NiceName`.
    {
        let mut s = ole.create_stream("PROJECT").expect("PROJECT stream");
        s.write_all(b"BaseClass=NiceName\r\n")
            .expect("write PROJECT");
    }

    ole.create_storage("VBA").expect("VBA storage");
    {
        // Minimal decompressed dir stream describing one module record group:
        // MODULENAME = NiceName
        // MODULESTREAMNAME = Wrong (intentionally)
        // MODULESTREAMNAMEUNICODE = storage_name (correct)
        let dir_decompressed = {
            let mut out = Vec::new();
            push_record(&mut out, 0x0003, &1252u16.to_le_bytes()); // PROJECTCODEPAGE

            push_record(&mut out, 0x0019, b"NiceName"); // MODULENAME

            // MODULESTREAMNAME: intentionally wrong MBCS name, with no trailing reserved u16.
            // When followed immediately by a MODULESTREAMNAMEUNICODE record, `DirStream` treats the
            // 0x0032 record header as the reserved marker and parses the Unicode tail.
            push_record(&mut out, 0x001A, b"Wrong");

            let mut unicode_name = utf16le_bytes(storage_name);
            // Some producers include a trailing UTF-16 NUL; our parser strips it.
            unicode_name.extend_from_slice(&0u16.to_le_bytes());
            push_record(&mut out, 0x0032, &unicode_name); // MODULESTREAMNAMEUNICODE

            push_record(&mut out, 0x0021, &3u16.to_le_bytes()); // MODULETYPE (UserForm)
            out
        };
        let dir_container = compress_container(&dir_decompressed);
        let mut s = ole.create_stream("VBA/dir").expect("dir stream");
        s.write_all(&dir_container).expect("write dir");
    }

    // Create only the Unicode-named designer storage (not the wrong ANSI one).
    ole.create_storage(storage_name)
        .expect("designer storage");
    {
        let mut s = ole
            .create_stream(format!("{storage_name}/Payload"))
            .expect("designer stream");
        s.write_all(b"X").expect("write designer bytes");
    }

    let vba_project_bin = ole.into_inner().into_inner();
    let normalized = project_normalized_data(&vba_project_bin).expect("ProjectNormalizedData");

    let mut expected_padded = Vec::new();
    expected_padded.extend_from_slice(b"X");
    expected_padded.extend(std::iter::repeat_n(0u8, 1022));

    let idx_designer =
        find_subslice(&normalized, &expected_padded).expect("expected designer bytes");
    let idx_tokens = find_subslice(&normalized, b"BaseClassNiceName")
        .expect("expected BaseClass tokens");
    assert!(
        idx_designer < idx_tokens,
        "expected designer bytes to appear before BaseClass property tokens"
    );
}

#[test]
fn project_normalized_data_resolves_designer_storage_using_modulestreamnameunicode_record_id_0048() {
    // Regression: Some `VBA/dir` streams store a Unicode module stream name in a separate record
    // with id 0x0048 immediately following MODULESTREAMNAME (0x001A). `DirStream` supports this
    // nonstandard layout; ensure BaseClass designer storage resolution in `project_normalized_data()`
    // also benefits from it.

    let storage_name = "ユニコード名";

    let cursor = Cursor::new(Vec::new());
    let mut ole = cfb::CompoundFile::create(cursor).expect("create cfb");

    // PROJECT stream refers to the designer module identifier `NiceName`.
    {
        let mut s = ole.create_stream("PROJECT").expect("PROJECT stream");
        s.write_all(b"BaseClass=NiceName\r\n")
            .expect("write PROJECT");
    }

    ole.create_storage("VBA").expect("VBA storage");
    {
        // Minimal decompressed dir stream describing one module record group:
        // MODULENAME = NiceName
        // MODULESTREAMNAME = Wrong (intentionally)
        // 0x0048 Unicode record = storage_name (correct)
        let dir_decompressed = {
            let mut out = Vec::new();
            push_record(&mut out, 0x0003, &1252u16.to_le_bytes()); // PROJECTCODEPAGE

            push_record(&mut out, 0x0019, b"NiceName"); // MODULENAME

            // MODULESTREAMNAME: intentionally wrong MBCS name with trailing reserved u16.
            let mut stream_name = Vec::new();
            stream_name.extend_from_slice(b"Wrong");
            stream_name.extend_from_slice(&0u16.to_le_bytes());
            push_record(&mut out, 0x001A, &stream_name);

            // Nonstandard Unicode stream-name bytes stored in record id 0x0048.
            push_record(&mut out, 0x0048, &unicode_record_data(storage_name));

            push_record(&mut out, 0x0021, &3u16.to_le_bytes()); // MODULETYPE (UserForm)
            out
        };
        let dir_container = compress_container(&dir_decompressed);
        let mut s = ole.create_stream("VBA/dir").expect("dir stream");
        s.write_all(&dir_container).expect("write dir");
    }

    // Create only the Unicode-named designer storage (not the wrong ANSI one).
    ole.create_storage(storage_name)
        .expect("designer storage");
    {
        let mut s = ole
            .create_stream(format!("{storage_name}/Payload"))
            .expect("designer stream");
        s.write_all(b"Z").expect("write designer bytes");
    }

    let vba_project_bin = ole.into_inner().into_inner();
    let normalized = project_normalized_data(&vba_project_bin).expect("ProjectNormalizedData");

    let mut expected_padded = Vec::new();
    expected_padded.extend_from_slice(b"Z");
    expected_padded.extend(std::iter::repeat_n(0u8, 1022));

    let idx_designer =
        find_subslice(&normalized, &expected_padded).expect("expected designer bytes");
    let idx_tokens = find_subslice(&normalized, b"BaseClassNiceName")
        .expect("expected BaseClass tokens");
    assert!(
        idx_designer < idx_tokens,
        "expected designer bytes to appear before BaseClass property tokens"
    );
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

    let mut expected = Vec::new();
    expected.extend_from_slice(b"Y");
    expected.extend(std::iter::repeat_n(0u8, 1022));
    expected.extend_from_slice(b"X");
    expected.extend(std::iter::repeat_n(0u8, 1022));

    let pos = find_subslice(&normalized, &expected)
        .expect("expected normalized output to contain padded designer stream bytes");
    let forms = &normalized[pos..pos + expected.len()];

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
