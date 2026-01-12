use std::io::{Cursor, Write};

use formula_vba::{
    compress_container, content_normalized_data, contents_hash_v3, forms_normalized_data,
    project_normalized_data_v3, v3_content_normalized_data,
};

fn push_record(out: &mut Vec<u8>, id: u16, data: &[u8]) {
    out.extend_from_slice(&id.to_le_bytes());
    out.extend_from_slice(&(data.len() as u32).to_le_bytes());
    out.extend_from_slice(data);
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
        let mut s = ole
            .create_stream("VBA/Module1")
            .expect("module stream");
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

fn build_project_with_projectcompatversion(include_compat: bool) -> (Vec<u8>, [u8; 4]) {
    // Distinctive compat version payload so the regression assertion is unambiguous.
    let compat_version = 0xDEADBEEFu32.to_le_bytes();

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

        // PROJECTVERSION (0x0009): Major(u16) + Minor(u16).
        let mut version = Vec::new();
        version.extend_from_slice(&1u16.to_le_bytes());
        version.extend_from_slice(&0u16.to_le_bytes());
        push_record(&mut out, 0x0009, &version);

        // PROJECTCOMPATVERSION (0x004A): present in many real-world files but must be skipped by
        // the MS-OVBA §2.4.2.5 V3ContentNormalizedData pseudocode.
        if include_compat {
            push_record(&mut out, 0x004A, &compat_version);
        }

        // PROJECTCONSTANTS (0x000C): MBCS string.
        push_record(&mut out, 0x000C, b"Constants");

        // ---- ProjectModules (minimal) ----
        push_record(&mut out, 0x0019, b"Module1"); // MODULENAME

        // MODULESTREAMNAME + reserved u16.
        let mut stream_name = Vec::new();
        stream_name.extend_from_slice(b"Module1");
        stream_name.extend_from_slice(&0u16.to_le_bytes());
        push_record(&mut out, 0x001A, &stream_name);

        push_record(&mut out, 0x0021, &0u16.to_le_bytes()); // MODULETYPE (standard)
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
            .create_stream("VBA/Module1")
            .expect("module stream");
        s.write_all(&module_container).expect("write module");
    }

    (ole.into_inner().into_inner(), compat_version)
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
fn project_normalized_data_v3_appends_padded_forms_normalized_data_when_designer_present() {
    let vba_bin = build_project_with_designer_storage();

    let content_v3 = v3_content_normalized_data(&vba_bin).expect("V3ContentNormalizedData");
    let forms = forms_normalized_data(&vba_bin).expect("FormsNormalizedData");
    let project = project_normalized_data_v3(&vba_bin).expect("ProjectNormalizedData v3");

    let mut expected_content_v3 = Vec::new();
    expected_content_v3.extend_from_slice(b"Sub FormHello()\nEnd Sub\n\nUserForm1\n");
    assert_eq!(content_v3, expected_content_v3);

    let mut expected_forms = Vec::new();
    expected_forms.extend_from_slice(b"ABC");
    expected_forms.extend(std::iter::repeat(0u8).take(1020));
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
}
