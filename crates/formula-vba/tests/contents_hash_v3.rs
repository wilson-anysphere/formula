use std::io::{Cursor, Write};

use formula_vba::{
    compress_container, contents_hash_v3, forms_normalized_data, project_normalized_data_v3_transcript,
    v3_content_normalized_data,
};
use sha2::{Digest as _, Sha256};

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

fn build_two_module_project_v3(module_order: [&str; 2]) -> Vec<u8> {
    // Distinct module source so we can assert ordering in V3ContentNormalizedData directly.
    let module_a_code = b"'MODULE-A\r\nSub A()\r\nEnd Sub\r\n";
    let module_b_code = b"'MODULE-B\r\nSub B()\r\nEnd Sub\r\n";

    let module_a_container = compress_container(module_a_code);
    let module_b_container = compress_container(module_b_code);

    let dir_decompressed = {
        let mut out = Vec::new();

        // REFERENCECONTROL (v3 includes additional reference record types).
        let libid_twiddled = b"REFCTRL-V3";
        let reserved1: u32 = 0;
        let reserved2: u16 = 0;
        let mut reference_control = Vec::new();
        reference_control.extend_from_slice(&(libid_twiddled.len() as u32).to_le_bytes());
        reference_control.extend_from_slice(libid_twiddled);
        reference_control.extend_from_slice(&reserved1.to_le_bytes());
        reference_control.extend_from_slice(&reserved2.to_le_bytes());
        push_record(&mut out, 0x002F, &reference_control);

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

    // Create module streams in alphabetical order (A then B) so we can assert we use the ordering
    // from the `VBA/dir` stream, not the OLE directory enumeration order.
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

fn build_contents_hash_v3_project(module_source: &[u8], designer_stream_bytes: &[u8]) -> Vec<u8> {
    let module_container = compress_container(module_source);
    let userform_code = b"Sub FormHello()\r\nEnd Sub\r\n";
    let userform_container = compress_container(userform_code);

    // `PROJECT` must reference the designer module via `BaseClass=` for FormsNormalizedData.
    let project_stream = b"Name=\"VBAProject\"\r\nBaseClass=\"UserForm1\"\r\n";

    // Minimal decompressed `VBA/dir` stream for:
    // - one procedural module (`Module1`; TypeRecord.Id=0x0021)
    // - one non-procedural module (`UserForm1`; TypeRecord.Id=0x0022) whose designer storage is
    //   `UserForm1/*`
    let dir_decompressed = {
        let mut out = Vec::new();

        // Standard module.
        push_record(&mut out, 0x0019, b"Module1");
        let mut stream_name = Vec::new();
        stream_name.extend_from_slice(b"Module1");
        stream_name.extend_from_slice(&0u16.to_le_bytes());
        push_record(&mut out, 0x001A, &stream_name);
        // Procedural module type record (Id=0x0021): reserved u16
        push_record(&mut out, 0x0021, &0u16.to_le_bytes());
        push_record(&mut out, 0x0031, &0u32.to_le_bytes()); // text offset 0

        // UserForm module (designer). `FormsNormalizedData` will include `UserForm1/*` streams.
        push_record(&mut out, 0x0019, b"UserForm1");
        let mut stream_name = Vec::new();
        stream_name.extend_from_slice(b"UserForm1");
        stream_name.extend_from_slice(&0u16.to_le_bytes());
        push_record(&mut out, 0x001A, &stream_name);
        // Non-procedural module type record (Id=0x0022): reserved u16 (ignored by v3 transcript)
        push_record(&mut out, 0x0022, &0u16.to_le_bytes());
        push_record(&mut out, 0x0031, &0u32.to_le_bytes());

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
        let mut s = ole.create_stream("VBA/Module1").expect("module stream");
        s.write_all(&module_container).expect("write module bytes");
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
        s.write_all(designer_stream_bytes)
            .expect("write designer bytes");
    }

    ole.into_inner().into_inner()
}

#[test]
fn v3_content_normalized_data_includes_v3_reference_records_and_module_metadata_and_uses_dir_order()
{
    // Deliberately non-alphabetical order: B then A.
    let vba_bin = build_two_module_project_v3(["ModuleB", "ModuleA"]);
    let normalized = v3_content_normalized_data(&vba_bin).expect("V3ContentNormalizedData");

    // V3 reference record inclusion.
    assert!(
        normalized.contains(&b'R'),
        "expected V3ContentNormalizedData to include bytes from the REFERENCECONTROL record"
    );
    assert!(
        find_subslice(&normalized, b"REFCTRL-V3").is_some(),
        "expected V3ContentNormalizedData to include REFERENCECONTROL record payload bytes"
    );

    // V3 module metadata inclusion (module identity/metadata).
    assert!(
        find_subslice(&normalized, b"ModuleA").is_some(),
        "expected module name/stream name bytes to be present"
    );
    assert!(
        find_subslice(&normalized, b"ModuleB").is_some(),
        "expected module name/stream name bytes to be present"
    );

    // Ordering must follow module ordering as recorded in `VBA/dir`.
    let pos_b = find_subslice(&normalized, b"'MODULE-B").expect("ModuleB code should be present");
    let pos_a = find_subslice(&normalized, b"'MODULE-A").expect("ModuleA code should be present");
    assert!(
        pos_b < pos_a,
        "expected ModuleB bytes to appear before ModuleA bytes in V3ContentNormalizedData"
    );

    // Swapping module order in the `dir` stream should swap the order in the normalized data too.
    let vba_bin_swapped = build_two_module_project_v3(["ModuleA", "ModuleB"]);
    let normalized_swapped =
        v3_content_normalized_data(&vba_bin_swapped).expect("V3ContentNormalizedData");

    assert_ne!(
        normalized, normalized_swapped,
        "changing module stored order should change V3ContentNormalizedData"
    );

    let pos_a2 =
        find_subslice(&normalized_swapped, b"'MODULE-A").expect("ModuleA code should be present");
    let pos_b2 =
        find_subslice(&normalized_swapped, b"'MODULE-B").expect("ModuleB code should be present");
    assert!(
        pos_a2 < pos_b2,
        "expected ModuleA bytes to appear before ModuleB bytes when dir order is A then B"
    );
}

#[test]
fn v3_content_normalized_data_respects_module_text_offset_when_stream_has_prefix() {
    // Real-world module streams can contain a binary header prefix before the MS-OVBA
    // CompressedContainer. MODULEOFFSET/TextOffset (0x0031) in the `dir` stream indicates where the
    // compressed source begins.
    let prefix_len = 20usize;
    let prefix = vec![0x00u8; prefix_len];

    // Include an Attribute line so we can also confirm module source normalization happens.
    let module_source = concat!(
        "Attribute VB_Name = \"Module1\"\r\n",
        "Sub Hello()\r\n",
        "End Sub\r\n",
    );
    let module_container = compress_container(module_source.as_bytes());

    let mut module_stream = prefix.clone();
    module_stream.extend_from_slice(&module_container);

    let dir_decompressed = {
        let mut out = Vec::new();
        push_record(&mut out, 0x0019, b"Module1"); // MODULENAME
        let mut stream_name = Vec::new();
        stream_name.extend_from_slice(b"Module1");
        stream_name.extend_from_slice(&0u16.to_le_bytes());
        push_record(&mut out, 0x001A, &stream_name); // MODULESTREAMNAME (+ reserved u16)
        push_record(&mut out, 0x0021, &0u16.to_le_bytes()); // MODULETYPE (procedural)
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

    let normalized = v3_content_normalized_data(&vba_bin).expect("V3ContentNormalizedData");
    let mut expected = Vec::new();
    expected.extend_from_slice(&0x0021u16.to_le_bytes());
    expected.extend_from_slice(&0u16.to_le_bytes());
    // V3 module source normalization uses LF separators and appends the module name when
    // `HashModuleNameFlag` becomes true.
    expected.extend_from_slice(b"Sub Hello()\nEnd Sub\n\nModule1\n");
    assert_eq!(normalized, expected);

    // If TextOffset is wrong, we should fail (proves we actually use 0x0031, rather than scanning
    // for a compressed container signature).
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

    let err = v3_content_normalized_data(&vba_bin_wrong_offset).expect_err("expected error");
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
fn project_normalized_data_v3_includes_project_stream_properties_v3_content_and_forms() {
    let vba_bin = build_contents_hash_v3_project(b"Sub Hello()\r\nEnd Sub\r\n", b"ABC");

    let content = v3_content_normalized_data(&vba_bin).expect("V3ContentNormalizedData");
    let forms = forms_normalized_data(&vba_bin).expect("FormsNormalizedData");
    assert!(
        !forms.is_empty(),
        "expected non-empty FormsNormalizedData when PROJECT contains BaseClass= and designer storage is present"
    );
    let project =
        project_normalized_data_v3_transcript(&vba_bin).expect("ProjectNormalizedData v3");

    // ProjectNormalizedData v3 includes filtered PROJECT stream properties before the v3
    // dir/module transcript.
    let mut expected = Vec::new();
    expected.extend_from_slice(b"Name=\"VBAProject\"\r\nBaseClass=\"UserForm1\"\r\n");
    expected.extend_from_slice(&content);
    expected.extend_from_slice(&forms);
    assert_eq!(project, expected);
}

#[test]
fn contents_hash_v3_matches_explicit_normalized_transcript_sha256() {
    // Module source includes:
    // - an Attribute line that must be stripped
    // - mixed newline styles (CRLF / CR-only / lone-LF)
    // - a final line without a newline terminator (must be normalized to LF)
    let module_source = concat!(
        "Attribute VB_Name = \"Module1\"\r\n",
        "Option Explicit\r",
        "Print \"Attribute\"\n",
        "Sub Foo()\r\n",
        "End Sub",
    )
    .as_bytes()
    .to_vec();

    let designer_bytes = b"FORMDATA";
    let vba_project_bin = build_contents_hash_v3_project(&module_source, designer_bytes);

    // ---- Expected normalized transcript ----
    //
    // This test targets `contents_hash_v3`, which computes the MS-OVBA v3 `ContentsHashV3` value:
    //
    // `ContentsHashV3 = SHA-256(ProjectNormalizedData)`
    // `ProjectNormalizedData = (filtered PROJECT stream properties) || V3ContentNormalizedData || FormsNormalizedData`
    //
    // Note: signature binding verification may use other digest algorithms (inferred from the
    // signed digest bytes) for robustness; `contents_hash_v3` always computes the spec-defined
    // SHA-256 digest.
    //
    // V3ContentNormalizedData includes (for procedural modules) `MODULETYPE.Id || MODULETYPE.Reserved`
    // followed by LF-normalized module source (Attribute filtering per MS-OVBA §2.4.2.5), and the
    // module name + LF when `HashModuleNameFlag` becomes true.
    let mut expected = Vec::new();

    // PROJECT stream prefix: Name + BaseClass (both included by the filter).
    expected.extend_from_slice(b"Name=\"VBAProject\"\r\nBaseClass=\"UserForm1\"\r\n");

    // Module1 prefix: (TypeRecord.Id || Reserved)
    expected.extend_from_slice(&0x0021u16.to_le_bytes());
    expected.extend_from_slice(&0u16.to_le_bytes());

    // Module1 normalized source.
    expected.extend_from_slice(
        concat!(
            "Option Explicit\n",
            "Print \"Attribute\"\n",
            "Sub Foo()\n",
            "End Sub\n",
            "Module1\n",
        )
        .as_bytes(),
    );

    // UserForm1 source (non-procedural: TypeRecord bytes omitted). Ends with CRLF in the fixture,
    // and MS-OVBA §2.4.2.5 line splitting appends a trailing empty line, so we expect the extra LF.
    expected.extend_from_slice(b"Sub FormHello()\nEnd Sub\n\nUserForm1\n");

    // FormsNormalizedData: one 1023-byte block for the designer stream.
    expected.extend_from_slice(designer_bytes);
    expected.extend(std::iter::repeat_n(0u8, 1023 - designer_bytes.len()));

    let actual_project_normalized = project_normalized_data_v3_transcript(&vba_project_bin)
        .expect("ProjectNormalizedData v3");
    assert_eq!(
        actual_project_normalized, expected,
        "expected ProjectNormalizedData v3 transcript bytes to match MS-OVBA §2.4.2"
    );

    let actual_digest = contents_hash_v3(&vba_project_bin).expect("ContentsHashV3");
    let expected_digest_from_transcript = Sha256::digest(&expected).to_vec();
    assert_eq!(
        actual_digest, expected_digest_from_transcript,
        "expected ContentsHashV3 to equal SHA-256(ProjectNormalizedData v3)"
    );
    // Hard-coded expected digest bytes to keep this test deterministic and to catch
    // accidental transcript changes.
    let expected_digest: [u8; 32] = [
        0x17, 0x17, 0x57, 0x02, 0x34, 0xf6, 0xf0, 0x2f, 0x26, 0x15, 0x60, 0xef, 0xfd, 0xe2,
        0x37, 0x31, 0x33, 0x75, 0x0d, 0x72, 0x4b, 0x0c, 0x99, 0x1c, 0x5d, 0x99, 0xf7, 0xe6,
        0x7f, 0xe3, 0xb6, 0xea,
    ];
    assert_eq!(
        expected_digest_from_transcript.as_slice(),
        expected_digest.as_ref(),
        "digest constant should match the digest of the explicit transcript above"
    );
    assert_eq!(actual_digest.as_slice(), expected_digest.as_ref());
    assert_eq!(
        Sha256::digest(&expected).as_slice(),
        expected_digest.as_ref(),
        "hard-coded digest must match the explicit normalized transcript"
    );
}

#[test]
fn contents_hash_v3_regressions_attribute_stripping_and_forms_inclusion() {
    let designer_bytes = b"FORMDATA";

    let module_source = concat!(
        "Attribute VB_Name = \"Module1\"\r\n",
        "Option Explicit\r\n",
        "Sub Foo()\r\n",
        "End Sub\r\n",
    )
    .as_bytes()
    .to_vec();
    let vba_project_bin = build_contents_hash_v3_project(&module_source, designer_bytes);
    let digest = contents_hash_v3(&vba_project_bin).expect("base digest");

    // Changing a stripped Attribute line should NOT affect the hash.
    let module_source_attribute_changed = concat!(
        "Attribute VB_Name = \"RenamedModule\"\r\n",
        "Option Explicit\r\n",
        "Sub Foo()\r\n",
        "End Sub\r\n",
    )
    .as_bytes()
    .to_vec();
    let vba_project_bin_attr_changed =
        build_contents_hash_v3_project(&module_source_attribute_changed, designer_bytes);
    let digest_attr_changed =
        contents_hash_v3(&vba_project_bin_attr_changed).expect("digest with attribute line changed");
    assert_eq!(
        digest_attr_changed, digest,
        "stripped Attribute lines must not influence ContentsHashV3"
    );

    // Changing a non-Attribute code line should affect the hash.
    let module_source_code_changed = concat!(
        "Attribute VB_Name = \"Module1\"\r\n",
        "Option Compare Database\r\n",
        "Sub Foo()\r\n",
        "End Sub\r\n",
    )
    .as_bytes()
    .to_vec();
    let vba_project_bin_code_changed =
        build_contents_hash_v3_project(&module_source_code_changed, designer_bytes);
    let digest_code_changed =
        contents_hash_v3(&vba_project_bin_code_changed).expect("digest with code line changed");
    assert_ne!(
        digest_code_changed, digest,
        "non-Attribute code changes must influence ContentsHashV3"
    );

    // Changing designer stream bytes must affect the hash (V3-specific inclusion of FormsNormalizedData).
    let designer_bytes_changed = b"FORMDATA2";
    let vba_project_bin_forms_changed =
        build_contents_hash_v3_project(&module_source, designer_bytes_changed);
    let digest_forms_changed =
        contents_hash_v3(&vba_project_bin_forms_changed).expect("digest with designer bytes changed");
    assert_ne!(
        digest_forms_changed, digest,
        "designer stream bytes must influence ContentsHashV3"
    );
}
