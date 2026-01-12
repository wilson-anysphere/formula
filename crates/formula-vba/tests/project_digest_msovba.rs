use std::io::{Cursor, Write};

use formula_vba::{
    compress_container, compute_vba_project_digest, content_normalized_data, forms_normalized_data,
    DigestAlg,
};
use md5::{Digest as _, Md5};

fn push_record(out: &mut Vec<u8>, id: u16, data: &[u8]) {
    out.extend_from_slice(&id.to_le_bytes());
    out.extend_from_slice(&(data.len() as u32).to_le_bytes());
    out.extend_from_slice(data);
}

#[test]
fn compute_vba_project_digest_errors_when_transcript_cannot_be_computed() {
    // Regression test: `compute_vba_project_digest` should be based on the MS-OVBA transcript
    // (`ContentNormalizedData || FormsNormalizedData`), not a fallback that hashes raw OLE streams.
    //
    // Build a valid OLE container but omit `VBA/dir`, which is required for transcript computation.
    let cursor = Cursor::new(Vec::new());
    let mut ole = cfb::CompoundFile::create(cursor).expect("create cfb");
    ole.create_storage("VBA").expect("VBA storage");
    {
        let mut s = ole
            .create_stream("VBA/Module1")
            .expect("create module stream");
        s.write_all(b"not-a-valid-compressed-module").expect("write");
    }

    let vba_bin = ole.into_inner().into_inner();

    assert!(
        compute_vba_project_digest(&vba_bin, DigestAlg::Md5).is_err(),
        "expected digest computation to fail when MS-OVBA transcript cannot be produced"
    );
}

#[test]
fn compute_vba_project_digest_matches_msovba_transcript_content_plus_forms() {
    // --- Build a minimal, self-contained vbaProject.bin ---
    //
    // ContentNormalizedData inputs:
    // - PROJECTNAME (0x0004)
    // - PROJECTCONSTANTS (0x000C)
    // - one module record group (MODULENAME / MODULESTREAMNAME / MODULETEXTOFFSET)
    //
    // FormsNormalizedData inputs:
    // - one root-level designer storage stream (UserForm1/X) so FormsNormalizedData is non-empty.

    let project_name = b"Project1";
    let project_constants = b"Answer=42";

    // This module also has a matching root-level designer storage (`UserForm1/*`) so that
    // `forms_normalized_data()` includes its streams.
    let module_name = b"UserForm1";

    // Module source includes:
    // - Attribute lines (must be stripped, case-insensitive)
    // - CRLF, CR-only, and lone-LF newlines (must normalize to CRLF)
    // - a final line without a trailing newline (still must be terminated with CRLF)
    let module_source = concat!(
        "Attribute VB_Name = \"UserForm1\"\r\n",
        "Option Explicit\r",
        "Print \"Attribute\"\n",
        "aTtRiBuTe VB_Base = \"0{00000000-0000-0000-0000-000000000000}\"\r\n",
        "Sub Foo()\r\n",
        "End Sub",
    )
    .as_bytes()
    .to_vec();
    let module_container = compress_container(&module_source);

    // Decompressed `VBA/dir` stream data.
    let dir_decompressed = {
        let mut out = Vec::new();
        push_record(&mut out, 0x0004, project_name); // PROJECTNAME
        push_record(&mut out, 0x000C, project_constants); // PROJECTCONSTANTS

        // Module record group (single module).
        push_record(&mut out, 0x0019, module_name); // MODULENAME
        let mut stream_name = Vec::new();
        stream_name.extend_from_slice(module_name);
        stream_name.extend_from_slice(&0u16.to_le_bytes()); // reserved u16
        push_record(&mut out, 0x001A, &stream_name); // MODULESTREAMNAME
        push_record(&mut out, 0x0021, &0x0003u16.to_le_bytes()); // MODULETYPE (UserForm)
        push_record(&mut out, 0x0031, &0u32.to_le_bytes()); // MODULETEXTOFFSET (0)

        out
    };
    let dir_container = compress_container(&dir_decompressed);

    // Designer stream bytes for FormsNormalizedData.
    let designer_stream = b"DESIGNER".to_vec(); // len=8 => pads with 1015 zeros to 1023 bytes

    let cursor = Cursor::new(Vec::new());
    let mut ole = cfb::CompoundFile::create(cursor).expect("create cfb");
    {
        // Minimal PROJECT stream so `forms_normalized_data()` can locate designer modules via
        // `BaseClass=` records (MS-OVBA §2.3.1.7).
        let mut s = ole.create_stream("PROJECT").expect("PROJECT stream");
        s.write_all(b"BaseClass=UserForm1\r\n")
            .expect("write PROJECT");
    }
    ole.create_storage("VBA").expect("VBA storage");
    {
        let mut s = ole.create_stream("VBA/dir").expect("dir stream");
        s.write_all(&dir_container).expect("write dir");
    }
    {
        let mut s = ole
            .create_stream("VBA/UserForm1")
            .expect("module stream");
        s.write_all(&module_container).expect("write module");
    }
    ole.create_storage("UserForm1")
        .expect("UserForm1 designer storage");
    {
        let mut s = ole
            .create_stream("UserForm1/X")
            .expect("UserForm1/X stream");
        s.write_all(&designer_stream).expect("write designer stream");
    }

    let vba_bin = ole.into_inner().into_inner();

    // --- Build expected transcript bytes manually ---
    //
    // ContentNormalizedData (subset we currently implement):
    //   PROJECTNAME.data || PROJECTCONSTANTS.data || NormalizedModuleSource
    let expected_module_normalized = concat!(
        "Option Explicit\r\n",
        "Print \"Attribute\"\r\n",
        "Sub Foo()\r\n",
        "End Sub\r\n",
    )
    .as_bytes()
    .to_vec();

    let expected_content = [
        project_name.as_slice(),
        project_constants.as_slice(),
        expected_module_normalized.as_slice(),
    ]
    .concat();

    // FormsNormalizedData:
    //   concatenate designer stream bytes in lexicographic path order, padding each
    //   stream to a 1023-byte multiple.
    let mut expected_forms = designer_stream.clone();
    let rem = expected_forms.len() % 1023;
    assert_ne!(rem, 0, "test requires non-zero padding");
    expected_forms.extend(std::iter::repeat_n(0u8, 1023 - rem));

    let expected_transcript = [expected_content.as_slice(), expected_forms.as_slice()].concat();

    let expected_md5 = Md5::digest(&expected_transcript).to_vec();
    let actual_md5 = compute_vba_project_digest(&vba_bin, DigestAlg::Md5).expect("digest");

    let actual_content = content_normalized_data(&vba_bin).expect("ContentNormalizedData");
    assert_eq!(actual_content, expected_content);
    let actual_forms = forms_normalized_data(&vba_bin).expect("FormsNormalizedData");
    assert_eq!(actual_forms, expected_forms);

    assert_eq!(actual_md5, expected_md5);
}
