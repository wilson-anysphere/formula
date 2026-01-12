use std::io::{Cursor, Write};

use encoding_rs::WINDOWS_1251;
use formula_vba::{compress_container, project_normalized_data};

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
fn project_normalized_data_project_properties_parse_name_and_value_tokens_without_separators() {
    // Regression test for MS-OVBA §2.4.2.6 `ProjectNormalizedData` `ProjectProperties`:
    //
    // The spec pseudocode appends the **property name token bytes** and then the **property value
    // token bytes**. It's easy to accidentally hash the raw line bytes (e.g. `Name="..."\r\n`),
    // which would incorrectly include separators (`=`), quotes, and newline bytes.

    // Minimal decompressed `VBA/dir` stream: enough to:
    // - satisfy `project_normalized_data()` (it requires `VBA/dir`), and
    // - allow `FormsNormalizedData` to resolve `BaseClass=UserForm1` to the `UserForm1` designer
    //   storage via MODULENAME → MODULESTREAMNAME.
    let dir_decompressed = {
        let mut out = Vec::new();
        // PROJECTCODEPAGE (u16 LE): Windows-1252.
        push_record(&mut out, 0x0003, &1252u16.to_le_bytes());

        // Module group for the designer module identifier `UserForm1`.
        push_record(&mut out, 0x0019, b"UserForm1"); // MODULENAME
        let mut stream_name = Vec::new();
        stream_name.extend_from_slice(b"UserForm1");
        stream_name.extend_from_slice(&0u16.to_le_bytes()); // reserved u16
        push_record(&mut out, 0x001A, &stream_name); // MODULESTREAMNAME
        push_record(&mut out, 0x0021, &0x0003u16.to_le_bytes()); // MODULETYPE (UserForm)
        out
    };
    let dir_container = compress_container(&dir_decompressed);

    let cursor = Cursor::new(Vec::new());
    let mut ole = cfb::CompoundFile::create(cursor).expect("create cfb");

    // `PROJECT` stream properties (MS-OVBA example formatting).
    let project_stream = concat!(
        "Name=\"VBAProject\"\r\n",
        "BaseClass=UserForm1\r\n",
        "HelpFile=\"c:\\example path\\example.hlp\"\r\n",
        "HelpContextID=\"1\"\r\n",
    );
    {
        let mut s = ole.create_stream("PROJECT").expect("PROJECT stream");
        s.write_all(project_stream.as_bytes())
            .expect("write PROJECT");
    }

    ole.create_storage("VBA").expect("VBA storage");
    {
        let mut s = ole.create_stream("VBA/dir").expect("dir stream");
        s.write_all(&dir_container).expect("write dir");
    }

    // Designer storage referenced by `BaseClass=UserForm1`.
    ole.create_storage("UserForm1")
        .expect("create designer storage");
    {
        let mut s = ole
            .create_stream("UserForm1/f")
            .expect("create designer stream");
        s.write_all(b"DESIGNER").expect("write designer bytes");
    }

    let vba_project_bin = ole.into_inner().into_inner();
    let normalized =
        project_normalized_data(&vba_project_bin).expect("compute ProjectNormalizedData");

    // Property tokens must appear as `property_name || property_value` with:
    // - no '=' separators
    // - no surrounding quotes for quoted values
    // - no newline bytes
    let expected_name = b"NameVBAProject";
    let expected_base_class = b"BaseClassUserForm1";
    let expected_help_file = b"HelpFilec:\\example path\\example.hlp";
    let expected_help_context_id = b"HelpContextID1";

    let pos_name =
        find_subslice(&normalized, expected_name).expect("Name property tokens should be present");
    let pos_base_class = find_subslice(&normalized, expected_base_class)
        .expect("BaseClass property tokens should be present");
    let pos_help_file = find_subslice(&normalized, expected_help_file)
        .expect("HelpFile property tokens should be present");
    let pos_help_context_id = find_subslice(&normalized, expected_help_context_id)
        .expect("HelpContextID property tokens should be present");

    assert!(
        pos_name < pos_base_class
            && pos_base_class < pos_help_file
            && pos_help_file < pos_help_context_id,
        "expected ProjectProperties tokens to appear in the same order as in the PROJECT stream"
    );

    // Ensure the raw `=` separators from the input lines are not present.
    assert!(
        !normalized.windows(b"Name=".len()).any(|w| w == b"Name="),
        "expected ProjectNormalizedData to omit the '=' separator from the Name line"
    );
    assert!(
        !normalized
            .windows(b"BaseClass=".len())
            .any(|w| w == b"BaseClass="),
        "expected ProjectNormalizedData to omit the '=' separator from the BaseClass line"
    );
    assert!(
        !normalized
            .windows(b"HelpFile=".len())
            .any(|w| w == b"HelpFile="),
        "expected ProjectNormalizedData to omit the '=' separator from the HelpFile line"
    );
    assert!(
        !normalized
            .windows(b"HelpContextID=".len())
            .any(|w| w == b"HelpContextID="),
        "expected ProjectNormalizedData to omit the '=' separator from the HelpContextID line"
    );

    // Ensure quoted values are appended without quote bytes, and newline bytes are not carried over
    // from the `PROJECT` stream line endings.
    assert!(
        !normalized.contains(&b'"'),
        "expected ProjectNormalizedData to omit quotes from quoted property values"
    );
    assert!(
        !normalized.contains(&b'\r') && !normalized.contains(&b'\n'),
        "expected ProjectNormalizedData to omit NWLN bytes from PROJECT stream property lines"
    );

    // `BaseClass` must also contribute the normalized designer storage bytes.
    let mut expected_designer_storage = Vec::new();
    expected_designer_storage.extend_from_slice(b"DESIGNER");
    expected_designer_storage.extend(std::iter::repeat(0u8).take(1023 - b"DESIGNER".len()));

    let pos_designer = find_subslice(&normalized, &expected_designer_storage)
        .expect("expected NormalizeDesignerStorage(UserForm1) bytes to be present");
    assert!(
        pos_base_class < pos_designer,
        "expected BaseClass property tokens to appear before the normalized designer storage bytes"
    );
}

#[test]
fn project_normalized_data_project_properties_preserves_non_ascii_mbcs_bytes_verbatim() {
    // Regression test: `ProjectProperties` name/value tokens are appended as **raw MBCS bytes**.
    // A naive implementation might decode to a Rust `String` and then append UTF-8 bytes, which
    // would corrupt the binding transcript for non-ASCII projects.

    let project_name = "Проект"; // "project" in Russian (non-ASCII)
    let (mbcs, _, _) = WINDOWS_1251.encode(project_name);
    let mbcs_bytes = mbcs.as_ref();
    let utf8_bytes = project_name.as_bytes();
    assert_ne!(
        mbcs_bytes, utf8_bytes,
        "test precondition: Windows-1251 bytes must differ from UTF-8 bytes"
    );

    let mut project_stream_bytes = Vec::new();
    project_stream_bytes.extend_from_slice(b"Name=\"");
    project_stream_bytes.extend_from_slice(mbcs_bytes);
    project_stream_bytes.extend_from_slice(b"\"\r\n");

    // `project_normalized_data()` requires a `VBA/dir` stream to exist, but this test does not need
    // any dir record payload bytes.
    let dir_container = compress_container(&[]);

    let cursor = Cursor::new(Vec::new());
    let mut ole = cfb::CompoundFile::create(cursor).expect("create cfb");
    {
        let mut s = ole.create_stream("PROJECT").expect("PROJECT stream");
        s.write_all(&project_stream_bytes)
            .expect("write PROJECT bytes");
    }
    ole.create_storage("VBA").expect("VBA storage");
    {
        let mut s = ole.create_stream("VBA/dir").expect("dir stream");
        s.write_all(&dir_container).expect("write dir");
    }

    let vba_project_bin = ole.into_inner().into_inner();
    let normalized =
        project_normalized_data(&vba_project_bin).expect("compute ProjectNormalizedData");

    let expected = [b"Name".as_slice(), mbcs_bytes].concat();
    let not_expected = [b"Name".as_slice(), utf8_bytes].concat();

    assert_eq!(
        normalized, expected,
        "expected ProjectNormalizedData to equal the concatenated Name token bytes (raw MBCS)"
    );
    assert!(
        find_subslice(&normalized, &not_expected).is_none(),
        "expected ProjectNormalizedData to NOT contain UTF-8 bytes for the same Name value"
    );

    // Negative assertions for the token parsing: ensure no separators/quotes/newlines from the raw
    // `PROJECT` stream line are included.
    assert!(
        !normalized.windows(b"Name=".len()).any(|w| w == b"Name="),
        "expected ProjectNormalizedData to omit '=' separator bytes"
    );
    assert!(
        !normalized.contains(&b'"'),
        "expected ProjectNormalizedData to omit quote bytes"
    );
    assert!(
        !normalized.contains(&b'\r') && !normalized.contains(&b'\n'),
        "expected ProjectNormalizedData to omit NWLN bytes"
    );
}

#[test]
fn project_normalized_data_project_properties_excludes_project_id_property_entirely() {
    // MS-OVBA §2.4.2.6 explicitly excludes the ProjectId (`ID=...`) property from the transcript.
    // This test ensures the entire property is omitted (both the name token and the value bytes).

    let project_stream = concat!(
        "ID=\"{11111111-2222-3333-4444-555555555555}\"\r\n",
        "Name=\"VBAProject\"\r\n",
    );

    // `project_normalized_data()` requires `VBA/dir` to exist, but the contents can be empty.
    let dir_container = compress_container(&[]);

    let cursor = Cursor::new(Vec::new());
    let mut ole = cfb::CompoundFile::create(cursor).expect("create cfb");
    {
        let mut s = ole.create_stream("PROJECT").expect("PROJECT stream");
        s.write_all(project_stream.as_bytes())
            .expect("write PROJECT");
    }
    ole.create_storage("VBA").expect("VBA storage");
    {
        let mut s = ole.create_stream("VBA/dir").expect("dir stream");
        s.write_all(&dir_container).expect("write dir");
    }

    let vba_project_bin = ole.into_inner().into_inner();
    let normalized =
        project_normalized_data(&vba_project_bin).expect("compute ProjectNormalizedData");

    // With an empty `VBA/dir` and no designers, the output should be exactly the Name token bytes.
    assert_eq!(normalized, b"NameVBAProject");
}
