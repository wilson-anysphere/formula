use std::io::{Cursor, Write};

use formula_vba::{compress_container, forms_normalized_data};

fn push_record(out: &mut Vec<u8>, id: u16, data: &[u8]) {
    out.extend_from_slice(&id.to_le_bytes());
    out.extend_from_slice(&(data.len() as u32).to_le_bytes());
    out.extend_from_slice(data);
}

fn build_dir_stream(designer_modules: &[(&str, &str)]) -> Vec<u8> {
    // Build a minimal decompressed `VBA/dir` stream with MODULE record groups for the designer
    // modules. FormsNormalizedData needs MODULENAME (module identifier) → MODULESTREAMNAME (designer
    // storage name mapping).
    let dir_decompressed = {
        let mut out = Vec::new();
        // PROJECTCODEPAGE (u16 LE) - default to Windows-1252.
        push_record(&mut out, 0x0003, &1252u16.to_le_bytes());

        for (module_name, stream_name) in designer_modules {
            // MODULENAME
            push_record(&mut out, 0x0019, module_name.as_bytes());

            // MODULESTREAMNAME + reserved u16 (some producers emit this; our parser trims it).
            let mut stream_name_record = Vec::new();
            stream_name_record.extend_from_slice(stream_name.as_bytes());
            stream_name_record.extend_from_slice(&0u16.to_le_bytes());
            push_record(&mut out, 0x001A, &stream_name_record);

            // MODULETYPE (0x0003 = UserForm / designer module).
            push_record(&mut out, 0x0021, &3u16.to_le_bytes());

            // MODULETEXTOFFSET (unused here, but common in real `dir` streams).
            push_record(&mut out, 0x0031, &0u32.to_le_bytes());
        }

        out
    };

    compress_container(&dir_decompressed)
}

fn build_project_stream_for_designer_modules(module_names_in_order: &[&str]) -> Vec<u8> {
    // MS-OVBA §2.3.1.7 ProjectDesignerModule:
    //   BaseClass=<ModuleIdentifier>
    //
    // FormsNormalizedData iterates these properties in PROJECT stream order.
    let mut s = String::new();
    s.push_str("ID=\"{00000000-0000-0000-0000-000000000000}\"\r\n");
    for name in module_names_in_order {
        s.push_str("BaseClass=");
        s.push_str(name);
        s.push_str("\r\n");
    }
    s.into_bytes()
}

fn utf16le_bytes(s: &str) -> Vec<u8> {
    let mut out = Vec::new();
    for u in s.encode_utf16() {
        out.extend_from_slice(&u.to_le_bytes());
    }
    out
}

#[test]
fn forms_normalized_data_pads_stream_to_1023_byte_blocks() {
    let cursor = Cursor::new(Vec::new());
    let mut ole = cfb::CompoundFile::create(cursor).expect("create compound file");

    {
        let project_bytes = build_project_stream_for_designer_modules(&["UserForm1"]);
        let mut s = ole.create_stream("PROJECT").expect("PROJECT stream");
        s.write_all(&project_bytes).expect("write PROJECT");
    }
    ole.create_storage("VBA").expect("create VBA storage");
    {
        let dir_container = build_dir_stream(&[("UserForm1", "UserForm1")]);
        let mut s = ole.create_stream("VBA/dir").expect("dir stream");
        s.write_all(&dir_container).expect("write dir");
    }

    ole.create_storage("UserForm1")
        .expect("create designer storage");
    {
        let mut s = ole
            .create_stream("UserForm1/Payload")
            .expect("create stream");
        s.write_all(b"ABC").expect("write stream bytes");
    }

    let vba_project_bin = ole.into_inner().into_inner();
    let normalized = forms_normalized_data(&vba_project_bin).expect("compute FormsNormalizedData");

    // MS-OVBA pads the final block to 1023 bytes with zeros.
    let mut expected = Vec::new();
    expected.extend_from_slice(b"ABC");
    expected.extend(std::iter::repeat(0u8).take(1020));

    assert_eq!(normalized.len(), 1023);
    assert_eq!(normalized, expected);
}

#[test]
fn forms_normalized_data_traverses_nested_storages_in_storage_element_order() {
    let cursor = Cursor::new(Vec::new());
    let mut ole = cfb::CompoundFile::create(cursor).expect("create compound file");

    {
        let project_bytes = build_project_stream_for_designer_modules(&["UserForm1"]);
        let mut s = ole.create_stream("PROJECT").expect("PROJECT stream");
        s.write_all(&project_bytes).expect("write PROJECT");
    }
    ole.create_storage("VBA").expect("create VBA storage");
    {
        let dir_container = build_dir_stream(&[("UserForm1", "UserForm1")]);
        let mut s = ole.create_stream("VBA/dir").expect("dir stream");
        s.write_all(&dir_container).expect("write dir");
    }

    ole.create_storage("UserForm1")
        .expect("create designer storage");
    ole.create_storage("UserForm1/Child")
        .expect("create nested storage");

    // Write the sibling stream first, then the nested one; enumeration order should come from the
    // compound file's directory tree ordering, not insertion order.
    {
        let mut s = ole
            .create_stream("UserForm1/Y")
            .expect("create sibling stream");
        s.write_all(b"Y").expect("write Y");
    }
    {
        let mut s = ole
            .create_stream("UserForm1/Child/X")
            .expect("create nested stream");
        s.write_all(b"X").expect("write X");
    }

    let vba_project_bin = ole.into_inner().into_inner();
    let normalized = forms_normalized_data(&vba_project_bin).expect("compute FormsNormalizedData");

    // Within a storage, MS-CFB orders siblings by name length first (then case-insensitive code point
    // order). With names `Y` (len=1) and `Child` (len=5), this yields:
    // - `UserForm1/Y`
    // - `UserForm1/Child/X`
    let mut expected = Vec::new();
    expected.extend_from_slice(b"Y");
    expected.extend(std::iter::repeat(0u8).take(1022));
    expected.extend_from_slice(b"X");
    expected.extend(std::iter::repeat(0u8).take(1022));

    assert_eq!(normalized.len(), 1023 * 2);
    assert_eq!(normalized, expected);
}

#[test]
fn forms_normalized_data_uses_modulestreamname_to_find_designer_storage() {
    let cursor = Cursor::new(Vec::new());
    let mut ole = cfb::CompoundFile::create(cursor).expect("create compound file");

    // Root-level "designer" storage named by MODULESTREAMNAME, not by MODULENAME / BaseClass.
    ole.create_storage("UserForm1")
        .expect("create designer storage");
    {
        let mut s = ole
            .create_stream("UserForm1/Payload")
            .expect("create stream");
        s.write_all(b"ABC").expect("write stream bytes");
    }

    {
        let mut s = ole.create_stream("PROJECT").expect("PROJECT stream");
        s.write_all(b"BaseClass=NiceName\r\n")
            .expect("write PROJECT");
    }
    ole.create_storage("VBA").expect("VBA storage");
    {
        let dir_container = build_dir_stream(&[("NiceName", "UserForm1")]);
        let mut s = ole.create_stream("VBA/dir").expect("dir stream");
        s.write_all(&dir_container).expect("write dir");
    }

    let vba_project_bin = ole.into_inner().into_inner();
    let normalized = forms_normalized_data(&vba_project_bin).expect("compute FormsNormalizedData");

    let mut expected = Vec::new();
    expected.extend_from_slice(b"ABC");
    expected.extend(std::iter::repeat(0u8).take(1020));

    assert_eq!(normalized.len(), 1023);
    assert_eq!(normalized, expected);
}

#[test]
fn forms_normalized_data_uses_modulestreamnameunicode_to_find_designer_storage() {
    let cursor = Cursor::new(Vec::new());
    let mut ole = cfb::CompoundFile::create(cursor).expect("create compound file");

    // Root-level "designer" storage named by MODULESTREAMNAMEUNICODE, not by MODULENAME / BaseClass.
    let storage_name = "Форма1";
    ole.create_storage(storage_name)
        .expect("create designer storage");
    {
        let mut s = ole
            .create_stream(format!("{storage_name}/Payload"))
            .expect("create stream");
        s.write_all(b"ABC").expect("write stream bytes");
    }

    {
        let mut s = ole.create_stream("PROJECT").expect("PROJECT stream");
        s.write_all(b"BaseClass=NiceName\r\n")
            .expect("write PROJECT");
    }
    ole.create_storage("VBA").expect("VBA storage");
    {
        // Provide an incorrect MODULESTREAMNAME but a correct MODULESTREAMNAMEUNICODE.
        let mut dir_decompressed = Vec::new();
        push_record(&mut dir_decompressed, 0x0003, &1252u16.to_le_bytes());
        push_record(&mut dir_decompressed, 0x0019, b"NiceName");

        let mut wrong_stream_name = Vec::new();
        wrong_stream_name.extend_from_slice(b"Wrong");
        wrong_stream_name.extend_from_slice(&0u16.to_le_bytes());
        push_record(&mut dir_decompressed, 0x001A, &wrong_stream_name);

        let mut unicode_name = utf16le_bytes(storage_name);
        // Add a trailing NUL, which some producers emit and our parser strips.
        unicode_name.extend_from_slice(&0u16.to_le_bytes());
        push_record(&mut dir_decompressed, 0x0032, &unicode_name);

        push_record(&mut dir_decompressed, 0x0021, &3u16.to_le_bytes());
        push_record(&mut dir_decompressed, 0x0031, &0u32.to_le_bytes());

        let dir_container = compress_container(&dir_decompressed);
        let mut s = ole.create_stream("VBA/dir").expect("dir stream");
        s.write_all(&dir_container).expect("write dir");
    }

    let vba_project_bin = ole.into_inner().into_inner();
    let normalized = forms_normalized_data(&vba_project_bin).expect("compute FormsNormalizedData");

    let mut expected = Vec::new();
    expected.extend_from_slice(b"ABC");
    expected.extend(std::iter::repeat(0u8).take(1020));

    assert_eq!(normalized.len(), 1023);
    assert_eq!(normalized, expected);
}

#[test]
fn forms_normalized_data_matches_baseclass_case_insensitively() {
    let cursor = Cursor::new(Vec::new());
    let mut ole = cfb::CompoundFile::create(cursor).expect("create compound file");

    ole.create_storage("UserForm1")
        .expect("create designer storage");
    {
        let mut s = ole
            .create_stream("UserForm1/Payload")
            .expect("create stream");
        s.write_all(b"ABC").expect("write stream bytes");
    }

    // Some writers appear to emit the key/value using different casing; matching should be
    // case-insensitive.
    {
        let mut s = ole.create_stream("PROJECT").expect("PROJECT stream");
        s.write_all(b"baseclass=NICENAME\r\n")
            .expect("write PROJECT");
    }
    ole.create_storage("VBA").expect("VBA storage");
    {
        let dir_container = build_dir_stream(&[("NiceName", "UserForm1")]);
        let mut s = ole.create_stream("VBA/dir").expect("dir stream");
        s.write_all(&dir_container).expect("write dir");
    }

    let vba_project_bin = ole.into_inner().into_inner();
    let normalized = forms_normalized_data(&vba_project_bin).expect("compute FormsNormalizedData");

    let mut expected = Vec::new();
    expected.extend_from_slice(b"ABC");
    expected.extend(std::iter::repeat(0u8).take(1020));

    assert_eq!(normalized.len(), 1023);
    assert_eq!(normalized, expected);
}

#[test]
fn forms_normalized_data_uses_project_stream_baseclass_order_and_ignores_unlisted_storages() {
    let cursor = Cursor::new(Vec::new());
    let mut ole = cfb::CompoundFile::create(cursor).expect("create compound file");

    // Deliberately non-alphabetical order.
    {
        let project_bytes = build_project_stream_for_designer_modules(&["FormB", "FormA"]);
        let mut s = ole.create_stream("PROJECT").expect("PROJECT stream");
        s.write_all(&project_bytes).expect("write PROJECT");
    }

    ole.create_storage("VBA").expect("create VBA storage");
    {
        let dir_container = build_dir_stream(&[("FormA", "FormA"), ("FormB", "FormB")]);
        let mut s = ole.create_stream("VBA/dir").expect("dir stream");
        s.write_all(&dir_container).expect("write dir");
    }

    // Designer storages referenced by PROJECT stream.
    ole.create_storage("FormA").expect("create FormA storage");
    ole.create_storage("FormB").expect("create FormB storage");
    {
        let mut s = ole.create_stream("FormA/Data").expect("create FormA stream");
        s.write_all(b"A").expect("write A");
    }
    {
        let mut s = ole.create_stream("FormB/Data").expect("create FormB stream");
        s.write_all(b"B").expect("write B");
    }

    // An extra root-level storage that is *not* listed as a ProjectDesignerModule.
    ole.create_storage("Ignored").expect("create extra storage");
    {
        let mut s = ole
            .create_stream("Ignored/Data")
            .expect("create extra stream");
        s.write_all(b"Z").expect("write Z");
    }

    let vba_project_bin = ole.into_inner().into_inner();
    let normalized = forms_normalized_data(&vba_project_bin).expect("compute FormsNormalizedData");

    // PROJECT stream order: FormB then FormA.
    let mut expected = Vec::new();
    expected.extend_from_slice(b"B");
    expected.extend(std::iter::repeat(0u8).take(1022));
    expected.extend_from_slice(b"A");
    expected.extend(std::iter::repeat(0u8).take(1022));

    assert_eq!(normalized, expected);
}

#[test]
fn forms_normalized_data_empty_stream_contributes_no_bytes() {
    let cursor = Cursor::new(Vec::new());
    let mut ole = cfb::CompoundFile::create(cursor).expect("create compound file");

    {
        let project_bytes = build_project_stream_for_designer_modules(&["UserForm1"]);
        let mut s = ole.create_stream("PROJECT").expect("PROJECT stream");
        s.write_all(&project_bytes).expect("write PROJECT");
    }
    ole.create_storage("VBA").expect("create VBA storage");
    {
        let dir_container = build_dir_stream(&[("UserForm1", "UserForm1")]);
        let mut s = ole.create_stream("VBA/dir").expect("dir stream");
        s.write_all(&dir_container).expect("write dir");
    }

    ole.create_storage("UserForm1")
        .expect("create designer storage");
    // Empty stream.
    {
        let _s = ole
            .create_stream("UserForm1/Empty")
            .expect("create empty stream");
    }
    // Non-empty stream.
    {
        let mut s = ole
            .create_stream("UserForm1/Data")
            .expect("create data stream");
        s.write_all(b"A").expect("write A");
    }

    let vba_project_bin = ole.into_inner().into_inner();
    let normalized = forms_normalized_data(&vba_project_bin).expect("compute FormsNormalizedData");

    assert_eq!(
        normalized.len(),
        1023,
        "expected only the non-empty stream to contribute a single 1023-byte block"
    );
    assert_eq!(normalized[0], b'A');
}
