use std::io::{Cursor, Write};

use formula_vba::{
    compress_container, forms_normalized_data, project_normalized_data_v3,
    v3_content_normalized_data,
};

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
        push_record(&mut out, 0x002F, b"REFCTRL-V3");

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
fn project_normalized_data_v3_is_v3_content_plus_forms_normalized_data() {
    // Minimal VBA project with:
    // - one module
    // - one designer stream, so FormsNormalizedData is non-empty
    let module_code = b"Sub Hello()\r\nEnd Sub\r\n";
    let module_container = compress_container(module_code);
    let userform_code = b"Sub FormHello()\r\nEnd Sub\r\n";
    let userform_container = compress_container(userform_code);

    let dir_decompressed = {
        let mut out = Vec::new();

        // Standard module.
        push_record(&mut out, 0x0019, b"Module1");
        let mut stream_name = Vec::new();
        stream_name.extend_from_slice(b"Module1");
        stream_name.extend_from_slice(&0u16.to_le_bytes());
        push_record(&mut out, 0x001A, &stream_name);
        push_record(&mut out, 0x0021, &0u16.to_le_bytes());
        push_record(&mut out, 0x0031, &0u32.to_le_bytes());

        // UserForm (designer) module. The PROJECT stream references this by `BaseClass=`.
        push_record(&mut out, 0x0019, b"UserForm1");
        let mut stream_name = Vec::new();
        stream_name.extend_from_slice(b"UserForm1");
        stream_name.extend_from_slice(&0u16.to_le_bytes());
        push_record(&mut out, 0x001A, &stream_name);
        // MODULETYPE = UserForm (0x0003 per MS-OVBA).
        push_record(&mut out, 0x0021, &0x0003u16.to_le_bytes());
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
        s.write_all(b"Name=\"VBAProject\"\r\nBaseClass=\"UserForm1\"\r\n")
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
    {
        let mut s = ole
            .create_stream("VBA/UserForm1")
            .expect("userform module stream");
        s.write_all(&userform_container)
            .expect("write userform module stream");
    }
    {
        let mut s = ole
            .create_stream("UserForm1/Payload")
            .expect("designer stream");
        s.write_all(b"ABC").expect("write designer payload");
    }

    let vba_bin = ole.into_inner().into_inner();

    let content = v3_content_normalized_data(&vba_bin).expect("V3ContentNormalizedData");
    let forms = forms_normalized_data(&vba_bin).expect("FormsNormalizedData");
    assert!(
        !forms.is_empty(),
        "expected non-empty FormsNormalizedData when PROJECT contains BaseClass= and designer storage is present"
    );
    let project = project_normalized_data_v3(&vba_bin).expect("ProjectNormalizedData v3");

    let mut expected = Vec::new();
    expected.extend_from_slice(&content);
    expected.extend_from_slice(&forms);
    assert_eq!(project, expected);
}
