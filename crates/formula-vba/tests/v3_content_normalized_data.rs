use std::io::{Cursor, Write};

use formula_vba::{
    compress_container, content_normalized_data, forms_normalized_data, project_normalized_data_v3,
    v3_content_normalized_data,
};

fn push_record(out: &mut Vec<u8>, id: u16, data: &[u8]) {
    out.extend_from_slice(&id.to_le_bytes());
    out.extend_from_slice(&(data.len() as u32).to_le_bytes());
    out.extend_from_slice(data);
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

        // MODULETYPE = UserForm (0x0003 per MS-OVBA).
        push_record(&mut out, 0x0021, &0x0003u16.to_le_bytes());
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

#[test]
fn v3_content_normalized_data_includes_module_metadata_even_without_designers() {
    let vba_bin = build_project_no_designers();

    let v3 = v3_content_normalized_data(&vba_bin).expect("V3ContentNormalizedData");
    let content = content_normalized_data(&vba_bin).expect("ContentNormalizedData");
    assert_eq!(content, b"Sub Foo()\r\nEnd Sub\r\n".to_vec());

    // Per MS-OVBA v3, the module transcript includes:
    // MODULENAME || MODULESTREAMNAME(trimmed) || MODULETYPE || normalized_source
    let mut expected = Vec::new();
    expected.extend_from_slice(b"Module1");
    expected.extend_from_slice(b"Module1");
    expected.extend_from_slice(&0u16.to_le_bytes());
    expected.extend_from_slice(b"Sub Foo()\r\nEnd Sub\r\n");

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
    expected_content_v3.extend_from_slice(b"UserForm1");
    expected_content_v3.extend_from_slice(b"UserForm1");
    expected_content_v3.extend_from_slice(&0x0003u16.to_le_bytes());
    expected_content_v3.extend_from_slice(b"Sub FormHello()\r\nEnd Sub\r\n");
    assert_eq!(content_v3, expected_content_v3);

    let mut expected_forms = Vec::new();
    expected_forms.extend_from_slice(b"ABC");
    expected_forms.extend(std::iter::repeat(0u8).take(1020));
    assert_eq!(forms, expected_forms);

    let expected_project = [expected_content_v3.as_slice(), expected_forms.as_slice()].concat();
    assert_eq!(project, expected_project);
}
