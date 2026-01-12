use std::io::{Cursor, Write};

use formula_vba::{compress_container, v3_content_normalized_data};

fn push_record(out: &mut Vec<u8>, id: u16, data: &[u8]) {
    out.extend_from_slice(&id.to_le_bytes());
    out.extend_from_slice(&(data.len() as u32).to_le_bytes());
    out.extend_from_slice(data);
}

#[test]
fn v3_content_normalized_data_accepts_tlv_projectversion_record() {
    // Some producers encode PROJECTVERSION (0x0009) as a normal TLV record:
    //   Id(u16) || Size(u32) || Data(Size)
    //
    // MS-OVBA defines PROJECTVERSION as fixed-length (no Size field):
    //   Id(u16) || Reserved(u32) || VersionMajor(u32) || VersionMinor(u16)
    //
    // V3ContentNormalizedData should normalize both encodings to the same transcript bytes:
    // exclude the TLV Size field and include the fixed-length payload fields.
    let module_code = b"Sub Foo()\r\nEnd Sub\r\n";
    let module_container = compress_container(module_code);

    let reserved: u32 = 0x0000_0004;
    let version_major: u32 = 0x1122_3344;
    let version_minor: u16 = 0x5566;

    let dir_decompressed = {
        let mut out = Vec::new();

        // PROJECTVERSION encoded as TLV: Size=10, Data=Reserved||Major||Minor.
        let mut projectversion_payload = Vec::new();
        projectversion_payload.extend_from_slice(&reserved.to_le_bytes());
        projectversion_payload.extend_from_slice(&version_major.to_le_bytes());
        projectversion_payload.extend_from_slice(&version_minor.to_le_bytes());
        push_record(&mut out, 0x0009, &projectversion_payload);

        // Minimal module record group.
        push_record(&mut out, 0x0019, b"Module1"); // MODULENAME
        let mut stream_name = Vec::new();
        stream_name.extend_from_slice(b"Module1");
        stream_name.extend_from_slice(&0u16.to_le_bytes());
        push_record(&mut out, 0x001A, &stream_name); // MODULESTREAMNAME
        push_record(&mut out, 0x0021, &0u16.to_le_bytes()); // MODULETYPE (procedural)
        push_record(&mut out, 0x0031, &0u32.to_le_bytes()); // MODULETEXTOFFSET (0)

        // Dir stream terminator.
        push_record(&mut out, 0x0010, &[]);

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
            .create_stream("VBA/Module1")
            .expect("module stream");
        s.write_all(&module_container).expect("write module");
    }
    let vba_bin = ole.into_inner().into_inner();

    let normalized = v3_content_normalized_data(&vba_bin).expect("V3ContentNormalizedData");

    let mut expected = Vec::new();
    // PROJECTVERSION normalized bytes (no TLV size field).
    expected.extend_from_slice(&0x0009u16.to_le_bytes());
    expected.extend_from_slice(&reserved.to_le_bytes());
    expected.extend_from_slice(&version_major.to_le_bytes());
    expected.extend_from_slice(&version_minor.to_le_bytes());
    // Procedural module TypeRecord: Id || Reserved(u16).
    expected.extend_from_slice(&0x0021u16.to_le_bytes());
    expected.extend_from_slice(&0u16.to_le_bytes());
    // Normalized module source (LF-only) + module name + LF (HashModuleNameFlag=true).
    expected.extend_from_slice(b"Sub Foo()\nEnd Sub\n\nModule1\n");
    // Dir terminator trailer: Id || Size(0).
    expected.extend_from_slice(&0x0010u16.to_le_bytes());
    expected.extend_from_slice(&0u32.to_le_bytes());

    assert_eq!(normalized, expected);
}

