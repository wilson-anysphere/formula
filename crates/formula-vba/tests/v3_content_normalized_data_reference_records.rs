use std::io::{Cursor, Write};

use formula_vba::{compress_container, v3_content_normalized_data};

// MS-OVBA §2.4.2.5 V3ContentNormalizedData reference handling regression tests.
//
// These tests are intentionally byte-level and aim to catch subtle transcript regressions:
// - accidentally including record size fields (e.g. SizeTwiddled / SizeExtended / Size),
// - omitting REFERENCENAME unicode name bytes, and
// - omitting the optional NameRecordExtended (REFERENCENAME) inside REFERENCECONTROL.

fn push_record(out: &mut Vec<u8>, id: u16, data: &[u8]) {
    out.extend_from_slice(&id.to_le_bytes());
    out.extend_from_slice(&(data.len() as u32).to_le_bytes());
    out.extend_from_slice(data);
}

fn push_reference_name(dir: &mut Vec<u8>, expected: &mut Vec<u8>, name: &str) {
    let name_bytes = name.as_bytes();
    let name_unicode: Vec<u8> = name
        .encode_utf16()
        .flat_map(|u| u.to_le_bytes())
        .collect();

    // REFERENCENAME (0x0016): Id + SizeOfName + Name
    push_record(dir, 0x0016, name_bytes);

    expected.extend_from_slice(&0x0016u16.to_le_bytes());
    expected.extend_from_slice(&(name_bytes.len() as u32).to_le_bytes());
    expected.extend_from_slice(name_bytes);

    // The "NameUnicode" suffix is encoded as the 0x003E marker + u32 size + UTF-16LE bytes.
    // The v3 transcript incorporates these bytes and should not omit them.
    push_record(dir, 0x003E, &name_unicode);
    expected.extend_from_slice(&0x003Eu16.to_le_bytes());
    expected.extend_from_slice(&(name_unicode.len() as u32).to_le_bytes());
    expected.extend_from_slice(&name_unicode);
}

#[test]
fn v3_content_normalized_data_reference_handling_includes_names_and_excludes_record_size_fields() {
    let mut dir_decompressed = Vec::new();
    let mut expected = Vec::new();

    // ---- Reference 1: REFERENCECONTROL (0x002F) ----
    push_reference_name(&mut dir_decompressed, &mut expected, "ControlRef");

    // REFERENCECONTROL record payload (after the record-size field / SizeTwiddled):
    // SizeOfLibidTwiddled (u32) + LibidTwiddled + Reserved1 (u32) + Reserved2 (u16).
    let libid_twiddled = b"LIBID-TWIDDLED";
    let reserved1: u32 = 0x00000000;
    let reserved2: u16 = 0x0000;
    let mut reference_control = Vec::new();
    reference_control.extend_from_slice(&(libid_twiddled.len() as u32).to_le_bytes());
    reference_control.extend_from_slice(libid_twiddled);
    reference_control.extend_from_slice(&reserved1.to_le_bytes());
    reference_control.extend_from_slice(&reserved2.to_le_bytes());
    push_record(&mut dir_decompressed, 0x002F, &reference_control);

    // V3 transcript: include Id + payload, but NOT SizeTwiddled.
    expected.extend_from_slice(&0x002Fu16.to_le_bytes());
    expected.extend_from_slice(&reference_control);

    // NameRecordExtended (optional REFERENCENAME inside REFERENCECONTROL) including NameUnicode.
    push_reference_name(&mut dir_decompressed, &mut expected, "ExtTypeLib");

    // REFERENCEEXTENDED record (0x0030) that pairs with the control reference.
    let libid_extended = b"LIBID-EXTENDED";
    let reserved4: u32 = 0x00000000;
    let reserved5: u16 = 0x0000;
    let original_type_lib: [u8; 16] = [
        0x00, 0x11, 0x22, 0x33, 0x44, 0x55, 0x66, 0x77, 0x88, 0x99, 0xAA, 0xBB, 0xCC, 0xDD,
        0xEE, 0xFF,
    ];
    let cookie: u32 = 0xA1B2C3D4;
    let mut reference_extended = Vec::new();
    reference_extended.extend_from_slice(&(libid_extended.len() as u32).to_le_bytes());
    reference_extended.extend_from_slice(libid_extended);
    reference_extended.extend_from_slice(&reserved4.to_le_bytes());
    reference_extended.extend_from_slice(&reserved5.to_le_bytes());
    reference_extended.extend_from_slice(&original_type_lib);
    reference_extended.extend_from_slice(&cookie.to_le_bytes());
    push_record(&mut dir_decompressed, 0x0030, &reference_extended);

    // V3 transcript: include 0x0030 (Reserved3) + payload, but NOT SizeExtended.
    expected.extend_from_slice(&0x0030u16.to_le_bytes());
    expected.extend_from_slice(&reference_extended);

    // ---- Reference 2: REFERENCEORIGINAL (0x0033) ----
    push_reference_name(&mut dir_decompressed, &mut expected, "OriginalRef");

    let libid_original = b"LIBID-ORIGINAL";
    push_record(&mut dir_decompressed, 0x0033, libid_original);

    // V3 transcript: include Id + SizeOfLibidOriginal + LibidOriginal (and NOT any record Size field).
    expected.extend_from_slice(&0x0033u16.to_le_bytes());
    expected.extend_from_slice(&(libid_original.len() as u32).to_le_bytes());
    expected.extend_from_slice(libid_original);

    // ---- Reference 3: REFERENCEREGISTERED (0x000D) ----
    push_reference_name(&mut dir_decompressed, &mut expected, "RegisteredRef");

    let libid_registered = b"LIBID-REGISTERED";
    let mut reference_registered = Vec::new();
    reference_registered.extend_from_slice(&(libid_registered.len() as u32).to_le_bytes());
    reference_registered.extend_from_slice(libid_registered);
    reference_registered.extend_from_slice(&0u32.to_le_bytes()); // Reserved1
    reference_registered.extend_from_slice(&0u16.to_le_bytes()); // Reserved2
    push_record(&mut dir_decompressed, 0x000D, &reference_registered);

    // V3 transcript: include Id + payload, but NOT the record Size field.
    expected.extend_from_slice(&0x000Du16.to_le_bytes());
    expected.extend_from_slice(&reference_registered);

    // ---- Reference 4: REFERENCEPROJECT (0x000E) ----
    push_reference_name(&mut dir_decompressed, &mut expected, "ProjectRef");

    let libid_absolute = b"ABS-PATH";
    let libid_relative = b"REL-PATH";
    let major: u32 = 0x01020304;
    let minor: u16 = 0x0506;
    let mut reference_project = Vec::new();
    reference_project.extend_from_slice(&(libid_absolute.len() as u32).to_le_bytes());
    reference_project.extend_from_slice(libid_absolute);
    reference_project.extend_from_slice(&(libid_relative.len() as u32).to_le_bytes());
    reference_project.extend_from_slice(libid_relative);
    reference_project.extend_from_slice(&major.to_le_bytes());
    reference_project.extend_from_slice(&minor.to_le_bytes());
    push_record(&mut dir_decompressed, 0x000E, &reference_project);

    // V3 transcript: include Id + payload, but NOT the record Size field.
    expected.extend_from_slice(&0x000Eu16.to_le_bytes());
    expected.extend_from_slice(&reference_project);

    // ---- Build OLE with the `VBA/dir` stream ----
    let dir_container = compress_container(&dir_decompressed);
    let cursor = Cursor::new(Vec::new());
    let mut ole = cfb::CompoundFile::create(cursor).expect("create cfb");
    ole.create_storage("VBA").expect("VBA storage");
    {
        let mut s = ole.create_stream("VBA/dir").expect("dir stream");
        s.write_all(&dir_container).expect("write dir");
    }
    let vba_bin = ole.into_inner().into_inner();

    let actual = v3_content_normalized_data(&vba_bin).expect("V3ContentNormalizedData");
    assert_eq!(actual, expected);
}

