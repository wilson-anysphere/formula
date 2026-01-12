use std::io::{Cursor, Write};

use formula_vba::{compress_container, v3_content_normalized_data};

fn push_record(out: &mut Vec<u8>, id: u16, data: &[u8]) {
    out.extend_from_slice(&id.to_le_bytes());
    out.extend_from_slice(&(data.len() as u32).to_le_bytes());
    out.extend_from_slice(data);
}

fn build_vba_project_with_dir(dir_decompressed: &[u8]) -> Vec<u8> {
    let dir_container = compress_container(dir_decompressed);

    let cursor = Cursor::new(Vec::new());
    let mut ole = cfb::CompoundFile::create(cursor).expect("create cfb");
    ole.create_storage("VBA").expect("VBA storage");
    {
        let mut s = ole.create_stream("VBA/dir").expect("dir stream");
        s.write_all(&dir_container).expect("write dir");
    }

    ole.into_inner().into_inner()
}

#[test]
fn v3_content_normalized_data_includes_all_reference_record_types_in_order() {
    // Build a decompressed `VBA/dir` stream containing each supported reference record type.
    //
    // MS-OVBA §2.4.2.5 V3ContentNormalizedData is a *field-level* transcript:
    // - it includes record IDs for reference records,
    // - it incorporates the specified payload fields (and some size-of-libid fields),
    // - it does *not* perform the legacy "copy bytes until first NUL" truncation used by
    //   ContentNormalizedData (v1).
    //
    // We deliberately include little-endian integers that contain embedded NUL bytes to ensure the
    // implementation preserves them.
    // REFERENCECONTROL (0x002F) payload: u32-len-prefixed libid + reserved1(u32) + reserved2(u16).
    //
    // reserved1=1 => 0x01 0x00 0x00 0x00 includes embedded NUL bytes that must be preserved.
    let libid_twiddled = b"ControlLib";
    let reserved1: u32 = 1;
    let reserved2: u16 = 0;
    let mut reference_control = Vec::new();
    reference_control.extend_from_slice(&(libid_twiddled.len() as u32).to_le_bytes());
    reference_control.extend_from_slice(libid_twiddled);
    reference_control.extend_from_slice(&reserved1.to_le_bytes());
    reference_control.extend_from_slice(&reserved2.to_le_bytes());

    // REFERENCEPROJECT (0x000E) payload: two u32-len-prefixed strings + major(u32) + minor(u16).
    //
    // major=1 => 0x01 0x00 0x00 0x00 includes embedded NUL bytes that must be preserved.
    let libid_absolute = b"ProjLib";
    let libid_relative = b"";
    let major: u32 = 1;
    let minor: u16 = 0;
    let mut reference_project = Vec::new();
    reference_project.extend_from_slice(&(libid_absolute.len() as u32).to_le_bytes());
    reference_project.extend_from_slice(libid_absolute);
    reference_project.extend_from_slice(&(libid_relative.len() as u32).to_le_bytes());
    reference_project.extend_from_slice(libid_relative);
    reference_project.extend_from_slice(&major.to_le_bytes());
    reference_project.extend_from_slice(&minor.to_le_bytes());

    let dir_decompressed = {
        let mut out = Vec::new();

        // REFERENCENAME (0x0016): Id + SizeOfName + Name.
        push_record(&mut out, 0x0016, b"RefName");
        // REFERENCENAMEUNICODE marker / payload (0x003E): copied as raw bytes.
        // Use UTF-16LE bytes that include NULs to ensure we do not treat this as NUL-terminated.
        push_record(&mut out, 0x003E, &[b'U', 0x00, b'N', 0x00]);

        // REFERENCECONTROL (0x002F)
        push_record(&mut out, 0x002F, &reference_control);

        // REFERENCEEXTENDED (0x0030): incorporated as raw payload bytes.
        push_record(&mut out, 0x0030, b"EXTENDED");

        // REFERENCEORIGINAL (0x0033): u32-len-prefixed libid.
        //
        // In spec-compliant dir streams, `SizeOfLibidOriginal` is the record header `len`; the
        // payload is the libid bytes (no u32 length prefix inside the payload).
        push_record(&mut out, 0x0033, b"OrigLib");

        // REFERENCEREGISTERED (0x000D): record bytes are incorporated directly.
        push_record(&mut out, 0x000D, b"{REG}");

        // REFERENCEPROJECT (0x000E)
        push_record(&mut out, 0x000E, &reference_project);

        out
    };

    let vba_bin = build_vba_project_with_dir(&dir_decompressed);
    let normalized = v3_content_normalized_data(&vba_bin).expect("V3ContentNormalizedData");

    // Expected output is the concatenation of reference transcripts in on-disk order.
    let mut expected = Vec::new();
    // REFERENCENAME: Id + SizeOfName + Name
    expected.extend_from_slice(&0x0016u16.to_le_bytes());
    expected.extend_from_slice(&(b"RefName".len() as u32).to_le_bytes());
    expected.extend_from_slice(b"RefName");
    // REFERENCENAMEUNICODE: Id + SizeOfNameUnicode + NameUnicode (UTF-16LE bytes).
    expected.extend_from_slice(&0x003Eu16.to_le_bytes());
    expected.extend_from_slice(&(4u32).to_le_bytes());
    expected.extend_from_slice(&[b'U', 0x00, b'N', 0x00]);
    // REFERENCECONTROL: Id + payload (no record SizeTwiddled field)
    expected.extend_from_slice(&0x002Fu16.to_le_bytes());
    expected.extend_from_slice(&reference_control);
    // REFERENCEEXTENDED: Id + payload (no record SizeExtended field)
    expected.extend_from_slice(&0x0030u16.to_le_bytes());
    expected.extend_from_slice(b"EXTENDED");
    // REFERENCEORIGINAL: Id + SizeOfLibidOriginal + LibidOriginal
    expected.extend_from_slice(&0x0033u16.to_le_bytes());
    expected.extend_from_slice(&(b"OrigLib".len() as u32).to_le_bytes());
    expected.extend_from_slice(b"OrigLib");
    // REFERENCEREGISTERED: Id + payload (no record Size field)
    expected.extend_from_slice(&0x000Du16.to_le_bytes());
    expected.extend_from_slice(b"{REG}");
    // REFERENCEPROJECT: Id + payload (no record Size field)
    expected.extend_from_slice(&0x000Eu16.to_le_bytes());
    expected.extend_from_slice(&reference_project);

    assert_eq!(normalized, expected);
}
