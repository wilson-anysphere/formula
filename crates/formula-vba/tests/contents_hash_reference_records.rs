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
    // We deliberately choose numeric little-endian fields that contain an early NUL (0x00) to
    // exercise the MS-OVBA "copy bytes until first NUL" normalization behavior.
    let dir_decompressed = {
        let mut out = Vec::new();

        // REFERENCENAME (0x0016): copied as raw bytes.
        push_record(&mut out, 0x0016, b"RefName");
        // REFERENCENAMEUNICODE marker / payload (0x003E): copied as raw bytes.
        // Use UTF-16LE bytes that include NULs to ensure we do not treat this as NUL-terminated.
        push_record(&mut out, 0x003E, &[b'U', 0x00, b'N', 0x00]);

        // REFERENCECONTROL (0x002F): u32-len-prefixed libid + reserved1(u32) + reserved2(u16).
        //
        // reserved1=1 => 0x01 0x00 0x00 0x00 causes the normalization to stop immediately after the
        // low byte (0x01) of reserved1.
        let libid_twiddled = b"ControlLib";
        let reserved1: u32 = 1;
        let reserved2: u16 = 0;
        let mut reference_control = Vec::new();
        reference_control.extend_from_slice(&(libid_twiddled.len() as u32).to_le_bytes());
        reference_control.extend_from_slice(libid_twiddled);
        reference_control.extend_from_slice(&reserved1.to_le_bytes());
        reference_control.extend_from_slice(&reserved2.to_le_bytes());
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

        // REFERENCEPROJECT (0x000E): two u32-len-prefixed strings + major(u32) + minor(u16).
        //
        // major=1 => 0x01 0x00 0x00 0x00 causes early termination (after copying 0x01).
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
        push_record(&mut out, 0x000E, &reference_project);

        out
    };

    let vba_bin = build_vba_project_with_dir(&dir_decompressed);
    let normalized = v3_content_normalized_data(&vba_bin).expect("V3ContentNormalizedData");

    // Expected output is the concatenation of normalized reference record bytes in on-disk order:
    // - 0x0016: raw payload bytes
    // - 0x003E: raw payload bytes
    // - 0x002F: TempBuffer = LibidTwiddled || Reserved1 || Reserved2; copy until first 0x00
    // - 0x0030: raw payload bytes
    // - 0x0033: LibidOriginal; copy until first 0x00
    // - 0x000D: raw bytes
    // - 0x000E: TempBuffer = LibidAbsolute || LibidRelative || MajorVersion || MinorVersion; copy until first 0x00
    let expected_name = b"RefName".as_slice();
    let expected_name_unicode = [b'U', 0x00, b'N', 0x00].as_slice();
    let expected_control = b"ControlLib\x01".as_slice();
    let expected_extended = b"EXTENDED".as_slice();
    let expected_original = b"OrigLib".as_slice();
    let expected_registered = b"{REG}".as_slice();
    let expected_project = b"ProjLib\x01".as_slice();
    let expected = [
        expected_name,
        expected_name_unicode,
        expected_control,
        expected_extended,
        expected_original,
        expected_registered,
        expected_project,
    ]
    .concat();

    assert_eq!(normalized, expected);
}
