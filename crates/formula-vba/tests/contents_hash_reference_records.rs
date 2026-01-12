use std::io::{Cursor, Write};

use formula_vba::{compress_container, v3_content_normalized_data};

fn push_record(out: &mut Vec<u8>, id: u16, data: &[u8]) {
    out.extend_from_slice(&id.to_le_bytes());
    out.extend_from_slice(&(data.len() as u32).to_le_bytes());
    out.extend_from_slice(data);
}

fn push_record_with_size(out: &mut Vec<u8>, id: u16, size: u32, data: &[u8]) {
    out.extend_from_slice(&id.to_le_bytes());
    out.extend_from_slice(&size.to_le_bytes());
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

fn reference_registered_payload(libid: &[u8]) -> Vec<u8> {
    // REFERENCEREGISTERED (0x000D) payload:
    // SizeOfLibid (u32) + Libid (bytes) + Reserved1 (u32) + Reserved2 (u16).
    let mut out = Vec::new();
    out.extend_from_slice(&(libid.len() as u32).to_le_bytes());
    out.extend_from_slice(libid);
    out.extend_from_slice(&0u32.to_le_bytes());
    out.extend_from_slice(&0u16.to_le_bytes());
    out
}

fn reference_extended_payload(libid: &[u8]) -> Vec<u8> {
    // REFERENCEEXTENDED (0x0030) payload:
    // SizeOfLibidExtended (u32) + LibidExtended (bytes) + Reserved4 (u32) + Reserved5 (u16)
    // + OriginalTypeLib (GUID, 16 bytes) + Cookie (u32).
    let mut out = Vec::new();
    out.extend_from_slice(&(libid.len() as u32).to_le_bytes());
    out.extend_from_slice(libid);
    out.extend_from_slice(&0u32.to_le_bytes()); // Reserved4
    out.extend_from_slice(&0u16.to_le_bytes()); // Reserved5
    out.extend_from_slice(&[0u8; 16]); // OriginalTypeLib
    out.extend_from_slice(&0u32.to_le_bytes()); // Cookie
    out
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

    let reference_extended = reference_extended_payload(b"EXTENDED");
    let reference_registered = reference_registered_payload(b"{REG}");

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
        push_record(&mut out, 0x0030, &reference_extended);

        // REFERENCEORIGINAL (0x0033): u32-len-prefixed libid.
        //
        // In spec-compliant dir streams, `SizeOfLibidOriginal` is the record header `len`; the
        // payload is the libid bytes (no u32 length prefix inside the payload).
        push_record(&mut out, 0x0033, b"OrigLib");

        // REFERENCEREGISTERED (0x000D): payload bytes are incorporated directly (excluding the record
        // size field).
        push_record(&mut out, 0x000D, &reference_registered);

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
    expected.extend_from_slice(&reference_extended);
    // REFERENCEORIGINAL: Id + SizeOfLibidOriginal + LibidOriginal
    expected.extend_from_slice(&0x0033u16.to_le_bytes());
    expected.extend_from_slice(&(b"OrigLib".len() as u32).to_le_bytes());
    expected.extend_from_slice(b"OrigLib");
    // REFERENCEREGISTERED: Id + payload (no record Size field)
    expected.extend_from_slice(&0x000Du16.to_le_bytes());
    expected.extend_from_slice(&reference_registered);
    // REFERENCEPROJECT: Id + payload (no record Size field)
    expected.extend_from_slice(&0x000Eu16.to_le_bytes());
    expected.extend_from_slice(&reference_project);

    assert_eq!(normalized, expected);
}

#[test]
fn v3_content_normalized_data_skips_referenceoriginal_embedded_referencecontrol() {
    // MS-OVBA `REFERENCEORIGINAL` embeds an immediate `REFERENCECONTROL` record; that embedded
    // control record must not contribute to the v3 transcript.
    let dir_decompressed = {
        let mut out = Vec::new();

        // NameRecord for the reference.
        push_record(&mut out, 0x0016, b"OrigName");
        push_record(&mut out, 0x003E, &[b'O', 0x00, b'K', 0x00]);

        // REFERENCEORIGINAL payload is just the libid bytes (spec form).
        push_record(&mut out, 0x0033, b"OrigLib");

        // Embedded REFERENCECONTROL (should be skipped).
        let libid_twiddled = b"EmbeddedCtrl";
        let reserved1: u32 = 1;
        let reserved2: u16 = 0;
        let mut embedded_control = Vec::new();
        embedded_control.extend_from_slice(&(libid_twiddled.len() as u32).to_le_bytes());
        embedded_control.extend_from_slice(libid_twiddled);
        embedded_control.extend_from_slice(&reserved1.to_le_bytes());
        embedded_control.extend_from_slice(&reserved2.to_le_bytes());
        push_record(&mut out, 0x002F, &embedded_control);

        // Embedded NameRecordExtended (should also be skipped).
        push_record(&mut out, 0x0016, b"ExtName");
        push_record(&mut out, 0x003E, &[b'E', 0x00]);

        // Embedded control tail / extended bytes (should be skipped).
        let embedded_extended = reference_extended_payload(b"EMB_EXT");
        push_record(&mut out, 0x0030, &embedded_extended);

        // Next reference (registered) should still be incorporated.
        let reference_registered = reference_registered_payload(b"{REG}");
        push_record(&mut out, 0x000D, &reference_registered);

        out
    };

    let vba_bin = build_vba_project_with_dir(&dir_decompressed);
    let normalized = v3_content_normalized_data(&vba_bin).expect("V3ContentNormalizedData");

    let mut expected = Vec::new();
    // REFERENCENAME: Id + SizeOfName + Name
    expected.extend_from_slice(&0x0016u16.to_le_bytes());
    expected.extend_from_slice(&(b"OrigName".len() as u32).to_le_bytes());
    expected.extend_from_slice(b"OrigName");
    // REFERENCENAMEUNICODE: Id + SizeOfNameUnicode + NameUnicode
    expected.extend_from_slice(&0x003Eu16.to_le_bytes());
    expected.extend_from_slice(&(4u32).to_le_bytes());
    expected.extend_from_slice(&[b'O', 0x00, b'K', 0x00]);
    // REFERENCEORIGINAL: Id + SizeOfLibidOriginal + LibidOriginal
    expected.extend_from_slice(&0x0033u16.to_le_bytes());
    expected.extend_from_slice(&(b"OrigLib".len() as u32).to_le_bytes());
    expected.extend_from_slice(b"OrigLib");
    // Embedded REFERENCECONTROL (+ name/extended tail) is skipped.
    // Next reference (registered): Id + payload
    expected.extend_from_slice(&0x000Du16.to_le_bytes());
    expected.extend_from_slice(&reference_registered_payload(b"{REG}"));

    assert_eq!(normalized, expected);
}

#[test]
fn v3_content_normalized_data_skips_referenceoriginal_embedded_referencecontrol_without_name_record_extended(
) {
    // Variant of the previous regression test where the embedded REFERENCECONTROL record omits the
    // optional NameRecordExtended (REFERENCENAME). We should still skip the embedded control record
    // and its 0x0030 tail, and continue parsing subsequent top-level references correctly.
    let dir_decompressed = {
        let mut out = Vec::new();

        // NameRecord for the reference.
        push_record(&mut out, 0x0016, b"OrigName");
        push_record(&mut out, 0x003E, &[b'O', 0x00, b'K', 0x00]);

        // REFERENCEORIGINAL payload is just the libid bytes (spec form).
        push_record(&mut out, 0x0033, b"OrigLib");

        // Embedded REFERENCECONTROL (should be skipped).
        let libid_twiddled = b"EmbeddedCtrl";
        let reserved1: u32 = 1;
        let reserved2: u16 = 0;
        let mut embedded_control = Vec::new();
        embedded_control.extend_from_slice(&(libid_twiddled.len() as u32).to_le_bytes());
        embedded_control.extend_from_slice(libid_twiddled);
        embedded_control.extend_from_slice(&reserved1.to_le_bytes());
        embedded_control.extend_from_slice(&reserved2.to_le_bytes());
        push_record(&mut out, 0x002F, &embedded_control);

        // Embedded control tail / extended bytes (should be skipped).
        let embedded_extended = reference_extended_payload(b"EMB_EXT");
        push_record(&mut out, 0x0030, &embedded_extended);

        // Next reference (registered) should still be incorporated.
        let reference_registered = reference_registered_payload(b"{REG}");
        push_record(&mut out, 0x000D, &reference_registered);

        out
    };

    let vba_bin = build_vba_project_with_dir(&dir_decompressed);
    let normalized = v3_content_normalized_data(&vba_bin).expect("V3ContentNormalizedData");

    let mut expected = Vec::new();
    // REFERENCENAME
    expected.extend_from_slice(&0x0016u16.to_le_bytes());
    expected.extend_from_slice(&(b"OrigName".len() as u32).to_le_bytes());
    expected.extend_from_slice(b"OrigName");
    // NameUnicode
    expected.extend_from_slice(&0x003Eu16.to_le_bytes());
    expected.extend_from_slice(&(4u32).to_le_bytes());
    expected.extend_from_slice(&[b'O', 0x00, b'K', 0x00]);
    // REFERENCEORIGINAL
    expected.extend_from_slice(&0x0033u16.to_le_bytes());
    expected.extend_from_slice(&(b"OrigLib".len() as u32).to_le_bytes());
    expected.extend_from_slice(b"OrigLib");
    // Embedded control is skipped.
    // Next reference (registered)
    expected.extend_from_slice(&0x000Du16.to_le_bytes());
    expected.extend_from_slice(&reference_registered_payload(b"{REG}"));

    assert_eq!(normalized, expected);
}

#[test]
fn v3_content_normalized_data_skips_referenceoriginal_embedded_referencecontrol_with_malformed_record_size_fields(
) {
    // Like `..._without_name_record_extended`, but uses incorrect record-size values (0) for the
    // embedded REFERENCECONTROL (0x002F) and REFERENCEEXTENDED (0x0030) records. These size fields
    // MUST be ignored by MS-OVBA v3 transcript construction, and should not be used for record
    // framing.
    let dir_decompressed = {
        let mut out = Vec::new();

        // NameRecord for the reference.
        push_record(&mut out, 0x0016, b"OrigName");
        push_record(&mut out, 0x003E, &[b'O', 0x00, b'K', 0x00]);

        // REFERENCEORIGINAL payload is just the libid bytes (spec form).
        push_record(&mut out, 0x0033, b"OrigLib");

        // Embedded REFERENCECONTROL (should be skipped); malformed SizeTwiddled.
        let libid_twiddled = b"EmbeddedCtrl";
        let reserved1: u32 = 1;
        let reserved2: u16 = 0;
        let mut embedded_control = Vec::new();
        embedded_control.extend_from_slice(&(libid_twiddled.len() as u32).to_le_bytes());
        embedded_control.extend_from_slice(libid_twiddled);
        embedded_control.extend_from_slice(&reserved1.to_le_bytes());
        embedded_control.extend_from_slice(&reserved2.to_le_bytes());
        push_record_with_size(&mut out, 0x002F, 0, &embedded_control);

        // Embedded control tail / extended bytes (should be skipped); malformed SizeExtended.
        let embedded_extended = reference_extended_payload(b"EMB_EXT");
        push_record_with_size(&mut out, 0x0030, 0, &embedded_extended);

        // Next reference (registered) should still be incorporated.
        let reference_registered = reference_registered_payload(b"{REG}");
        push_record(&mut out, 0x000D, &reference_registered);

        out
    };

    let vba_bin = build_vba_project_with_dir(&dir_decompressed);
    let normalized = v3_content_normalized_data(&vba_bin).expect("V3ContentNormalizedData");

    let mut expected = Vec::new();
    expected.extend_from_slice(&0x0016u16.to_le_bytes());
    expected.extend_from_slice(&(b"OrigName".len() as u32).to_le_bytes());
    expected.extend_from_slice(b"OrigName");
    expected.extend_from_slice(&0x003Eu16.to_le_bytes());
    expected.extend_from_slice(&(4u32).to_le_bytes());
    expected.extend_from_slice(&[b'O', 0x00, b'K', 0x00]);
    expected.extend_from_slice(&0x0033u16.to_le_bytes());
    expected.extend_from_slice(&(b"OrigLib".len() as u32).to_le_bytes());
    expected.extend_from_slice(b"OrigLib");
    // Embedded control records are skipped.
    expected.extend_from_slice(&0x000Du16.to_le_bytes());
    expected.extend_from_slice(&reference_registered_payload(b"{REG}"));

    assert_eq!(normalized, expected);
}
