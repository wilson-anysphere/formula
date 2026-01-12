use std::io::{Cursor, Write};

use formula_vba::{compress_container, content_normalized_data};

fn push_record(out: &mut Vec<u8>, id: u16, data: &[u8]) {
    out.extend_from_slice(&id.to_le_bytes());
    out.extend_from_slice(&(data.len() as u32).to_le_bytes());
    out.extend_from_slice(data);
}

#[test]
fn content_normalized_data_includes_reference_record_allowlist_and_normalization() {
    // MS-OVBA §2.4.2.1 ContentNormalizedData incorporates only a subset of REFERENCE records.
    //
    // Included:
    // - 0x000D REFERENCEREGISTERED (raw bytes)
    // - 0x000E REFERENCEPROJECT (TempBuffer + copy-until-NUL)
    // - 0x002F REFERENCECONTROL (TempBuffer + copy-until-NUL)
    // - 0x0033 REFERENCEORIGINAL (u32-len libid + copy-until-NUL)
    // - 0x0030 REFERENCEEXTENDED (raw bytes)
    //
    // Excluded:
    // - 0x0016 REFERENCENAME (should not contribute)
    let dir_decompressed = {
        let mut out = Vec::new();

        // Excluded record: REFERENCENAME (0x0016). Should not affect output.
        push_record(&mut out, 0x0016, b"EXCLUDED_REF_NAME");

        // Included record: REFERENCEREGISTERED (0x000D).
        push_record(&mut out, 0x000D, b"{REG}");

        // Included record: REFERENCECONTROL (0x002F).
        // Structure used by the normalization pseudocode:
        // - u32 len + bytes (LibidTwiddled)
        // - reserved1 (u32)
        // - reserved2 (u16)
        //
        // Choose reserved values that contain an early NUL byte so copy-until-NUL stops quickly.
        let libid_twiddled = b"CtrlLib";
        let reserved1: u32 = 1; // 01 00 00 00
        let reserved2: u16 = 0;
        let mut reference_control = Vec::new();
        reference_control.extend_from_slice(&(libid_twiddled.len() as u32).to_le_bytes());
        reference_control.extend_from_slice(libid_twiddled);
        reference_control.extend_from_slice(&reserved1.to_le_bytes());
        reference_control.extend_from_slice(&reserved2.to_le_bytes());
        push_record(&mut out, 0x002F, &reference_control);

        // Included record: REFERENCEPROJECT (0x000E).
        // Structure used by the normalization pseudocode:
        // - u32 len + bytes (LibidAbsolute)
        // - u32 len + bytes (LibidRelative)
        // - major (u32)
        // - minor (u16)
        //
        // Choose major=1 so copy-until-NUL stops after copying 0x01.
        let libid_absolute = b"ProjLib";
        let libid_relative = b"";
        let major: u32 = 1; // 01 00 00 00
        let minor: u16 = 0;
        let mut reference_project = Vec::new();
        reference_project.extend_from_slice(&(libid_absolute.len() as u32).to_le_bytes());
        reference_project.extend_from_slice(libid_absolute);
        reference_project.extend_from_slice(&(libid_relative.len() as u32).to_le_bytes());
        reference_project.extend_from_slice(libid_relative);
        reference_project.extend_from_slice(&major.to_le_bytes());
        reference_project.extend_from_slice(&minor.to_le_bytes());
        push_record(&mut out, 0x000E, &reference_project);

        // Included record: REFERENCEORIGINAL (0x0033).
        // Structure used by the normalization pseudocode:
        // - u32 len + bytes (LibidOriginal)
        let libid_original = b"ORIG";
        let mut reference_original = Vec::new();
        reference_original.extend_from_slice(&(libid_original.len() as u32).to_le_bytes());
        reference_original.extend_from_slice(libid_original);
        push_record(&mut out, 0x0033, &reference_original);

        // Included record: REFERENCEEXTENDED (0x0030).
        push_record(&mut out, 0x0030, b"EXT");

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
    let vba_bin = ole.into_inner().into_inner();

    let normalized = content_normalized_data(&vba_bin).expect("ContentNormalizedData");

    let expected = [
        b"{REG}".as_slice(),
        b"CtrlLib\x01".as_slice(),
        b"ProjLib\x01".as_slice(),
        b"ORIG".as_slice(),
        b"EXT".as_slice(),
    ]
    .concat();

    assert_eq!(normalized, expected);
    assert!(
        !normalized.windows(b"EXCLUDED_REF_NAME".len()).any(|w| w == b"EXCLUDED_REF_NAME"),
        "REFERENCENAME (0x0016) must not contribute to ContentNormalizedData"
    );
}
