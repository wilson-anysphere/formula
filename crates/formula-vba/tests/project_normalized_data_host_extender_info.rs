use std::io::{Cursor, Write};

use formula_vba::{compress_container, project_normalized_data};

fn contains_subslice(haystack: &[u8], needle: &[u8]) -> bool {
    if needle.is_empty() {
        return true;
    }
    haystack.windows(needle.len()).any(|w| w == needle)
}

#[test]
fn project_normalized_data_includes_host_extender_info_and_strips_nwln_and_excludes_project_id() {
    // MS-OVBA §2.4.2.6 requires:
    // - excluding ProjectId ("ID=") from ProjectProperties output, and
    // - if "[Host Extender Info]" exists, appending:
    //   - the literal bytes "Host Extender Info" (no brackets), and
    //   - HostExtenderRef bytes with NWLN removed.

    let excluded_project_id_guid = "11111111-2222-3333-4444-555555555555";

    let host_extender_ref_1 = "HostExtenderRef=&H00000001={3832D640-CF90-11CF-8E43-00A0C911005A};VBE;&H00000000";
    let host_extender_ref_2 = "HostExtenderRef=&H00000002={3832D641-CF90-11CF-8E43-00A0C911005A};VBE;&H00000000";

    let project_stream = format!(
        concat!(
            "ID=\"{{{excluded_project_id_guid}}}\"\r\n",
            "Name=\"VBAProject\"\r\n",
            "\r\n",
            "[Host Extender Info]\r\n",
            "{host_extender_ref_1}\r\n",
            "{host_extender_ref_2}\r\n",
            "\r\n",
            "[Workspace]\r\n",
            "ThisWorkbook=0, 0, 0, 0\r\n",
        ),
        excluded_project_id_guid = excluded_project_id_guid,
        host_extender_ref_1 = host_extender_ref_1,
        host_extender_ref_2 = host_extender_ref_2,
    );

    let cursor = Cursor::new(Vec::new());
    let mut ole = cfb::CompoundFile::create(cursor).expect("create cfb");

    {
        let mut s = ole.create_stream("PROJECT").expect("PROJECT stream");
        s.write_all(project_stream.as_bytes())
            .expect("write PROJECT stream");
    }

    // Include a minimal valid `VBA/dir` stream so the normalization implementation can load it if
    // it needs to (even though this test doesn't include any `BaseClass=` designer modules).
    ole.create_storage("VBA").expect("VBA storage");
    {
        let dir_container = compress_container(&[]);
        let mut s = ole.create_stream("VBA/dir").expect("dir stream");
        s.write_all(&dir_container).expect("write dir stream");
    }

    let vba_project_bin = ole.into_inner().into_inner();

    let normalized =
        project_normalized_data(&vba_project_bin).expect("compute ProjectNormalizedData");

    // Property filtering: ProjectId must not be included in the transcript.
    assert!(
        !contains_subslice(&normalized, excluded_project_id_guid.as_bytes()),
        "expected ProjectId GUID bytes to be excluded from ProjectNormalizedData"
    );

    // Host Extender Info header: output MUST include `Host Extender Info` without brackets.
    assert!(
        contains_subslice(&normalized, b"Host Extender Info"),
        "expected ProjectNormalizedData to include Host Extender Info header bytes"
    );
    assert!(
        !contains_subslice(&normalized, b"[Host Extender Info]"),
        "expected ProjectNormalizedData to NOT include bracketed section header bytes"
    );

    // HostExtenderRef bytes MUST have NWLN removed, so the two lines are concatenated.
    let expected_host_extender_bytes =
        format!("Host Extender Info{host_extender_ref_1}{host_extender_ref_2}");
    assert!(
        contains_subslice(&normalized, expected_host_extender_bytes.as_bytes()),
        "expected HostExtenderRef lines to be concatenated without newlines"
    );

    // Regression guard: the original (newline-separated) representation MUST NOT appear.
    let not_expected =
        format!("Host Extender Info{host_extender_ref_1}\r\n{host_extender_ref_2}");
    assert!(
        !contains_subslice(&normalized, not_expected.as_bytes()),
        "expected NWLN to be removed from HostExtenderRef"
    );
}

#[test]
fn project_normalized_data_host_extender_info_header_and_key_matching_is_case_insensitive() {
    // Some producers may vary the casing of `[Host Extender Info]` and `HostExtenderRef=...`.
    // For robustness we treat these case-insensitively, but must still preserve the raw line
    // bytes in the transcript (only NWLN removed).

    let host_extender_ref = "HOSTEXTENDERREF=MyHostExtender";

    let project_stream = format!(
        concat!(
            "ID=\"{{11111111-2222-3333-4444-555555555555}}\"\r\n",
            "Name=\"VBAProject\"\r\n",
            "\r\n",
            "[HOST EXTENDER INFO]\r\n",
            "{host_extender_ref}\r\n",
            "\r\n",
            "[Workspace]\r\n",
            "ThisWorkbook=SHOULD_NOT_APPEAR_IN_HASH\r\n",
        ),
        host_extender_ref = host_extender_ref,
    );

    let cursor = Cursor::new(Vec::new());
    let mut ole = cfb::CompoundFile::create(cursor).expect("create cfb");

    {
        let mut s = ole.create_stream("PROJECT").expect("PROJECT stream");
        s.write_all(project_stream.as_bytes())
            .expect("write PROJECT stream");
    }

    // Include a minimal `VBA/dir` stream so `project_normalized_data()` can load it (contents are
    // irrelevant for this test).
    ole.create_storage("VBA").expect("VBA storage");
    {
        let dir_container = compress_container(&[]);
        let mut s = ole.create_stream("VBA/dir").expect("dir stream");
        s.write_all(&dir_container).expect("write dir stream");
    }

    let vba_project_bin = ole.into_inner().into_inner();
    let normalized =
        project_normalized_data(&vba_project_bin).expect("compute ProjectNormalizedData");

    // Host Extender Info header is always emitted as the literal bytes `Host Extender Info`.
    assert!(
        contains_subslice(&normalized, b"Host Extender Info"),
        "expected ProjectNormalizedData to include Host Extender Info header bytes"
    );

    // The HostExtenderRef line bytes must be present (with original casing preserved).
    assert!(
        contains_subslice(&normalized, host_extender_ref.as_bytes()),
        "expected ProjectNormalizedData to include the HostExtenderRef line bytes"
    );

    // Newlines must be removed.
    assert!(
        !contains_subslice(
            &normalized,
            format!("{host_extender_ref}\r\n").as_bytes()
        ),
        "expected HostExtenderRef bytes to be appended without NWLN"
    );

    // Workspace section must be ignored.
    assert!(
        !contains_subslice(&normalized, b"ThisWorkbook=SHOULD_NOT_APPEAR_IN_HASH"),
        "expected [Workspace] section lines to be ignored"
    );
}

#[test]
fn project_normalized_data_host_extender_info_strips_utf8_bom() {
    // Some producers may include a UTF-8 BOM at the start of the PROJECT stream. Ensure this does
    // not prevent `[Host Extender Info]` section detection.

    let project_stream = b"\xEF\xBB\xBF[Host Extender Info]\r\nHostExtenderRef=MyHostExtender\r\n";

    let cursor = Cursor::new(Vec::new());
    let mut ole = cfb::CompoundFile::create(cursor).expect("create cfb");

    {
        let mut s = ole.create_stream("PROJECT").expect("PROJECT stream");
        s.write_all(project_stream).expect("write PROJECT stream");
    }

    // `project_normalized_data()` requires a `VBA/dir` stream to exist.
    ole.create_storage("VBA").expect("VBA storage");
    {
        let dir_container = compress_container(&[]);
        let mut s = ole.create_stream("VBA/dir").expect("dir stream");
        s.write_all(&dir_container).expect("write dir stream");
    }

    let vba_project_bin = ole.into_inner().into_inner();
    let normalized =
        project_normalized_data(&vba_project_bin).expect("compute ProjectNormalizedData");

    assert_eq!(
        normalized,
        b"Host Extender InfoHostExtenderRef=MyHostExtender",
        "expected Host Extender Info contribution even when PROJECT stream starts with a UTF-8 BOM"
    );
    assert!(
        !normalized.contains(&0xEF) && !normalized.contains(&0xBB) && !normalized.contains(&0xBF),
        "expected BOM bytes to be stripped from the transcript"
    );
}
