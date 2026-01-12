#[path = "../examples/shared/dir_record_names.rs"]
mod dir_record_names;

#[test]
fn unicode_dir_record_ids_match_ms_ovba_table() {
    // MS-OVBA §2.3.4 (dir Stream) record IDs for Unicode/alternate string variants that are
    // important when debugging non-ASCII projects.
    assert_eq!(
        dir_record_names::record_name(0x0040),
        Some("PROJECTDOCSTRINGUNICODE")
    );
    assert_eq!(
        dir_record_names::record_name(0x0041),
        Some("PROJECTDOCSTRINGUNICODE (alt id 0x0041)")
    );
    assert_eq!(
        dir_record_names::record_name(0x003D),
        Some("PROJECTHELPFILEPATH2")
    );
    assert_eq!(
        dir_record_names::record_name(0x0042),
        Some("PROJECTHELPFILEPATH2 (alt id 0x0042)")
    );
    assert_eq!(
        dir_record_names::record_name(0x003C),
        Some("PROJECTCONSTANTSUNICODE")
    );
    assert_eq!(
        dir_record_names::record_name(0x0043),
        Some("PROJECTCONSTANTSUNICODE (alt id 0x0043)")
    );

    assert_eq!(dir_record_names::record_name(0x0047), Some("MODULENAMEUNICODE"));
    assert_eq!(
        dir_record_names::record_name(0x0032),
        Some("MODULESTREAMNAMEUNICODE")
    );
    assert_eq!(
        dir_record_names::record_name(0x0048),
        Some("MODULEDOCSTRINGUNICODE")
    );
    assert_eq!(
        dir_record_names::record_name(0x0049),
        Some("MODULEHELPFILEPATHUNICODE")
    );
    assert_eq!(
        dir_record_names::record_name(0x003E),
        Some("REFERENCENAMEUNICODE (reserved marker)")
    );
    assert_eq!(
        dir_record_names::record_name(0x004A),
        Some("PROJECTCOMPATVERSION")
    );
}
