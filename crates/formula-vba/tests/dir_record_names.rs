#[path = "../examples/shared/dir_record_names.rs"]
mod dir_record_names;

#[test]
fn unicode_dir_record_ids_match_ms_ovba_table() {
    // MS-OVBA §2.3.4 (dir Stream) record IDs for Unicode string variants.
    assert_eq!(
        dir_record_names::record_name(0x0040),
        Some("PROJECTNAMEUNICODE")
    );
    assert_eq!(
        dir_record_names::record_name(0x0041),
        Some("PROJECTDOCSTRINGUNICODE")
    );
    assert_eq!(
        dir_record_names::record_name(0x0042),
        Some("PROJECTHELPFILEPATHUNICODE")
    );
    assert_eq!(
        dir_record_names::record_name(0x0043),
        Some("PROJECTCONSTANTSUNICODE")
    );

    assert_eq!(dir_record_names::record_name(0x0047), Some("MODULENAMEUNICODE"));
    assert_eq!(
        dir_record_names::record_name(0x0048),
        Some("MODULESTREAMNAMEUNICODE")
    );
    assert_eq!(
        dir_record_names::record_name(0x0049),
        Some("MODULEDOCSTRINGUNICODE")
    );
    assert_eq!(
        dir_record_names::record_name(0x004A),
        Some("PROJECTCOMPATVERSION / MODULEHELPFILEPATHUNICODE")
    );
}
