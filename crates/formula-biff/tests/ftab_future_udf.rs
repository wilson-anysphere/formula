use formula_biff::{function_id_from_name, FTAB_USER_DEFINED};

#[test]
fn minifs_maxifs_map_to_user_defined_function_id() {
    for name in ["MINIFS", "MAXIFS", "_xlfn.MINIFS", "_xlfn.MAXIFS"] {
        assert_eq!(
            function_id_from_name(name),
            Some(FTAB_USER_DEFINED),
            "{name} should map to the BIFF UDF sentinel id"
        );
    }
}

