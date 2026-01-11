use formula_xlsb::ftab::{function_id_from_name, function_name_from_id, FTAB_USER_DEFINED};

#[test]
fn formula_engine_functions_have_ftab_coverage() {
    for spec in formula_engine::functions::iter_function_specs() {
        let id = function_id_from_name(spec.name)
            .unwrap_or_else(|| panic!("missing BIFF function-id mapping for {}", spec.name));

        // Built-in functions should round-trip id -> name -> id. For newer
        // forward-compatible functions we intentionally map to the BIFF UDF
        // sentinel (255) because they are typically encoded with `_xlfn.`.
        if id != FTAB_USER_DEFINED {
            let roundtrip_name = function_name_from_id(id)
                .unwrap_or_else(|| panic!("missing BIFF name mapping for id {id} ({})", spec.name));
            assert_eq!(roundtrip_name, spec.name, "ftab mismatch for {}", spec.name);
        }
    }
}

#[test]
fn xlfn_prefix_is_ignored_for_lookup() {
    let base = function_id_from_name("XLOOKUP").expect("XLOOKUP should have an encoding");
    let prefixed = function_id_from_name("_xlfn.XLOOKUP").expect("_xlfn.XLOOKUP should have an encoding");
    assert_eq!(base, prefixed);

    // Unknown `_xlfn.` functions should still be encodable as BIFF UDF calls.
    assert_eq!(
        function_id_from_name("_xlfn.SOME_FUTURE_FUNCTION"),
        Some(FTAB_USER_DEFINED)
    );
}

