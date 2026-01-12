use formula_xlsx::pivots::preserve::apply_preserved_pivot_caches_to_workbook_xml;
use roxmltree::Document;

fn count(haystack: &str, needle: &str) -> usize {
    haystack.match_indices(needle).count()
}

const PIVOT_CACHES: &str =
    r#"<pivotCaches count="1"><pivotCache cacheId="1" r:id="rId10"/></pivotCaches>"#;

#[test]
fn inserts_pivot_caches_into_minimal_workbook() {
    let workbook = r#"<workbook xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main"><sheets/></workbook>"#;
    let updated =
        apply_preserved_pivot_caches_to_workbook_xml(workbook, PIVOT_CACHES).expect("patch");

    assert_eq!(count(&updated, "<pivotCaches"), 1);
    assert!(
        updated.contains(r#"xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships""#),
        "should add xmlns:r when inserting r:id attributes"
    );

    let pivot_pos = updated.find("<pivotCaches").unwrap();
    let close_pos = updated.find("</workbook>").unwrap();
    assert!(pivot_pos < close_pos, "<pivotCaches> should be inside <workbook>");
}

#[test]
fn inserts_before_ext_lst() {
    let workbook = r#"<workbook xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main"><sheets/><definedNames/><calcPr/><extLst><ext/></extLst></workbook>"#;
    let updated =
        apply_preserved_pivot_caches_to_workbook_xml(workbook, PIVOT_CACHES).expect("patch");

    assert_eq!(count(&updated, "<pivotCaches"), 1);
    let pivot_pos = updated.find("<pivotCaches").unwrap();
    let ext_pos = updated.find("<extLst").unwrap();
    assert!(pivot_pos < ext_pos, "<pivotCaches> must be inserted before <extLst>");
}

#[test]
fn inserts_before_file_recovery_pr() {
    let workbook = r#"<workbook xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main"><sheets/><fileRecoveryPr/><extLst><ext/></extLst></workbook>"#;
    let updated =
        apply_preserved_pivot_caches_to_workbook_xml(workbook, PIVOT_CACHES).expect("patch");

    assert_eq!(count(&updated, "<pivotCaches"), 1);
    let pivot_pos = updated.find("<pivotCaches").unwrap();
    let fr_pos = updated.find("<fileRecoveryPr").unwrap();
    assert!(
        pivot_pos < fr_pos,
        "<pivotCaches> must be inserted before <fileRecoveryPr>"
    );
}

#[test]
fn merges_when_workbook_already_has_pivot_caches() {
    let workbook = r#"<workbook xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships"><sheets/><pivotCaches count="1"><pivotCache cacheId="1" r:id="rId10"/></pivotCaches><extLst><ext/></extLst></workbook>"#;
    let preserved = r#"<pivotCaches count="2"><pivotCache cacheId="1" r:id="rId10"/><pivotCache cacheId="2" r:id="rId11"/></pivotCaches>"#;

    let updated =
        apply_preserved_pivot_caches_to_workbook_xml(workbook, preserved).expect("patch");

    assert_eq!(count(&updated, "<pivotCaches"), 1);
    assert_eq!(count(&updated, r#"cacheId="1""#), 1);
    assert_eq!(count(&updated, r#"cacheId="2""#), 1);

    let pivot_pos = updated.find("<pivotCaches").unwrap();
    let ext_pos = updated.find("<extLst").unwrap();
    assert!(pivot_pos < ext_pos, "<pivotCaches> must remain before <extLst>");
}

#[test]
fn inserts_into_self_closing_prefixed_workbook_root() {
    let workbook =
        r#"<x:workbook xmlns:x="http://schemas.openxmlformats.org/spreadsheetml/2006/main"/>"#;
    let fragment = r#"<x:pivotCaches><x:pivotCache cacheId="1" r:id="rId1"/></x:pivotCaches>"#;

    let updated = apply_preserved_pivot_caches_to_workbook_xml(workbook, fragment).expect("patch");

    Document::parse(&updated).expect("output should be parseable XML");
    assert!(updated.contains("<x:pivotCaches"), "missing inserted block: {updated}");
    assert!(
        updated.contains("</x:workbook>"),
        "missing expanded root close tag: {updated}"
    );
    assert!(
        !updated.contains("</workbook>"),
        "introduced unprefixed close tag: {updated}"
    );
    assert!(
        updated.contains(r#"xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships""#),
        "missing xmlns:r declaration: {updated}"
    );
}

#[test]
fn inserts_into_self_closing_default_ns_workbook_root() {
    let workbook =
        r#"<workbook xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main"/>"#;

    let updated = apply_preserved_pivot_caches_to_workbook_xml(workbook, PIVOT_CACHES).expect("patch");

    Document::parse(&updated).expect("output should be parseable XML");
    assert_eq!(count(&updated, "<pivotCaches"), 1);
    let pivot_pos = updated.find("<pivotCaches").unwrap();
    let close_pos = updated.find("</workbook>").unwrap();
    assert!(
        pivot_pos < close_pos,
        "<pivotCaches> should be inside <workbook>: {updated}"
    );
}
