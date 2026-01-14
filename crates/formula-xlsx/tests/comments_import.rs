use std::fs;
use std::io::{Cursor, Write};

use formula_model::{CellRef, CommentKind, Workbook};
use zip::write::FileOptions;
use zip::CompressionMethod;

fn assert_fixture_comments_present(workbook: &Workbook) {
    let sheet = workbook
        .sheets
        .iter()
        .find(|s| s.name == "Sheet1")
        .expect("fixture should contain Sheet1");

    let mut note = None;
    let mut threaded = None;
    for (_anchor, comment) in sheet.iter_comments() {
        match comment.kind {
            CommentKind::Note => note = Some(comment),
            CommentKind::Threaded => threaded = Some(comment),
        }
    }

    let note = note.expect("fixture should contain a legacy note comment");
    assert_eq!(note.cell_ref, CellRef::from_a1("A1").unwrap());
    assert_eq!(note.author.name, "Alex");
    assert_eq!(note.content, "Legacy note");

    let threaded = threaded.expect("fixture should contain a threaded comment");
    assert_eq!(threaded.cell_ref, CellRef::from_a1("B2").unwrap());
    assert_eq!(threaded.author.name, "Alex");
    assert_eq!(threaded.content, "Thread root");
    assert!(threaded.resolved, "fixture threaded comment should be resolved");

    let reply = threaded
        .replies
        .first()
        .expect("fixture threaded comment should have a reply");
    assert_eq!(reply.author.name, "Sam");
    assert_eq!(reply.content, "First reply");
}

fn make_minimal_xlsx_with_missing_comment_part() -> Vec<u8> {
    // This fixture intentionally omits the referenced comment part (xl/comments1.xml).
    //
    // The loader should treat missing comment parts as best-effort (ignore), not as a hard error.
    const WORKBOOK_XML: &str = r#"<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<workbook xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main"
          xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">
  <sheets>
    <sheet name="Sheet1" sheetId="1" r:id="rId1"/>
  </sheets>
</workbook>
"#;

    const WORKBOOK_RELS_XML: &str = r#"<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
  <Relationship Id="rId1"
    Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/worksheet"
    Target="worksheets/sheet1.xml"/>
</Relationships>
"#;

    const SHEET_XML: &str = r#"<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main"/>
"#;

    const SHEET_RELS_XML: &str = r#"<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
  <Relationship Id="rId1"
    Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/comments"
    Target="../comments1.xml"/>
</Relationships>
"#;

    let cursor = Cursor::new(Vec::<u8>::new());
    let mut zip = zip::ZipWriter::new(cursor);
    let options: FileOptions<'_, ()> =
        FileOptions::default().compression_method(CompressionMethod::Stored);

    zip.start_file("xl/workbook.xml", options)
        .expect("zip entry creation should succeed");
    zip.write_all(WORKBOOK_XML.as_bytes())
        .expect("workbook xml write should succeed");

    zip.start_file("xl/_rels/workbook.xml.rels", options)
        .expect("zip entry creation should succeed");
    zip.write_all(WORKBOOK_RELS_XML.as_bytes())
        .expect("workbook rels write should succeed");

    zip.start_file("xl/worksheets/sheet1.xml", options)
        .expect("zip entry creation should succeed");
    zip.write_all(SHEET_XML.as_bytes())
        .expect("worksheet xml write should succeed");

    zip.start_file("xl/worksheets/_rels/sheet1.xml.rels", options)
        .expect("zip entry creation should succeed");
    zip.write_all(SHEET_RELS_XML.as_bytes())
        .expect("worksheet rels write should succeed");

    zip.finish()
        .expect("zip finalization should succeed")
        .into_inner()
}

fn make_minimal_xlsx_with_non_canonical_part_name_casing() -> Vec<u8> {
    // This fixture intentionally stores comment + persons parts with non-canonical (upper) casing
    // to ensure both the full loader (`load_from_bytes`) and fast reader
    // (`read_workbook_model_from_bytes`) tolerate common producer mistakes.
    const WORKBOOK_XML: &str = r#"<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<workbook xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main"
          xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">
  <sheets>
    <sheet name="Sheet1" sheetId="1" r:id="rId1"/>
  </sheets>
</workbook>
"#;

    // Note: no workbook-level `.../relationships/person` relationship is included; the loader
    // must discover `xl/persons/*.xml` parts by scanning the package.
    const WORKBOOK_RELS_XML: &str = r#"<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
  <Relationship Id="rId1"
    Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/worksheet"
    Target="worksheets/sheet1.xml"/>
</Relationships>
"#;

    const SHEET_XML: &str = r#"<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main"/>
"#;

    const SHEET_RELS_XML: &str = r#"<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
  <Relationship Id="rId1"
    Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/comments"
    Target="../comments1.xml"/>
  <Relationship Id="rId2"
    Type="http://schemas.microsoft.com/office/2017/10/relationships/threadedComment"
    Target="../threadedComments/threadedComments1.xml"/>
</Relationships>
"#;

    const LEGACY_COMMENTS_XML: &str = r#"<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<comments xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main">
  <authors>
    <author>Alex</author>
  </authors>
  <commentList>
    <comment ref="A1" authorId="0">
      <text><r><t xml:space="preserve">Legacy note</t></r></text>
    </comment>
  </commentList>
</comments>
"#;

    const THREADED_COMMENTS_XML: &str = r#"<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<threadedComments xmlns="http://schemas.microsoft.com/office/spreadsheetml/2018/threadedcomments">
  <threadedComment id="t1" ref="B2" personId="p1" done="1">
    <text><r><t xml:space="preserve">Thread root</t></r></text>
  </threadedComment>
  <threadedComment id="t2" parentId="t1" ref="B2" personId="p2">
    <text><r><t xml:space="preserve">First reply</t></r></text>
  </threadedComment>
</threadedComments>
"#;

    const PERSONS_XML: &str = r#"<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<persons xmlns="http://schemas.microsoft.com/office/spreadsheetml/2018/threadedcomments">
  <person id="p1" displayName="Alex"/>
  <person id="p2" displayName="Sam"/>
</persons>
"#;

    let cursor = Cursor::new(Vec::<u8>::new());
    let mut zip = zip::ZipWriter::new(cursor);
    let options: FileOptions<'_, ()> =
        FileOptions::default().compression_method(CompressionMethod::Stored);

    zip.start_file("xl/workbook.xml", options)
        .expect("zip entry creation should succeed");
    zip.write_all(WORKBOOK_XML.as_bytes())
        .expect("workbook xml write should succeed");

    zip.start_file("xl/_rels/workbook.xml.rels", options)
        .expect("zip entry creation should succeed");
    zip.write_all(WORKBOOK_RELS_XML.as_bytes())
        .expect("workbook rels write should succeed");

    zip.start_file("xl/worksheets/sheet1.xml", options)
        .expect("zip entry creation should succeed");
    zip.write_all(SHEET_XML.as_bytes())
        .expect("worksheet xml write should succeed");

    zip.start_file("xl/worksheets/_rels/sheet1.xml.rels", options)
        .expect("zip entry creation should succeed");
    zip.write_all(SHEET_RELS_XML.as_bytes())
        .expect("worksheet rels write should succeed");

    // Non-canonical casing on comment + persons parts.
    zip.start_file("XL/COMMENTS1.XML", options)
        .expect("zip entry creation should succeed");
    zip.write_all(LEGACY_COMMENTS_XML.as_bytes())
        .expect("legacy comments xml write should succeed");

    zip.start_file("XL/THREADEDCOMMENTS/THREADEDCOMMENTS1.XML", options)
        .expect("zip entry creation should succeed");
    zip.write_all(THREADED_COMMENTS_XML.as_bytes())
        .expect("threaded comments xml write should succeed");

    zip.start_file("XL/PERSONS/PERSONS1.XML", options)
        .expect("zip entry creation should succeed");
    zip.write_all(PERSONS_XML.as_bytes())
        .expect("persons xml write should succeed");

    zip.finish()
        .expect("zip finalization should succeed")
        .into_inner()
}

fn make_minimal_xlsx_with_percent_encoded_comment_targets() -> Vec<u8> {
    // This fixture stores comment parts using canonical names (e.g. `xl/comments1.xml`), but the
    // worksheet `.rels` references them with unnecessary percent-encoding (e.g. `comments%31.xml`).
    //
    // Relationship targets are URIs, so consumers should treat percent-encoded and decoded names
    // as equivalent. The fast reader already tolerates this via `open_zip_part`; the full loader
    // (`load_from_bytes`) should too.
    const WORKBOOK_XML: &str = r#"<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<workbook xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main"
          xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">
  <sheets>
    <sheet name="Sheet1" sheetId="1" r:id="rId1"/>
  </sheets>
</workbook>
"#;

    const WORKBOOK_RELS_XML: &str = r#"<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
  <Relationship Id="rId1"
    Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/worksheet"
    Target="worksheets/sheet1.xml"/>
</Relationships>
"#;

    const SHEET_XML: &str = r#"<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main"/>
"#;

    const SHEET_RELS_XML: &str = r#"<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
  <Relationship Id="rId1"
    Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/comments"
    Target="../comments%31.xml"/>
  <Relationship Id="rId2"
    Type="http://schemas.microsoft.com/office/2017/10/relationships/threadedComment"
    Target="../threadedComments/threadedComments%31.xml"/>
</Relationships>
"#;

    const LEGACY_COMMENTS_XML: &str = r#"<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<comments xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main">
  <authors>
    <author>Alex</author>
  </authors>
  <commentList>
    <comment ref="A1" authorId="0">
      <text><r><t xml:space="preserve">Legacy note</t></r></text>
    </comment>
  </commentList>
</comments>
"#;

    const THREADED_COMMENTS_XML: &str = r#"<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<threadedComments xmlns="http://schemas.microsoft.com/office/spreadsheetml/2018/threadedcomments">
  <threadedComment id="t1" ref="B2" personId="p1" done="1">
    <text><r><t xml:space="preserve">Thread root</t></r></text>
  </threadedComment>
  <threadedComment id="t2" parentId="t1" ref="B2" personId="p2">
    <text><r><t xml:space="preserve">First reply</t></r></text>
  </threadedComment>
</threadedComments>
"#;

    const PERSONS_XML: &str = r#"<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<persons xmlns="http://schemas.microsoft.com/office/spreadsheetml/2018/threadedcomments">
  <person id="p1" displayName="Alex"/>
  <person id="p2" displayName="Sam"/>
</persons>
"#;

    let cursor = Cursor::new(Vec::<u8>::new());
    let mut zip = zip::ZipWriter::new(cursor);
    let options: FileOptions<'_, ()> =
        FileOptions::default().compression_method(CompressionMethod::Stored);

    zip.start_file("xl/workbook.xml", options)
        .expect("zip entry creation should succeed");
    zip.write_all(WORKBOOK_XML.as_bytes())
        .expect("workbook xml write should succeed");

    zip.start_file("xl/_rels/workbook.xml.rels", options)
        .expect("zip entry creation should succeed");
    zip.write_all(WORKBOOK_RELS_XML.as_bytes())
        .expect("workbook rels write should succeed");

    zip.start_file("xl/worksheets/sheet1.xml", options)
        .expect("zip entry creation should succeed");
    zip.write_all(SHEET_XML.as_bytes())
        .expect("worksheet xml write should succeed");

    zip.start_file("xl/worksheets/_rels/sheet1.xml.rels", options)
        .expect("zip entry creation should succeed");
    zip.write_all(SHEET_RELS_XML.as_bytes())
        .expect("worksheet rels write should succeed");

    // Canonical part names (unescaped).
    zip.start_file("xl/comments1.xml", options)
        .expect("zip entry creation should succeed");
    zip.write_all(LEGACY_COMMENTS_XML.as_bytes())
        .expect("legacy comments xml write should succeed");

    zip.start_file("xl/threadedComments/threadedComments1.xml", options)
        .expect("zip entry creation should succeed");
    zip.write_all(THREADED_COMMENTS_XML.as_bytes())
        .expect("threaded comments xml write should succeed");

    zip.start_file("xl/persons/persons1.xml", options)
        .expect("zip entry creation should succeed");
    zip.write_all(PERSONS_XML.as_bytes())
        .expect("persons xml write should succeed");

    zip.finish()
        .expect("zip finalization should succeed")
        .into_inner()
}

#[test]
fn load_from_bytes_imports_sheet_comments() {
    let fixture_path = concat!(env!("CARGO_MANIFEST_DIR"), "/tests/fixtures/comments.xlsx");
    let bytes = fs::read(fixture_path).expect("fixture workbook should be readable");

    let doc = formula_xlsx::load_from_bytes(&bytes).expect("load_from_bytes should succeed");
    assert_fixture_comments_present(&doc.workbook);
}

#[test]
fn fast_reader_imports_sheet_comments() {
    let fixture_path = concat!(env!("CARGO_MANIFEST_DIR"), "/tests/fixtures/comments.xlsx");
    let bytes = fs::read(fixture_path).expect("fixture workbook should be readable");

    let workbook =
        formula_xlsx::read_workbook_model_from_bytes(&bytes).expect("fast reader should succeed");
    assert_fixture_comments_present(&workbook);
}

#[test]
fn load_from_bytes_ignores_missing_comment_parts() {
    let bytes = make_minimal_xlsx_with_missing_comment_part();
    let doc = formula_xlsx::load_from_bytes(&bytes)
        .expect("load_from_bytes should succeed even if comment part is missing");

    let sheet = doc
        .workbook
        .sheets
        .iter()
        .find(|s| s.name == "Sheet1")
        .expect("fixture should contain Sheet1");
    assert_eq!(
        sheet.iter_comments().count(),
        0,
        "missing comment parts should not create comments"
    );
}

#[test]
fn fast_reader_ignores_missing_comment_parts() {
    let bytes = make_minimal_xlsx_with_missing_comment_part();
    let workbook = formula_xlsx::read_workbook_model_from_bytes(&bytes)
        .expect("fast reader should succeed even if comment part is missing");

    let sheet = workbook
        .sheets
        .iter()
        .find(|s| s.name == "Sheet1")
        .expect("fixture should contain Sheet1");
    assert_eq!(
        sheet.iter_comments().count(),
        0,
        "missing comment parts should not create comments"
    );
}

#[test]
fn load_from_bytes_imports_comments_with_non_canonical_part_name_casing() {
    let bytes = make_minimal_xlsx_with_non_canonical_part_name_casing();
    let doc = formula_xlsx::load_from_bytes(&bytes)
        .expect("load_from_bytes should tolerate non-canonical part names");
    assert_fixture_comments_present(&doc.workbook);
}

#[test]
fn fast_reader_imports_comments_with_non_canonical_part_name_casing() {
    let bytes = make_minimal_xlsx_with_non_canonical_part_name_casing();
    let workbook = formula_xlsx::read_workbook_model_from_bytes(&bytes)
        .expect("fast reader should tolerate non-canonical part names");
    assert_fixture_comments_present(&workbook);
}

#[test]
fn load_from_bytes_imports_comments_with_percent_encoded_relationship_targets() {
    let bytes = make_minimal_xlsx_with_percent_encoded_comment_targets();
    let doc = formula_xlsx::load_from_bytes(&bytes)
        .expect("load_from_bytes should tolerate percent-encoded relationship targets");
    assert_fixture_comments_present(&doc.workbook);
}

#[test]
fn fast_reader_imports_comments_with_percent_encoded_relationship_targets() {
    let bytes = make_minimal_xlsx_with_percent_encoded_comment_targets();
    let workbook = formula_xlsx::read_workbook_model_from_bytes(&bytes)
        .expect("fast reader should tolerate percent-encoded relationship targets");
    assert_fixture_comments_present(&workbook);
}
