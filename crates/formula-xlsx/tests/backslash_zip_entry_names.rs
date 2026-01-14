use std::collections::HashMap;
use std::io::{Cursor, Read, Write};

use formula_xlsx::{
    patch_xlsx_streaming_workbook_cell_patches_with_part_overrides, strip_vba_project_streaming,
    PartOverride, WorkbookCellPatches,
};
use zip::write::FileOptions;
use zip::{CompressionMethod, ZipArchive, ZipWriter};

fn build_zip(entries: &[(&str, &[u8])]) -> Vec<u8> {
    let mut buf = Cursor::new(Vec::new());
    let mut zip = ZipWriter::new(&mut buf);
    let options = FileOptions::<()>::default().compression_method(CompressionMethod::Deflated);
    for (name, bytes) in entries {
        zip.start_file(name, options).expect("start_file");
        zip.write_all(bytes).expect("write file bytes");
    }
    zip.finish().expect("finish zip");
    buf.into_inner()
}

#[test]
fn streaming_part_override_matches_backslash_entry_names_and_preserves_raw_name(
) -> Result<(), Box<dyn std::error::Error>> {
    let input = build_zip(&[
        ("xl\\workbook.xml", b"<workbook before=\"1\"/>"),
        ("xl\\worksheets\\sheet1.xml", b"<worksheet/>"),
    ]);

    let mut overrides = HashMap::new();
    overrides.insert(
        "xl/workbook.xml".to_string(),
        PartOverride::Replace(b"<workbook after=\"1\"/>".to_vec()),
    );

    let patches = WorkbookCellPatches::default();
    let mut out = Cursor::new(Vec::new());
    patch_xlsx_streaming_workbook_cell_patches_with_part_overrides(
        Cursor::new(input),
        &mut out,
        &patches,
        &overrides,
    )?;
    let out_bytes = out.into_inner();

    // The replaced part should keep its original raw ZIP entry name (`\` separators).
    let mut archive = ZipArchive::new(Cursor::new(&out_bytes))?;
    let out_names: Vec<String> = archive.file_names().map(str::to_string).collect();
    assert!(
        out_names.iter().any(|n| n == "xl\\workbook.xml"),
        "expected output zip to preserve raw entry name xl\\\\workbook.xml"
    );
    assert!(
        !out_names.iter().any(|n| n == "xl/workbook.xml"),
        "expected output zip to not introduce a normalized xl/workbook.xml entry"
    );

    let mut updated = String::new();
    archive
        .by_name("xl\\workbook.xml")?
        .read_to_string(&mut updated)?;
    assert_eq!(updated, "<workbook after=\"1\"/>");

    Ok(())
}

#[test]
fn macro_strip_streaming_deletes_backslash_macrosheet_parts(
) -> Result<(), Box<dyn std::error::Error>> {
    let input = build_zip(&[
        ("xl\\macrosheets\\sheet1.xml", b"<worksheet/>"),
        ("xl\\workbook.xml", b"<workbook/>"),
    ]);

    let mut out = Cursor::new(Vec::new());
    strip_vba_project_streaming(Cursor::new(input), &mut out)?;
    let out_bytes = out.into_inner();

    let archive = ZipArchive::new(Cursor::new(&out_bytes))?;
    let out_names: Vec<String> = archive.file_names().map(str::to_string).collect();
    assert!(
        !out_names.iter().any(|n| n == "xl\\macrosheets\\sheet1.xml"),
        "expected xl\\\\macrosheets\\\\sheet1.xml to be removed"
    );
    assert!(
        out_names.iter().any(|n| n == "xl\\workbook.xml"),
        "expected non-macro parts to be preserved"
    );

    Ok(())
}
