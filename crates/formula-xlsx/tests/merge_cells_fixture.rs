use formula_model::{Alignment, CellRef, HorizontalAlignment, Range, VerticalAlignment};
use formula_xlsx::merge_cells::read_merge_cells_from_xlsx;
use formula_xlsx::styles::parse_cell_xfs_alignments;
use formula_xlsx::write_minimal_xlsx;
use pretty_assertions::assert_eq;
use std::fs;
use std::io::Read;
use std::io::Cursor;
use std::path::Path;
use zip::ZipArchive;

fn fixture_bytes() -> Vec<u8> {
    let path = Path::new(env!("CARGO_MANIFEST_DIR")).join("tests/fixtures/merged-cells.xlsx");
    fs::read(&path).expect("missing fixture xlsx")
}

#[test]
fn parses_fixture_merge_cells_and_alignment() {
    let bytes = fixture_bytes();
    let mut archive = ZipArchive::new(Cursor::new(bytes)).expect("zip open");

    let merges =
        read_merge_cells_from_xlsx(&mut archive, "xl/worksheets/sheet1.xml").expect("merges");
    assert_eq!(
        merges,
        vec![Range::new(CellRef::new(0, 0), CellRef::new(1, 1))]
    );

    let mut styles_xml = String::new();
    archive
        .by_name("xl/styles.xml")
        .expect("styles.xml")
        .read_to_string(&mut styles_xml)
        .expect("read styles");

    let alignments = parse_cell_xfs_alignments(&styles_xml).expect("alignments");
    assert!(
        alignments.len() >= 2,
        "expected at least 2 xfs, got {}",
        alignments.len()
    );

    assert_eq!(
        alignments[1],
        Alignment {
            horizontal: Some(HorizontalAlignment::Center),
            vertical: Some(VerticalAlignment::Center),
            wrap_text: true,
            text_rotation: 45,
        }
    );
}

#[test]
fn merge_cells_round_trip_preserved_on_save_load() {
    let bytes = fixture_bytes();
    let mut archive = ZipArchive::new(Cursor::new(bytes)).expect("zip open");

    let merges =
        read_merge_cells_from_xlsx(&mut archive, "xl/worksheets/sheet1.xml").expect("merges");

    let mut styles_xml = String::new();
    archive
        .by_name("xl/styles.xml")
        .expect("styles.xml")
        .read_to_string(&mut styles_xml)
        .expect("read styles");
    let alignments = parse_cell_xfs_alignments(&styles_xml).expect("alignments");

    let out = write_minimal_xlsx(&merges, &alignments).expect("write xlsx");
    let mut archive2 = ZipArchive::new(Cursor::new(out)).expect("zip open 2");

    let merges2 =
        read_merge_cells_from_xlsx(&mut archive2, "xl/worksheets/sheet1.xml").expect("merges2");
    assert_eq!(merges2, merges);
}
