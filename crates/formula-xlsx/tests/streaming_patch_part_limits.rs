use std::io::Cursor;

use formula_model::{CellRef, CellValue, Workbook};
use formula_xlsx::{patch_xlsx_streaming, write_workbook_to_writer, StreamingPatchError, WorksheetCellPatch, XlsxError};
use zip::ZipArchive;

fn patch_zip_entry_uncompressed_size(
    mut zip_bytes: Vec<u8>,
    entry_name: &str,
    new_uncompressed_size: u32,
) -> Vec<u8> {
    // Locate the end-of-central-directory record (EOCD) by scanning backwards from the end of the
    // file. The ZIP spec allows up to 64KiB of trailing comment.
    const EOCD_SIG: [u8; 4] = [0x50, 0x4B, 0x05, 0x06];
    let min_eocd = zip_bytes.len().saturating_sub(22);
    let search_min = zip_bytes.len().saturating_sub(22 + 0xFFFF);

    let mut eocd_offset = None;
    for i in (search_min..=min_eocd).rev() {
        if zip_bytes.get(i..i + 4) == Some(&EOCD_SIG) {
            eocd_offset = Some(i);
            break;
        }
    }
    let eocd_offset = eocd_offset.expect("expected EOCD record in test zip");

    let central_dir_size = u32::from_le_bytes(
        zip_bytes[eocd_offset + 12..eocd_offset + 16]
            .try_into()
            .unwrap(),
    ) as usize;
    let central_dir_offset = u32::from_le_bytes(
        zip_bytes[eocd_offset + 16..eocd_offset + 20]
            .try_into()
            .unwrap(),
    ) as usize;

    const CEN_SIG: [u8; 4] = [0x50, 0x4B, 0x01, 0x02];
    const LFH_SIG: [u8; 4] = [0x50, 0x4B, 0x03, 0x04];
    let mut cursor = central_dir_offset;
    let end = central_dir_offset + central_dir_size;
    while cursor < end {
        assert_eq!(
            zip_bytes.get(cursor..cursor + 4),
            Some(CEN_SIG.as_slice()),
            "expected central directory header signature"
        );

        let name_len = u16::from_le_bytes(zip_bytes[cursor + 28..cursor + 30].try_into().unwrap())
            as usize;
        let extra_len = u16::from_le_bytes(zip_bytes[cursor + 30..cursor + 32].try_into().unwrap())
            as usize;
        let comment_len =
            u16::from_le_bytes(zip_bytes[cursor + 32..cursor + 34].try_into().unwrap()) as usize;
        let local_header_offset =
            u32::from_le_bytes(zip_bytes[cursor + 42..cursor + 46].try_into().unwrap()) as usize;

        let name_start = cursor + 46;
        let name_end = name_start + name_len;
        let name = std::str::from_utf8(&zip_bytes[name_start..name_end])
            .expect("expected UTF-8 entry name");

        if name == entry_name {
            // Patch central directory header's uncompressed size (offset 24, 4 bytes).
            zip_bytes[cursor + 24..cursor + 28]
                .copy_from_slice(&new_uncompressed_size.to_le_bytes());

            // Patch local file header's uncompressed size too (offset 22, 4 bytes).
            assert_eq!(
                zip_bytes.get(local_header_offset..local_header_offset + 4),
                Some(LFH_SIG.as_slice()),
                "expected local file header signature"
            );
            zip_bytes[local_header_offset + 22..local_header_offset + 26]
                .copy_from_slice(&new_uncompressed_size.to_le_bytes());
            return zip_bytes;
        }

        cursor += 46 + name_len + extra_len + comment_len;
    }

    panic!("test zip did not contain expected entry: {entry_name}");
}

#[test]
fn patch_xlsx_streaming_rejects_oversized_workbook_rels_part() {
    // Build a small, valid XLSX package via the writer, then forge the declared uncompressed size
    // of `xl/_rels/workbook.xml.rels` to exceed the per-part limit. The streaming patcher should
    // fail fast (before allocating) when it tries to read that part to resolve sharedStrings.
    let mut workbook = Workbook::new();
    workbook.add_sheet("Sheet1").unwrap();

    let mut out = Cursor::new(Vec::new());
    write_workbook_to_writer(&workbook, &mut out).expect("write workbook");
    let mut bytes = out.into_inner();

    // Sanity check: the part exists in the zip produced by the writer.
    {
        let mut zip = ZipArchive::new(Cursor::new(bytes.as_slice())).expect("open zip");
        zip.by_name("xl/_rels/workbook.xml.rels")
            .expect("expected workbook rels part");
    }

    bytes = patch_zip_entry_uncompressed_size(bytes, "xl/_rels/workbook.xml.rels", 1_000_000_000);

    let patch = WorksheetCellPatch {
        worksheet_part: "xl/worksheets/sheet1.xml".to_string(),
        cell: CellRef::new(0, 0),
        value: CellValue::String("hello".to_string()),
        formula: None,
        xf_index: None,
        vm: None,
        cm: None,
    };

    let err = patch_xlsx_streaming(Cursor::new(bytes), Cursor::new(Vec::new()), &[patch])
        .expect_err("expected streaming patch to reject oversized workbook rels");
    match err {
        StreamingPatchError::Xlsx(XlsxError::PartTooLarge { part, size, max }) => {
            assert_eq!(part, "xl/_rels/workbook.xml.rels");
            assert_eq!(size, 1_000_000_000);
            assert!(size > max);
        }
        other => panic!("expected StreamingPatchError::Xlsx(PartTooLarge), got {other:?}"),
    }
}

