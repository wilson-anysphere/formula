use std::io::{Cursor, Write};

use formula_xlsx::{read_workbook_model_from_bytes, ReadError, XlsxError};
use zip::write::FileOptions;
use zip::ZipWriter;

fn build_zip_bytes(entries: &[(&str, &[u8])]) -> Vec<u8> {
    let cursor = Cursor::new(Vec::new());
    let mut zip = ZipWriter::new(cursor);
    let options = FileOptions::<()>::default().compression_method(zip::CompressionMethod::Stored);

    for (name, bytes) in entries {
        zip.start_file(*name, options).unwrap();
        zip.write_all(bytes).unwrap();
    }

    zip.finish().unwrap().into_inner()
}

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
fn read_workbook_model_rejects_parts_whose_declared_uncompressed_size_exceeds_limit() {
    // Construct a minimal XLSX-like ZIP container where `xl/workbook.xml` is a tiny payload but
    // forges ZIP metadata to claim an uncompressed size far above typical per-part limits.
    //
    // The model reader should reject this deterministically without attempting a huge allocation.
    let bytes = build_zip_bytes(&[("xl/workbook.xml", b"hello")]);
    let bytes = patch_zip_entry_uncompressed_size(bytes, "xl/workbook.xml", 1_000_000_000);

    let err = read_workbook_model_from_bytes(&bytes).expect_err("expected oversized part to error");
    match err {
        ReadError::Xlsx(XlsxError::PartTooLarge { part, size, max }) => {
            assert_eq!(part, "xl/workbook.xml");
            assert_eq!(size, 1_000_000_000);
            assert!(size > max);
        }
        other => panic!("expected ReadError::Xlsx(PartTooLarge), got {other:?}"),
    }
}

