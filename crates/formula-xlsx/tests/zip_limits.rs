use std::io::Cursor;
use std::io::Write;

use formula_xlsx::{read_part_from_reader_limited, XlsxError, XlsxPackage, XlsxPackageLimits};
use zip::write::FileOptions;
use zip::ZipArchive;
use zip::ZipWriter;

fn build_zip_bytes(entries: &[(&str, &[u8])]) -> Vec<u8> {
    let cursor = Cursor::new(Vec::new());
    let mut zip = ZipWriter::new(cursor);
    let options = FileOptions::<()>::default().compression_method(zip::CompressionMethod::Deflated);

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
fn read_part_from_reader_limited_rejects_oversized_parts() {
    let bytes = build_zip_bytes(&[("xl/vbaProject.bin", b"0123456789A")]); // 11 bytes

    let err = read_part_from_reader_limited(Cursor::new(bytes), "xl/vbaProject.bin", 10)
        .expect_err("expected part-too-large error");

    match err {
        XlsxError::PartTooLarge { part, size, max } => {
            assert_eq!(part, "xl/vbaProject.bin");
            assert_eq!(size, 11);
            assert_eq!(max, 10);
        }
        other => panic!("expected XlsxError::PartTooLarge, got {other:?}"),
    }
}

#[test]
fn read_part_from_reader_limited_does_not_trust_zip_uncompressed_size_metadata() {
    // Create an 11-byte payload but forge the ZIP headers to claim the uncompressed size is only
    // 10 bytes. The reader should still reject the part after observing >10 bytes while reading,
    // rather than trusting metadata alone.
    let bytes = build_zip_bytes(&[("xl/oversize.bin", b"0123456789A")]); // 11 bytes
    let bytes = patch_zip_entry_uncompressed_size(bytes, "xl/oversize.bin", 10);

    // Sanity check: zip metadata now claims the part is 10 bytes.
    {
        let mut archive = ZipArchive::new(Cursor::new(bytes.as_slice())).expect("open zip");
        let file = archive.by_name("xl/oversize.bin").expect("open part");
        assert_eq!(file.size(), 10);
    }

    let err = read_part_from_reader_limited(Cursor::new(bytes), "xl/oversize.bin", 10)
        .expect_err("expected part-too-large error");

    match err {
        XlsxError::PartTooLarge { part, size, max } => {
            assert_eq!(part, "xl/oversize.bin");
            assert_eq!(max, 10);
            assert_eq!(
                size, 11,
                "expected observed size to exceed max even though ZIP metadata was forged"
            );
        }
        other => panic!("expected XlsxError::PartTooLarge, got {other:?}"),
    }
}

#[test]
fn read_part_from_reader_limited_reads_small_parts() {
    let bytes = build_zip_bytes(&[("xl/workbook.xml", b"hello")]);

    let out =
        read_part_from_reader_limited(Cursor::new(bytes), "xl/workbook.xml", 10).unwrap();
    assert_eq!(out.as_deref(), Some(b"hello".as_slice()));
}

#[test]
fn xlsxpackage_from_bytes_limited_enforces_total_budget() {
    let bytes = build_zip_bytes(&[
        ("xl/a.bin", b"0123456789"), // 10 bytes
        ("xl/b.bin", b"0123456789"), // 10 bytes
        ("xl/c.bin", b"0123456789"), // 10 bytes
    ]);

    let limits = XlsxPackageLimits {
        max_part_bytes: 10,
        max_total_bytes: 20,
    };
    let err = XlsxPackage::from_bytes_limited(&bytes, limits)
        .expect_err("expected total-budget error");
    match err {
        XlsxError::PackageTooLarge { total, max } => {
            assert_eq!(max, 20);
            assert!(
                total > max,
                "expected reported total ({total}) to exceed max ({max})"
            );
        }
        other => panic!(
            "expected XlsxError::PackageTooLarge, got {other:?}"
        ),
    }
}

#[test]
fn xlsxpackage_from_bytes_limited_does_not_trust_zip_uncompressed_size_metadata_for_total_budget() {
    // Create an 11-byte payload but forge the ZIP headers to claim the uncompressed size is only
    // 10 bytes. When the total budget is 10 bytes, the package loader should still reject the
    // entry after observing >10 bytes while reading, rather than trusting metadata alone.
    let bytes = build_zip_bytes(&[("xl/a.bin", b"0123456789A")]); // 11 bytes
    let bytes = patch_zip_entry_uncompressed_size(bytes, "xl/a.bin", 10);

    let limits = XlsxPackageLimits {
        max_part_bytes: 20,
        max_total_bytes: 10,
    };
    let err = XlsxPackage::from_bytes_limited(&bytes, limits)
        .expect_err("expected total-budget error");
    match err {
        XlsxError::PackageTooLarge { total, max } => {
            assert_eq!(max, 10);
            assert!(
                total > max,
                "expected reported total ({total}) to exceed max ({max})"
            );
        }
        other => panic!("expected XlsxError::PackageTooLarge, got {other:?}"),
    }
}
