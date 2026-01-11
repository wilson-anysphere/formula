use std::io::{Cursor, Read};
use std::path::PathBuf;
use std::time::{SystemTime, UNIX_EPOCH};

use calamine::{open_workbook, Reader, Xls};

fn read_workbook_stream_from_xls_bytes(data: &[u8]) -> Vec<u8> {
    let cursor = Cursor::new(data.to_vec());
    let mut ole = cfb::CompoundFile::open(cursor).expect("open xls cfb");

    for candidate in ["/Workbook", "/Book", "Workbook", "Book"] {
        if let Ok(mut stream) = ole.open_stream(candidate) {
            let mut buf = Vec::new();
            stream
                .read_to_end(&mut buf)
                .expect("read workbook stream");
            return buf;
        }
    }

    panic!("fixture missing Workbook/Book stream");
}

fn cfb_sector_size(data: &[u8]) -> usize {
    let header = data.get(0..0x20).expect("cfb header");
    let shift = u16::from_le_bytes([header[0x1E], header[0x1F]]) as u32;
    1usize << shift
}

fn find_all(haystack: &[u8], needle: &[u8]) -> Vec<usize> {
    if needle.is_empty() || needle.len() > haystack.len() {
        return Vec::new();
    }

    haystack
        .windows(needle.len())
        .enumerate()
        .filter_map(|(idx, window)| (window == needle).then_some(idx))
        .collect()
}

fn read_biff_record_header(buf: &[u8], offset: usize) -> Option<(u16, usize)> {
    let header = buf.get(offset..offset + 4)?;
    let record_id = u16::from_le_bytes([header[0], header[1]]);
    let len = u16::from_le_bytes([header[2], header[3]]) as usize;
    Some((record_id, len))
}

fn detect_biff8(buf: &[u8]) -> bool {
    let Some((record_id, len)) = read_biff_record_header(buf, 0) else {
        return true;
    };
    if record_id != 0x0809 && record_id != 0x0009 {
        return true;
    }
    let data = buf.get(4..4 + len).unwrap_or(&[]);
    let Some(v) = data.get(0..2).map(|v| u16::from_le_bytes([v[0], v[1]])) else {
        return true;
    };
    v == 0x0600
}

/// Patch the first BoundSheet name in-place so it contains an embedded NUL byte.
///
/// Returns the sheet name bytes interpreted via `String::from_utf8_lossy`
/// (mirrors the legacy BIFF name decoding used by the importer before this fix).
fn patch_first_boundsheet_name(workbook_stream: &mut [u8]) -> (String, usize) {
    let biff8 = detect_biff8(workbook_stream);

    let mut offset = 0usize;
    loop {
        let Some((record_id, len)) = read_biff_record_header(workbook_stream, offset) else {
            break;
        };
        let data_start = offset + 4;
        let data_end = data_start + len;

        if record_id == 0x0085 {
            let data = workbook_stream
                .get_mut(data_start..data_end)
                .expect("boundsheet data in range");

            // [MS-XLS] BoundSheet8: 4 bytes BOF offset, 1 byte state, 1 byte type, then a string.
            assert!(
                data.len() >= 8,
                "expected BoundSheet record to contain at least offset/state/type/string header"
            );

            if biff8 {
                let string_data = &mut data[6..];
                let cch = string_data[0] as usize;
                let flags = string_data[1];
                let mut string_offset = 2usize;

                if flags & 0x08 != 0 {
                    // Skip rich-text run count (2 bytes)
                    string_offset += 2;
                }
                if flags & 0x04 != 0 {
                    // Skip extended string size (4 bytes)
                    string_offset += 4;
                }

                // We specifically want the compressed 8-bit path so that calamine's
                // codepage decoding can diverge from our naive UTF-8 decoding.
                assert_eq!(
                    flags & 0x01,
                    0,
                    "expected fixture sheet name to be stored in compressed 8-bit form"
                );

                let char_start = string_offset;
                let char_end = char_start + cch;
                assert!(
                    string_data.len() >= char_end,
                    "unexpected end of BIFF8 sheet name"
                );
                assert!(cch >= 3, "expected sheet name to be at least 3 chars");
                string_data[char_start + 2] = 0x00;

                let patched_name =
                    String::from_utf8_lossy(&string_data[char_start..char_end]).into_owned();
                let patched_offset = data_start + 6 + char_start + 2;
                return (patched_name, patched_offset);
            }

            // BIFF5 short string: 1 byte length then raw bytes.
            let string_data = &mut data[6..];
            let cch = string_data[0] as usize;
            assert!(
                string_data.len() >= 1 + cch,
                "unexpected end of BIFF5 sheet name"
            );
            assert!(cch >= 3, "expected sheet name to be at least 3 chars");
            string_data[1 + 2] = 0x00;
            let patched_name = String::from_utf8_lossy(&string_data[1..1 + cch]).into_owned();
            let patched_offset = data_start + 6 + 1 + 2;
            return (patched_name, patched_offset);
        }

        offset = data_end;
    }

    panic!("no BoundSheet record found to patch");
}

#[test]
fn imports_row_col_properties_even_when_biff_sheet_name_is_sanitized() {
    let fixture_path = PathBuf::from(env!("CARGO_MANIFEST_DIR"))
        .join("tests")
        .join("fixtures")
        .join("merged_hidden.xls");
    let fixture_bytes = std::fs::read(&fixture_path).expect("read fixture");

    let workbook_stream = read_workbook_stream_from_xls_bytes(&fixture_bytes);
    let mut patched_workbook_stream = workbook_stream.clone();
    let (biff_utf8_lossy_name, patch_offset) = patch_first_boundsheet_name(&mut patched_workbook_stream);

    let sector_size = cfb_sector_size(&fixture_bytes);
    let sector_base = patch_offset - (patch_offset % sector_size);
    let window_start = patch_offset.saturating_sub(32).max(sector_base);
    let window_end = (window_start + 64)
        .min(sector_base + sector_size)
        .min(workbook_stream.len());
    let patch_offset_in_window = patch_offset - window_start;
    let window = &workbook_stream[window_start..window_end];

    let matches = find_all(&fixture_bytes, window);
    assert_eq!(
        matches.len(),
        1,
        "expected patched workbook stream window to appear exactly once in the fixture CFB file"
    );
    let mut patched_file = fixture_bytes.clone();
    let file_offset = matches[0] + patch_offset_in_window;
    assert_eq!(
        patched_file[file_offset],
        workbook_stream[patch_offset],
        "expected located CFB window to contain original workbook byte"
    );
    patched_file[file_offset] = 0x00;

    let tmp_path = std::env::temp_dir().join(format!(
        "formula_xls_sheet_name_mismatch_{}_{}.xls",
        std::process::id(),
        SystemTime::now()
            .duration_since(UNIX_EPOCH)
            .expect("time")
            .as_nanos()
    ));
    std::fs::write(&tmp_path, &patched_file).expect("write patched xls");

    let calamine: Xls<_> = open_workbook(&tmp_path).expect("open patched xls");
    let calamine_sheets = calamine.sheets_metadata().to_vec();
    assert_eq!(calamine_sheets.len(), 1);
    let calamine_name = calamine_sheets[0].name.clone();
    drop(calamine);

    // The BIFF BoundSheet name is stored in 8-bit encoding; our importer used to decode
    // it as UTF-8-lossy, which can diverge from calamine's decoding/sanitization.
    assert!(
        biff_utf8_lossy_name.contains('\0'),
        "expected patched BIFF sheet name to contain an embedded NUL"
    );
    assert_ne!(
        calamine_name, biff_utf8_lossy_name,
        "expected calamine and naive BIFF UTF-8 decoding to diverge for the patched name"
    );

    let result = formula_xls::import_xls_path(&tmp_path).expect("import xls");
    std::fs::remove_file(&tmp_path).ok();

    let sheet = &result.workbook.sheets[0];

    // Row and column properties should still be applied to the correct sheet
    // even if the BIFF BoundSheet name differs from calamine's sheet name.
    assert_eq!(sheet.row_properties(0).unwrap().height, Some(20.0));
    assert!(sheet.row_properties(2).unwrap().hidden);

    assert_eq!(sheet.col_properties(0).unwrap().width, Some(20.0));
    assert!(sheet.col_properties(3).unwrap().hidden);
}
