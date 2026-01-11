use std::collections::BTreeSet;
use std::fs::File;
use std::io::{Read, Write};
use std::path::{Path, PathBuf};

use formula_xlsb::{CellValue, XlsbWorkbook};
use pretty_assertions::assert_eq;
use zip::write::FileOptions;
use zip::{CompressionMethod, ZipArchive, ZipWriter};

fn fixture_path() -> PathBuf {
    Path::new(env!("CARGO_MANIFEST_DIR")).join("tests/fixtures/simple.xlsb")
}

fn format_report(report: &xlsx_diff::DiffReport) -> String {
    report
        .differences
        .iter()
        .map(|d| d.to_string())
        .collect::<Vec<_>>()
        .join("\n")
}

#[test]
fn save_as_is_lossless_at_opc_part_level() {
    let fixture_path = fixture_path();
    let wb = XlsbWorkbook::open(&fixture_path).expect("open xlsb fixture");

    let tmpdir = tempfile::tempdir().expect("create temp dir");
    let out_path = tmpdir.path().join("roundtrip.xlsb");
    wb.save_as(&out_path).expect("save_as");

    let report = xlsx_diff::diff_workbooks(&fixture_path, &out_path).expect("diff workbooks");
    assert!(
        report.is_empty(),
        "expected no OPC part diffs, got:\n{}",
        format_report(&report)
    );
}

trait SaveWithEditsExt {
    fn save_with_edits(
        &self,
        dest: impl AsRef<Path>,
        sheet_index: usize,
        row: u32,
        col: u32,
        value: f64,
    ) -> Result<(), formula_xlsb::Error>;
}

impl SaveWithEditsExt for XlsbWorkbook {
    fn save_with_edits(
        &self,
        dest: impl AsRef<Path>,
        sheet_index: usize,
        row: u32,
        col: u32,
        value: f64,
    ) -> Result<(), formula_xlsb::Error> {
        let dest = dest.as_ref();
        self.save_as(dest)?;

        let sheet_part = self
            .sheet_metas()
            .get(sheet_index)
            .ok_or(formula_xlsb::Error::SheetIndexOutOfBounds(sheet_index))?
            .part_path
            .clone();

        let in_file = File::open(dest)?;
        let mut zip = ZipArchive::new(in_file)?;

        let tmpdir = tempfile::tempdir()?;
        let tmp_path = tmpdir.path().join("patched.xlsb");
        let out_file = File::create(&tmp_path)?;
        let mut writer = ZipWriter::new(out_file);

        let options = FileOptions::default().compression_method(CompressionMethod::Deflated);

        for i in 0..zip.len() {
            let mut entry = zip.by_index(i)?;
            let name = entry.name().to_string();

            if entry.is_dir() {
                writer.add_directory(name, options)?;
                continue;
            }

            let mut bytes = Vec::with_capacity(entry.size() as usize);
            entry.read_to_end(&mut bytes)?;

            writer.start_file(&name, options)?;
            if name == sheet_part {
                let patched = patch_sheet_numeric_cell(&bytes, row, col, value)?;
                writer.write_all(&patched)?;
            } else {
                writer.write_all(&bytes)?;
            }
        }

        writer.finish()?;
        std::fs::remove_file(dest)?;
        std::fs::rename(tmp_path, dest)?;

        Ok(())
    }
}

#[test]
fn patch_writer_changes_only_target_sheet_part() {
    let fixture_path = fixture_path();
    let wb = XlsbWorkbook::open(&fixture_path).expect("open xlsb fixture");

    let tmpdir = tempfile::tempdir().expect("create temp dir");
    let out_path = tmpdir.path().join("patched.xlsb");
    wb.save_with_edits(&out_path, 0, 0, 1, 123.0)
        .expect("save_with_edits");

    let patched = XlsbWorkbook::open(&out_path).expect("re-open patched workbook");
    let sheet = patched.read_sheet(0).expect("read patched sheet");
    let b1 = sheet
        .cells
        .iter()
        .find(|c| c.row == 0 && c.col == 1)
        .expect("B1 exists");
    assert_eq!(b1.value, CellValue::Number(123.0));

    let report = xlsx_diff::diff_workbooks(&fixture_path, &out_path).expect("diff workbooks");
    assert!(
        report
            .differences
            .iter()
            .any(|d| d.part == "xl/worksheets/sheet1.bin"),
        "expected worksheet part to change, got:\n{}",
        format_report(&report)
    );

    let unexpected_missing: Vec<_> = report
        .differences
        .iter()
        .filter(|d| d.kind == "missing_part" && !is_calc_chain_part(&d.part))
        .map(|d| d.part.clone())
        .collect();
    assert!(
        unexpected_missing.is_empty(),
        "unexpected missing parts: {unexpected_missing:?}\n{}",
        format_report(&report)
    );

    let parts: BTreeSet<String> = report.differences.iter().map(|d| d.part.clone()).collect();
    let unexpected_parts: Vec<_> = parts
        .iter()
        .filter(|part| !is_allowed_patch_diff_part(part))
        .cloned()
        .collect();

    assert!(
        unexpected_parts.is_empty(),
        "unexpected diff parts: {unexpected_parts:?}\n{}",
        format_report(&report)
    );
}

fn is_allowed_patch_diff_part(part: &str) -> bool {
    part == "xl/worksheets/sheet1.bin" || is_calc_chain_part(part)
}

fn is_calc_chain_part(part: &str) -> bool {
    part.starts_with("xl/calcChain.")
}

fn patch_sheet_numeric_cell(
    sheet_bytes: &[u8],
    target_row: u32,
    target_col: u32,
    value: f64,
) -> Result<Vec<u8>, formula_xlsb::Error> {
    const SHEETDATA: u32 = 0x0191;
    const SHEETDATA_END: u32 = 0x0192;
    const ROW: u32 = 0x0000;
    const NUM: u32 = 0x0002;
    const FLOAT: u32 = 0x0005;

    let mut out = sheet_bytes.to_vec();
    let mut offset = 0usize;
    let mut in_sheet_data = false;
    let mut current_row = 0u32;
    let mut patched = false;

    while offset < sheet_bytes.len() {
        let record_id = read_biff12_id(sheet_bytes, &mut offset)?;
        let len = read_biff12_len(sheet_bytes, &mut offset)? as usize;

        let data_start = offset;
        let data_end = offset
            .checked_add(len)
            .filter(|end| *end <= sheet_bytes.len())
            .ok_or_else(|| {
                formula_xlsb::Error::Io(std::io::Error::new(
                    std::io::ErrorKind::UnexpectedEof,
                    "invalid record length",
                ))
            })?;
        let data = &sheet_bytes[data_start..data_end];

        match record_id {
            SHEETDATA => in_sheet_data = true,
            SHEETDATA_END => break,
            ROW if in_sheet_data => {
                if data.len() < 4 {
                    return Err(formula_xlsb::Error::Io(std::io::Error::new(
                        std::io::ErrorKind::UnexpectedEof,
                        "row record too short",
                    )));
                }
                current_row = u32::from_le_bytes(data[0..4].try_into().unwrap());
            }
            NUM | FLOAT if in_sheet_data => {
                if data.len() < 8 {
                    return Err(formula_xlsb::Error::Io(std::io::Error::new(
                        std::io::ErrorKind::UnexpectedEof,
                        "cell record too short",
                    )));
                }
                let col = u32::from_le_bytes(data[0..4].try_into().unwrap());
                if current_row == target_row && col == target_col {
                    let value_offset = data_start + 8;
                    match record_id {
                        NUM => {
                            if data.len() < 12 {
                                return Err(formula_xlsb::Error::Io(std::io::Error::new(
                                    std::io::ErrorKind::UnexpectedEof,
                                    "rk cell record too short",
                                )));
                            }
                            let rk = encode_rk_number(value).ok_or_else(|| {
                                formula_xlsb::Error::Io(std::io::Error::new(
                                    std::io::ErrorKind::InvalidInput,
                                    "value not representable as RK",
                                ))
                            })?;
                            out[value_offset..value_offset + 4].copy_from_slice(&rk.to_le_bytes());
                        }
                        FLOAT => {
                            if data.len() < 16 {
                                return Err(formula_xlsb::Error::Io(std::io::Error::new(
                                    std::io::ErrorKind::UnexpectedEof,
                                    "float cell record too short",
                                )));
                            }
                            out[value_offset..value_offset + 8]
                                .copy_from_slice(&value.to_le_bytes());
                        }
                        _ => {}
                    }
                    patched = true;
                    break;
                }
            }
            _ => {}
        }

        offset = data_end;
    }

    if !patched {
        return Err(formula_xlsb::Error::Io(std::io::Error::new(
            std::io::ErrorKind::NotFound,
            format!("cell not found at row {target_row} col {target_col}"),
        )));
    }

    Ok(out)
}

fn read_biff12_id(bytes: &[u8], offset: &mut usize) -> Result<u32, formula_xlsb::Error> {
    let mut v: u32 = 0;
    for i in 0..4 {
        let b = *bytes.get(*offset).ok_or_else(|| {
            formula_xlsb::Error::Io(std::io::Error::new(
                std::io::ErrorKind::UnexpectedEof,
                "unexpected eof while reading record id",
            ))
        })?;
        *offset += 1;
        v |= (b as u32) << (8 * i);
        if b & 0x80 == 0 {
            break;
        }
    }
    Ok(v)
}

fn read_biff12_len(bytes: &[u8], offset: &mut usize) -> Result<u32, formula_xlsb::Error> {
    let mut v: u32 = 0;
    for i in 0..4 {
        let b = *bytes.get(*offset).ok_or_else(|| {
            formula_xlsb::Error::Io(std::io::Error::new(
                std::io::ErrorKind::UnexpectedEof,
                "unexpected eof while reading record length",
            ))
        })?;
        *offset += 1;
        v |= ((b & 0x7F) as u32) << (7 * i);
        if b & 0x80 == 0 {
            break;
        }
    }
    Ok(v)
}

fn encode_rk_number(value: f64) -> Option<u32> {
    if !value.is_finite() {
        return None;
    }

    let int = value.round();
    if (value - int).abs() <= f64::EPSILON && int >= i32::MIN as f64 && int <= i32::MAX as f64 {
        let i = int as i32;
        return Some(((i as u32) << 2) | 0x02);
    }

    let scaled = (value * 100.0).round();
    if ((value * 100.0) - scaled).abs() <= 1e-6
        && scaled >= i32::MIN as f64
        && scaled <= i32::MAX as f64
    {
        let i = scaled as i32;
        return Some(((i as u32) << 2) | 0x03);
    }

    None
}
