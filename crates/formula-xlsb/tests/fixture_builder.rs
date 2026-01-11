#![allow(dead_code)]

use std::collections::BTreeMap;
use std::io::{Cursor, Write};

use formula_xlsb::biff12_varint;
use zip::write::FileOptions;
use zip::{CompressionMethod, ZipWriter};

/// Minimal XLSB fixture builder.
///
/// The goal is not to be a complete XLSB writer, but to generate *just enough*
/// OPC + BIFF12 to exercise our reader with targeted cell and formula payloads.
pub struct XlsbFixtureBuilder {
    sheet_name: String,
    shared_strings: Vec<String>,
    // row -> (col -> cell)
    cells: BTreeMap<u32, BTreeMap<u32, CellSpec>>,
}

#[derive(Debug, Clone)]
enum CellSpec {
    Number(f64),
    Sst(u32),
    FormulaNum { cached: f64, rgce: Vec<u8>, extra: Vec<u8> },
}

impl XlsbFixtureBuilder {
    pub fn new() -> Self {
        Self {
            sheet_name: "Sheet1".to_string(),
            shared_strings: Vec::new(),
            cells: BTreeMap::new(),
        }
    }

    pub fn add_shared_string(&mut self, s: &str) -> u32 {
        let idx = self.shared_strings.len() as u32;
        self.shared_strings.push(s.to_string());
        idx
    }

    pub fn set_sheet_name(&mut self, name: &str) {
        self.sheet_name = name.to_string();
    }

    pub fn set_cell_number(&mut self, row: u32, col: u32, v: f64) {
        self.cells
            .entry(row)
            .or_default()
            .insert(col, CellSpec::Number(v));
    }

    pub fn set_cell_sst(&mut self, row: u32, col: u32, sst_idx: u32) {
        self.cells
            .entry(row)
            .or_default()
            .insert(col, CellSpec::Sst(sst_idx));
    }

    pub fn set_cell_formula_num(&mut self, row: u32, col: u32, cached: f64, rgce: Vec<u8>, extra: Vec<u8>) {
        self.cells.entry(row).or_default().insert(
            col,
            CellSpec::FormulaNum {
                cached,
                rgce,
                extra,
            },
        );
    }

    /// Build a full `.xlsb` (ZIP) into memory.
    pub fn build_bytes(&self) -> Vec<u8> {
        let workbook_bin = build_workbook_bin(&self.sheet_name);
        let sheet1_bin = build_sheet_bin(&self.cells);
        let styles_bin = build_styles_bin();
        let shared_strings_bin = if self.shared_strings.is_empty() {
            None
        } else {
            Some(build_shared_strings_bin(&self.shared_strings))
        };

        let content_types_xml = build_content_types_xml(shared_strings_bin.is_some());
        let rels_xml = build_root_rels_xml();
        let workbook_rels_xml = build_workbook_rels_xml(shared_strings_bin.is_some());

        let mut zip = ZipWriter::new(Cursor::new(Vec::<u8>::new()));
        let options = FileOptions::default().compression_method(CompressionMethod::Stored);

        zip.start_file("[Content_Types].xml", options)
            .expect("start [Content_Types].xml");
        zip.write_all(content_types_xml.as_bytes())
            .expect("write [Content_Types].xml");

        zip.start_file("_rels/.rels", options)
            .expect("start _rels/.rels");
        zip.write_all(rels_xml.as_bytes()).expect("write _rels/.rels");

        zip.start_file("xl/workbook.bin", options)
            .expect("start xl/workbook.bin");
        zip.write_all(&workbook_bin)
            .expect("write xl/workbook.bin");

        zip.start_file("xl/_rels/workbook.bin.rels", options)
            .expect("start xl/_rels/workbook.bin.rels");
        zip.write_all(workbook_rels_xml.as_bytes())
            .expect("write xl/_rels/workbook.bin.rels");

        zip.start_file("xl/worksheets/sheet1.bin", options)
            .expect("start xl/worksheets/sheet1.bin");
        zip.write_all(&sheet1_bin)
            .expect("write xl/worksheets/sheet1.bin");

        if let Some(shared_strings_bin) = shared_strings_bin {
            zip.start_file("xl/sharedStrings.bin", options)
                .expect("start xl/sharedStrings.bin");
            zip.write_all(&shared_strings_bin)
                .expect("write xl/sharedStrings.bin");
        }

        zip.start_file("xl/styles.bin", options)
            .expect("start xl/styles.bin");
        zip.write_all(&styles_bin).expect("write xl/styles.bin");

        zip.finish().expect("finish xlsb zip").into_inner()
    }
}

// -- OPC (ZIP plumbing) --------------------------------------------------------

fn build_content_types_xml(has_shared_strings: bool) -> String {
    // Keep it close to Excel output (and our checked-in `simple.xlsb`) so the
    // package is also easy to debug with standard tools.
    let mut xml = r#"<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">
  <Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/>
  <Default Extension="xml" ContentType="application/xml"/>
  <Default Extension="bin" ContentType="application/vnd.ms-excel.sheet.binary.main"/>
  <Override PartName="/xl/workbook.bin" ContentType="application/vnd.ms-excel.sheet.binary.main"/>
  <Override PartName="/xl/worksheets/sheet1.bin" ContentType="application/vnd.ms-excel.worksheet"/>
"#
    .to_string();

    if has_shared_strings {
        xml.push_str(
            r#"  <Override PartName="/xl/sharedStrings.bin" ContentType="application/vnd.ms-excel.sharedStrings"/>
"#,
        );
    }

    xml.push_str(
        r#"  <Override PartName="/xl/styles.bin" ContentType="application/vnd.ms-excel.styles"/>
</Types>
"#,
    );
    xml
}

fn build_root_rels_xml() -> String {
    r#"<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
  <Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="xl/workbook.bin"/>
</Relationships>
"#
    .to_string()
}

fn build_workbook_rels_xml(has_shared_strings: bool) -> String {
    let mut xml = r#"<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
  <Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/worksheet" Target="worksheets/sheet1.bin"/>
"#
    .to_string();

    if has_shared_strings {
        xml.push_str(
            r#"  <Relationship Id="rId2" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/sharedStrings" Target="sharedStrings.bin"/>
  <Relationship Id="rId3" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/styles" Target="styles.bin"/>
"#,
        );
    } else {
        xml.push_str(
            r#"  <Relationship Id="rId2" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/styles" Target="styles.bin"/>
"#,
        );
    }

    xml.push_str("</Relationships>\n");
    xml
}

// -- BIFF12 writer -------------------------------------------------------------

// Record IDs copied from MS-XLSB / our parser constants. (Not re-exported from the crate.)
mod biff12 {
    pub const BEGIN_BOOK: u32 = 0x0183;
    pub const END_BOOK: u32 = 0x0184;
    pub const BEGIN_SHEETS: u32 = 0x018F;

    pub const SHEETS_END: u32 = 0x0190;
    pub const SHEET: u32 = 0x019C;

    pub const WORKSHEET: u32 = 0x0181;
    pub const WORKSHEET_END: u32 = 0x0182;
    pub const SHEETDATA: u32 = 0x0191;
    pub const SHEETDATA_END: u32 = 0x0192;
    pub const DIMENSION: u32 = 0x0194;

    pub const ROW: u32 = 0x0000;
    pub const FLOAT: u32 = 0x0005;
    pub const STRING: u32 = 0x0007;
    pub const FORMULA_FLOAT: u32 = 0x0009;

    pub const SST: u32 = 0x019F;
    pub const SST_END: u32 = 0x01A0;
    pub const SI: u32 = 0x0013;

    pub const BEGIN_STYLES: u32 = 0x0296;
    pub const END_STYLES: u32 = 0x0297;
}

fn build_workbook_bin(sheet_name: &str) -> Vec<u8> {
    let mut out = Vec::<u8>::new();

    write_record(&mut out, biff12::BEGIN_BOOK, &[]);
    write_record(&mut out, biff12::BEGIN_SHEETS, &[]);

    let mut sheet = Vec::<u8>::new();
    sheet.extend_from_slice(&0u32.to_le_bytes()); // flags/state (unused by our parser)
    sheet.extend_from_slice(&1u32.to_le_bytes()); // sheet id
    write_utf16_string(&mut sheet, "rId1");
    write_utf16_string(&mut sheet, sheet_name);
    write_record(&mut out, biff12::SHEET, &sheet);

    write_record(&mut out, biff12::SHEETS_END, &[]);
    write_record(&mut out, biff12::END_BOOK, &[]);

    out
}

fn build_sheet_bin(cells: &BTreeMap<u32, BTreeMap<u32, CellSpec>>) -> Vec<u8> {
    let mut out = Vec::<u8>::new();

    write_record(&mut out, biff12::WORKSHEET, &[]);

    let (r1, r2, c1, c2) = compute_dimension(cells);
    let mut dim = Vec::<u8>::new();
    dim.extend_from_slice(&r1.to_le_bytes());
    dim.extend_from_slice(&r2.to_le_bytes());
    dim.extend_from_slice(&c1.to_le_bytes());
    dim.extend_from_slice(&c2.to_le_bytes());
    write_record(&mut out, biff12::DIMENSION, &dim);

    write_record(&mut out, biff12::SHEETDATA, &[]);

    for (row, cols) in cells {
        write_record(&mut out, biff12::ROW, &row.to_le_bytes());
        for (col, cell) in cols {
            match cell {
                CellSpec::Number(v) => {
                    let mut data = Vec::<u8>::new();
                    data.extend_from_slice(&col.to_le_bytes());
                    data.extend_from_slice(&0u32.to_le_bytes()); // style
                    data.extend_from_slice(&v.to_le_bytes());
                    write_record(&mut out, biff12::FLOAT, &data);
                }
                CellSpec::Sst(idx) => {
                    let mut data = Vec::<u8>::new();
                    data.extend_from_slice(&col.to_le_bytes());
                    data.extend_from_slice(&0u32.to_le_bytes()); // style
                    data.extend_from_slice(&idx.to_le_bytes());
                    write_record(&mut out, biff12::STRING, &data);
                }
                CellSpec::FormulaNum { cached, rgce, extra } => {
                    let mut data = Vec::<u8>::new();
                    data.extend_from_slice(&col.to_le_bytes());
                    data.extend_from_slice(&0u32.to_le_bytes()); // style
                    data.extend_from_slice(&cached.to_le_bytes());
                    data.extend_from_slice(&0u16.to_le_bytes()); // flags
                    data.extend_from_slice(&(rgce.len() as u32).to_le_bytes());
                    data.extend_from_slice(rgce);
                    data.extend_from_slice(extra);
                    write_record(&mut out, biff12::FORMULA_FLOAT, &data);
                }
            }
        }
    }

    write_record(&mut out, biff12::SHEETDATA_END, &[]);
    write_record(&mut out, biff12::WORKSHEET_END, &[]);

    out
}

fn build_shared_strings_bin(strings: &[String]) -> Vec<u8> {
    let mut out = Vec::<u8>::new();

    let count = strings.len() as u32;
    let mut sst = Vec::<u8>::new();
    // uniqueCount, totalCount
    sst.extend_from_slice(&count.to_le_bytes());
    sst.extend_from_slice(&count.to_le_bytes());
    write_record(&mut out, biff12::SST, &sst);

    for s in strings {
        let mut si = Vec::<u8>::new();
        si.push(0u8); // flags (rich text / phonetic) - not used by our parser.
        write_utf16_string(&mut si, s);
        write_record(&mut out, biff12::SI, &si);
    }

    write_record(&mut out, biff12::SST_END, &[]);
    out
}

fn build_styles_bin() -> Vec<u8> {
    let mut out = Vec::<u8>::new();
    write_record(&mut out, biff12::BEGIN_STYLES, &[]);
    write_record(&mut out, biff12::END_STYLES, &[]);
    out
}

fn compute_dimension(cells: &BTreeMap<u32, BTreeMap<u32, CellSpec>>) -> (u32, u32, u32, u32) {
    let mut min_row: Option<u32> = None;
    let mut max_row: Option<u32> = None;
    let mut min_col: Option<u32> = None;
    let mut max_col: Option<u32> = None;

    for (&r, cols) in cells {
        min_row = Some(min_row.map_or(r, |v| v.min(r)));
        max_row = Some(max_row.map_or(r, |v| v.max(r)));
        for (&c, _) in cols {
            min_col = Some(min_col.map_or(c, |v| v.min(c)));
            max_col = Some(max_col.map_or(c, |v| v.max(c)));
        }
    }

    (
        min_row.unwrap_or(0),
        max_row.unwrap_or(0),
        min_col.unwrap_or(0),
        max_col.unwrap_or(0),
    )
}

fn write_record(out: &mut Vec<u8>, id: u32, data: &[u8]) {
    biff12_varint::write_record_id(out, id).expect("write record id");
    biff12_varint::write_record_len(out, data.len() as u32).expect("write record len");
    out.extend_from_slice(data);
}

fn write_utf16_string(out: &mut Vec<u8>, s: &str) {
    let units: Vec<u16> = s.encode_utf16().collect();
    out.extend_from_slice(&(units.len() as u32).to_le_bytes());
    for u in units {
        out.extend_from_slice(&u.to_le_bytes());
    }
}
