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
    shared_strings_bin_override: Option<Vec<u8>>,
    // row -> (col -> cell)
    cells: BTreeMap<u32, BTreeMap<u32, CellSpec>>,
    row_record_trailing_bytes: Vec<u8>,
    extra_zip_parts: Vec<(String, Vec<u8>)>,
}

#[derive(Debug, Clone)]
enum CellSpec {
    Blank,
    Bool(bool),
    Error(u8),
    Number(f64),
    Rk(u32),
    Sst(u32),
    InlineString(String),
    FormulaNum {
        cached: f64,
        rgce: Vec<u8>,
        extra: Vec<u8>,
    },
    FormulaStr {
        cached: String,
        rgce: Vec<u8>,
    },
    FormulaBool {
        cached: bool,
        rgce: Vec<u8>,
    },
    FormulaErr {
        cached: u8,
        rgce: Vec<u8>,
    },
}

impl XlsbFixtureBuilder {
    pub fn new() -> Self {
        Self {
            sheet_name: "Sheet1".to_string(),
            shared_strings: Vec::new(),
            shared_strings_bin_override: None,
            cells: BTreeMap::new(),
            row_record_trailing_bytes: Vec::new(),
            extra_zip_parts: Vec::new(),
        }
    }

    pub fn add_shared_string(&mut self, s: &str) -> u32 {
        let idx = self.shared_strings.len() as u32;
        self.shared_strings.push(s.to_string());
        idx
    }

    /// Override the generated `xl/sharedStrings.bin` part bytes.
    pub fn set_shared_strings_bin_override(&mut self, bytes: Vec<u8>) {
        self.shared_strings_bin_override = Some(bytes);
    }

    pub fn set_sheet_name(&mut self, name: &str) {
        self.sheet_name = name.to_string();
    }

    pub fn set_row_record_trailing_bytes(&mut self, bytes: Vec<u8>) {
        self.row_record_trailing_bytes = bytes;
    }

    /// Add an arbitrary extra part to the generated XLSB ZIP package.
    ///
    /// This is useful for testing consumers that read auxiliary XML parts (e.g. table definitions).
    pub fn add_extra_zip_part(&mut self, name: impl Into<String>, bytes: Vec<u8>) {
        self.extra_zip_parts.push((name.into(), bytes));
    }

    pub fn set_cell_number(&mut self, row: u32, col: u32, v: f64) {
        self.cells
            .entry(row)
            .or_default()
            .insert(col, CellSpec::Number(v));
    }

    pub fn set_cell_blank(&mut self, row: u32, col: u32) {
        self.cells
            .entry(row)
            .or_default()
            .insert(col, CellSpec::Blank);
    }

    pub fn set_cell_bool(&mut self, row: u32, col: u32, v: bool) {
        self.cells
            .entry(row)
            .or_default()
            .insert(col, CellSpec::Bool(v));
    }

    pub fn set_cell_error(&mut self, row: u32, col: u32, code: u8) {
        self.cells
            .entry(row)
            .or_default()
            .insert(col, CellSpec::Error(code));
    }

    pub fn set_cell_number_rk(&mut self, row: u32, col: u32, v: f64) {
        let rk = encode_rk_number(v).unwrap_or_else(|| panic!("value {v} not representable as RK"));
        self.cells
            .entry(row)
            .or_default()
            .insert(col, CellSpec::Rk(rk));
    }

    pub fn set_cell_sst(&mut self, row: u32, col: u32, sst_idx: u32) {
        self.cells
            .entry(row)
            .or_default()
            .insert(col, CellSpec::Sst(sst_idx));
    }

    pub fn set_cell_inline_string(&mut self, row: u32, col: u32, s: &str) {
        self.cells
            .entry(row)
            .or_default()
            .insert(col, CellSpec::InlineString(s.to_string()));
    }

    pub fn set_cell_formula_num(
        &mut self,
        row: u32,
        col: u32,
        cached: f64,
        rgce: Vec<u8>,
        extra: Vec<u8>,
    ) {
        self.cells.entry(row).or_default().insert(
            col,
            CellSpec::FormulaNum {
                cached,
                rgce,
                extra,
            },
        );
    }

    pub fn set_cell_formula_str(
        &mut self,
        row: u32,
        col: u32,
        cached: impl Into<String>,
        rgce: Vec<u8>,
    ) {
        self.cells.entry(row).or_default().insert(
            col,
            CellSpec::FormulaStr {
                cached: cached.into(),
                rgce,
            },
        );
    }

    pub fn set_cell_formula_bool(&mut self, row: u32, col: u32, cached: bool, rgce: Vec<u8>) {
        self.cells
            .entry(row)
            .or_default()
            .insert(col, CellSpec::FormulaBool { cached, rgce });
    }

    pub fn set_cell_formula_err(&mut self, row: u32, col: u32, cached: u8, rgce: Vec<u8>) {
        self.cells
            .entry(row)
            .or_default()
            .insert(col, CellSpec::FormulaErr { cached, rgce });
    }

    /// Build a full `.xlsb` (ZIP) into memory.
    pub fn build_bytes(&self) -> Vec<u8> {
        let workbook_bin = build_workbook_bin(&self.sheet_name);
        let sheet1_bin = build_sheet_bin(&self.cells, &self.row_record_trailing_bytes);
        let shared_strings_bin = if let Some(bytes) = &self.shared_strings_bin_override {
            Some(bytes.clone())
        } else if self.shared_strings.is_empty() {
            None
        } else {
            Some(build_shared_strings_bin(&self.shared_strings, &self.cells))
        };

        let content_types_xml = build_content_types_xml(shared_strings_bin.is_some());
        let rels_xml = build_root_rels_xml();
        let workbook_rels_xml = build_workbook_rels_xml(shared_strings_bin.is_some());

        let mut zip = ZipWriter::new(Cursor::new(Vec::<u8>::new()));
        let options = FileOptions::<()>::default().compression_method(CompressionMethod::Stored);

        zip.start_file("[Content_Types].xml", options.clone())
            .expect("start [Content_Types].xml");
        zip.write_all(content_types_xml.as_bytes())
            .expect("write [Content_Types].xml");

        zip.start_file("_rels/.rels", options.clone())
            .expect("start _rels/.rels");
        zip.write_all(rels_xml.as_bytes())
            .expect("write _rels/.rels");

        zip.start_file("xl/workbook.bin", options.clone())
            .expect("start xl/workbook.bin");
        zip.write_all(&workbook_bin).expect("write xl/workbook.bin");

        zip.start_file("xl/_rels/workbook.bin.rels", options.clone())
            .expect("start xl/_rels/workbook.bin.rels");
        zip.write_all(workbook_rels_xml.as_bytes())
            .expect("write xl/_rels/workbook.bin.rels");

        zip.start_file("xl/worksheets/sheet1.bin", options.clone())
            .expect("start xl/worksheets/sheet1.bin");
        zip.write_all(&sheet1_bin)
            .expect("write xl/worksheets/sheet1.bin");

        if let Some(shared_strings_bin) = shared_strings_bin {
            zip.start_file("xl/sharedStrings.bin", options.clone())
                .expect("start xl/sharedStrings.bin");
            zip.write_all(&shared_strings_bin)
                .expect("write xl/sharedStrings.bin");
        }

        for (name, bytes) in &self.extra_zip_parts {
            zip.start_file(name, options.clone())
                .unwrap_or_else(|_| panic!("start {name}"));
            zip.write_all(bytes)
                .unwrap_or_else(|_| panic!("write {name}"));
        }

        zip.finish().expect("finish xlsb zip").into_inner()
    }
}

/// Helpers for building small `rgce` token streams in tests.
///
/// These helpers operate on the *zero-based* row/col indices used by XLSB.
pub mod rgce {
    /// PtgRef (`0x24`) encoded using the BIFF12 cell-address layout:
    /// `[row: u32][col: u16-with-relative-flags]`.
    pub fn push_ref(out: &mut Vec<u8>, row: u32, col: u32, abs_row: bool, abs_col: bool) {
        assert!(
            col <= 0x3FFF,
            "column index out of range for BIFF12 (max 16383)"
        );
        out.push(0x24);
        out.extend_from_slice(&row.to_le_bytes());

        let col_u16 = col as u16;
        let low = (col_u16 & 0x00FF) as u8;
        let mut high = ((col_u16 >> 8) as u8) & 0x3F;

        // Our decoder interprets bit=1 as "relative" (no `$`).
        if !abs_row {
            high |= 0x40;
        }
        if !abs_col {
            high |= 0x80;
        }

        out.push(low);
        out.push(high);
    }

    /// PtgInt (`0x1E`) literal.
    pub fn push_int(out: &mut Vec<u8>, n: u16) {
        out.push(0x1E);
        out.extend_from_slice(&n.to_le_bytes());
    }

    /// PtgNum (`0x1F`) literal.
    pub fn push_num(out: &mut Vec<u8>, n: f64) {
        out.push(0x1F);
        out.extend_from_slice(&n.to_le_bytes());
    }

    /// PtgAdd (`0x03`)
    pub fn push_add(out: &mut Vec<u8>) {
        out.push(0x03);
    }

    /// PtgMul (`0x05`)
    pub fn push_mul(out: &mut Vec<u8>) {
        out.push(0x05);
    }

    /// PtgParen (`0x15`)
    pub fn push_paren(out: &mut Vec<u8>) {
        out.push(0x15);
    }

    /// PtgUminus (`0x13`)
    pub fn push_unary_minus(out: &mut Vec<u8>) {
        out.push(0x13);
    }

    /// PtgArray (`0x20`) placeholder token.
    ///
    /// In real files this is typically followed by an `rgcb` payload after the `rgce` stream.
    pub fn array_placeholder() -> Vec<u8> {
        vec![0x20, 0, 0, 0, 0, 0, 0, 0]
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

    xml.push_str("</Types>\n");
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
"#,
        );
    }

    xml.push_str("</Relationships>\n");
    xml
}

// -- BIFF12 writer -------------------------------------------------------------

// Record IDs copied from MS-XLSB / our parser constants. (Not re-exported from the crate.)
mod biff12 {
    pub const BEGIN_BOOK: u32 = 0x0083;
    pub const END_BOOK: u32 = 0x0084;
    pub const BEGIN_SHEETS: u32 = 0x008F;

    pub const SHEETS_END: u32 = 0x0090;
    pub const SHEET: u32 = 0x009C;

    pub const WORKSHEET: u32 = 0x0081;
    pub const WORKSHEET_END: u32 = 0x0082;
    pub const SHEETDATA: u32 = 0x0091;
    pub const SHEETDATA_END: u32 = 0x0092;
    pub const DIMENSION: u32 = 0x0094;

    pub const ROW: u32 = 0x0000;
    pub const BLANK: u32 = 0x0001;
    pub const NUM: u32 = 0x0002;
    pub const BOOLERR: u32 = 0x0003;
    pub const BOOL: u32 = 0x0004;
    pub const FLOAT: u32 = 0x0005;
    pub const STRING: u32 = 0x0007;
    pub const CELL_ST: u32 = 0x0006;
    pub const FORMULA_STRING: u32 = 0x0008;
    pub const FORMULA_FLOAT: u32 = 0x0009;
    pub const FORMULA_BOOL: u32 = 0x000A;
    pub const FORMULA_BOOLERR: u32 = 0x000B;

    pub const SST: u32 = 0x009F;
    pub const SST_END: u32 = 0x00A0;
    pub const SI: u32 = 0x0013;

    // Stylesheet records are intentionally omitted from fixtures. `formula-xlsb` does not parse
    // them, and Calamine (used in our formula-comparison regression tests) currently fails to
    // open minimal XLSB packages that include `xl/styles.bin`.
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

fn build_sheet_bin(
    cells: &BTreeMap<u32, BTreeMap<u32, CellSpec>>,
    row_record_trailing_bytes: &[u8],
) -> Vec<u8> {
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
        if row_record_trailing_bytes.is_empty() {
            write_record(&mut out, biff12::ROW, &row.to_le_bytes());
        } else {
            let mut payload = Vec::with_capacity(4 + row_record_trailing_bytes.len());
            payload.extend_from_slice(&row.to_le_bytes());
            payload.extend_from_slice(row_record_trailing_bytes);
            write_record(&mut out, biff12::ROW, &payload);
        }
        for (col, cell) in cols {
            match cell {
                CellSpec::Blank => {
                    let mut data = Vec::<u8>::new();
                    data.extend_from_slice(&col.to_le_bytes());
                    data.extend_from_slice(&0u32.to_le_bytes()); // style
                    write_record(&mut out, biff12::BLANK, &data);
                }
                CellSpec::Bool(v) => {
                    let mut data = Vec::<u8>::new();
                    data.extend_from_slice(&col.to_le_bytes());
                    data.extend_from_slice(&0u32.to_le_bytes()); // style
                    data.push(u8::from(*v));
                    write_record(&mut out, biff12::BOOL, &data);
                }
                CellSpec::Error(code) => {
                    let mut data = Vec::<u8>::new();
                    data.extend_from_slice(&col.to_le_bytes());
                    data.extend_from_slice(&0u32.to_le_bytes()); // style
                    data.push(*code);
                    write_record(&mut out, biff12::BOOLERR, &data);
                }
                CellSpec::Number(v) => {
                    let mut data = Vec::<u8>::new();
                    data.extend_from_slice(&col.to_le_bytes());
                    data.extend_from_slice(&0u32.to_le_bytes()); // style
                    data.extend_from_slice(&v.to_le_bytes());
                    write_record(&mut out, biff12::FLOAT, &data);
                }
                CellSpec::Rk(rk) => {
                    let mut data = Vec::<u8>::new();
                    data.extend_from_slice(&col.to_le_bytes());
                    data.extend_from_slice(&0u32.to_le_bytes()); // style
                    data.extend_from_slice(&rk.to_le_bytes());
                    write_record(&mut out, biff12::NUM, &data);
                }
                CellSpec::Sst(idx) => {
                    let mut data = Vec::<u8>::new();
                    data.extend_from_slice(&col.to_le_bytes());
                    data.extend_from_slice(&0u32.to_le_bytes()); // style
                    data.extend_from_slice(&idx.to_le_bytes());
                    write_record(&mut out, biff12::STRING, &data);
                }
                CellSpec::InlineString(s) => {
                    let mut data = Vec::<u8>::new();
                    data.extend_from_slice(&col.to_le_bytes());
                    data.extend_from_slice(&0u32.to_le_bytes()); // style
                    write_utf16_string(&mut data, s);
                    write_record(&mut out, biff12::CELL_ST, &data);
                }
                CellSpec::FormulaNum {
                    cached,
                    rgce,
                    extra,
                } => {
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
                CellSpec::FormulaStr { cached, rgce } => {
                    let mut data = Vec::<u8>::new();
                    data.extend_from_slice(&col.to_le_bytes());
                    data.extend_from_slice(&0u32.to_le_bytes()); // style
                    let units: Vec<u16> = cached.encode_utf16().collect();
                    data.extend_from_slice(&(units.len() as u32).to_le_bytes()); // cch
                    data.extend_from_slice(&0u16.to_le_bytes()); // flags
                    for u in units {
                        data.extend_from_slice(&u.to_le_bytes());
                    }
                    data.extend_from_slice(&(rgce.len() as u32).to_le_bytes());
                    data.extend_from_slice(rgce);
                    write_record(&mut out, biff12::FORMULA_STRING, &data);
                }
                CellSpec::FormulaBool { cached, rgce } => {
                    let mut data = Vec::<u8>::new();
                    data.extend_from_slice(&col.to_le_bytes());
                    data.extend_from_slice(&0u32.to_le_bytes()); // style
                    data.push(if *cached { 1 } else { 0 });
                    data.extend_from_slice(&0u16.to_le_bytes()); // flags
                    data.extend_from_slice(&(rgce.len() as u32).to_le_bytes());
                    data.extend_from_slice(rgce);
                    write_record(&mut out, biff12::FORMULA_BOOL, &data);
                }
                CellSpec::FormulaErr { cached, rgce } => {
                    let mut data = Vec::<u8>::new();
                    data.extend_from_slice(&col.to_le_bytes());
                    data.extend_from_slice(&0u32.to_le_bytes()); // style
                    data.push(*cached);
                    data.extend_from_slice(&0u16.to_le_bytes()); // flags
                    data.extend_from_slice(&(rgce.len() as u32).to_le_bytes());
                    data.extend_from_slice(rgce);
                    write_record(&mut out, biff12::FORMULA_BOOLERR, &data);
                }
            }
        }
    }

    write_record(&mut out, biff12::SHEETDATA_END, &[]);
    write_record(&mut out, biff12::WORKSHEET_END, &[]);

    out
}

fn build_shared_strings_bin(
    strings: &[String],
    cells: &BTreeMap<u32, BTreeMap<u32, CellSpec>>,
) -> Vec<u8> {
    let mut out = Vec::<u8>::new();

    let unique_count = strings.len() as u32;
    let total_count: u32 = cells
        .values()
        .map(|cols| {
            cols.values()
                .filter(|cell| matches!(cell, CellSpec::Sst(_)))
                .count() as u32
        })
        .sum();
    let mut sst = Vec::<u8>::new();
    // BrtSST: [totalCount][uniqueCount]
    sst.extend_from_slice(&total_count.to_le_bytes());
    sst.extend_from_slice(&unique_count.to_le_bytes());
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
    let len = u32::try_from(data.len()).expect("record too large");
    biff12_varint::write_record_len(out, len).expect("write record len");
    out.extend_from_slice(data);
}

fn write_utf16_string(out: &mut Vec<u8>, s: &str) {
    let units: Vec<u16> = s.encode_utf16().collect();
    let len = u32::try_from(units.len()).expect("string too large");
    out.extend_from_slice(&len.to_le_bytes());
    for u in units {
        out.extend_from_slice(&u.to_le_bytes());
    }
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

    // Non-integer RK numbers store the top 30 bits of the IEEE754 f64 (with the low
    // 34 bits cleared) and set the low two bits to 0b00 (or 0b01 when scaled by 100).
    const LOW_34_MASK: u64 = (1u64 << 34) - 1;
    let bits = value.to_bits();
    if bits & LOW_34_MASK == 0 {
        let raw = (bits >> 32) as u32;
        if raw & 0x03 == 0 {
            return Some(raw);
        }
    }

    let scaled = value * 100.0;
    if scaled.is_finite() {
        let bits = scaled.to_bits();
        if bits & LOW_34_MASK == 0 {
            let raw = (bits >> 32) as u32;
            if raw & 0x03 == 0 {
                return Some(raw | 0x01);
            }
        }
    }

    None
}
