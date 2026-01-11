use crate::tables::{write_table_xml, TABLE_REL_TYPE};
use crate::styles::StylesPart;
use formula_model::{normalize_formula_text, Cell, CellRef, CellValue, Workbook, Worksheet};
use std::collections::{BTreeMap, HashMap};
use std::fs::File;
use std::io::{Seek, Write};
use std::path::Path;
use thiserror::Error;
use zip::ZipWriter;

#[derive(Debug, Error)]
pub enum XlsxWriteError {
    #[error("io error: {0}")]
    Io(#[from] std::io::Error),
    #[error("zip error: {0}")]
    Zip(#[from] zip::result::ZipError),
    #[error("invalid workbook: {0}")]
    Invalid(String),
}

pub fn write_workbook(workbook: &Workbook, path: impl AsRef<Path>) -> Result<(), XlsxWriteError> {
    let file = File::create(path)?;
    write_workbook_to_writer(workbook, file)
}

pub fn write_workbook_to_writer<W: Write + Seek>(workbook: &Workbook, writer: W) -> Result<(), XlsxWriteError> {
    let mut zip = ZipWriter::new(writer);
    let options = zip::write::FileOptions::<()>::default()
        .compression_method(zip::CompressionMethod::Deflated);

    let shared_strings = build_shared_strings(workbook);
    let mut style_table = workbook.styles.clone();
    let mut styles_part =
        StylesPart::parse_or_default(None, &mut style_table).map_err(|e| XlsxWriteError::Invalid(e.to_string()))?;
    let style_ids = workbook
        .sheets
        .iter()
        .flat_map(|sheet| sheet.iter_cells().map(|(_, cell)| cell.style_id))
        .filter(|style_id| *style_id != 0 && workbook.styles.get(*style_id).is_some());
    let style_to_xf = styles_part
        .xf_indices_for_style_ids(style_ids, &style_table)
        .map_err(|e| XlsxWriteError::Invalid(e.to_string()))?;
    let styles_xml = styles_part.to_xml_bytes();

    // Root relationships
    zip.start_file("_rels/.rels", options)?;
    zip.write_all(root_rels_xml().as_bytes())?;

    // Content types
    zip.start_file("[Content_Types].xml", options)?;
    zip.write_all(content_types_xml(workbook, &shared_strings).as_bytes())?;

    // Workbook
    zip.start_file("xl/workbook.xml", options)?;
    zip.write_all(workbook_xml(workbook).as_bytes())?;

    // Workbook relationships
    zip.start_file("xl/_rels/workbook.xml.rels", options)?;
    zip.write_all(workbook_rels_xml(workbook, !shared_strings.values.is_empty()).as_bytes())?;

    // Styles
    zip.start_file("xl/styles.xml", options)?;
    zip.write_all(&styles_xml)?;

    // Shared strings
    if !shared_strings.values.is_empty() {
        zip.start_file("xl/sharedStrings.xml", options)?;
        zip.write_all(shared_strings_xml(&shared_strings).as_bytes())?;
    }

    // Tables are written globally and then referenced from sheets.
    let mut next_table_part = 1usize;
    let mut table_parts_by_sheet: Vec<Vec<(String, String)>> = Vec::new(); // sheet_index -> [(rId, target)]

    for sheet in &workbook.sheets {
        let mut parts = Vec::new();
        for (table_idx, table) in sheet.tables.iter().enumerate() {
            let file_name = format!("table{next_table_part}.xml");
            next_table_part += 1;
            let part_path = format!("xl/tables/{file_name}");

            let rel_id = table
                .relationship_id
                .clone()
                .unwrap_or_else(|| format!("rId{}", table_idx + 1));
            parts.push((rel_id, format!("../tables/{file_name}")));

            zip.start_file(&part_path, options)?;
            let xml = write_table_xml(table).map_err(XlsxWriteError::Invalid)?;
            zip.write_all(xml.as_bytes())?;
        }
        table_parts_by_sheet.push(parts);
    }

    // Worksheets + relationships
    for (idx, sheet) in workbook.sheets.iter().enumerate() {
        let sheet_number = idx + 1;
        let sheet_path = format!("xl/worksheets/sheet{sheet_number}.xml");
        zip.start_file(&sheet_path, options)?;
        zip.write_all(
            sheet_xml(
                sheet,
                &shared_strings,
                &table_parts_by_sheet[idx],
                &style_to_xf,
            )
            .as_bytes(),
        )?;

        let rels_path = format!("xl/worksheets/_rels/sheet{sheet_number}.xml.rels");
        zip.start_file(&rels_path, options)?;
        zip.write_all(sheet_rels_xml(&table_parts_by_sheet[idx]).as_bytes())?;
    }

    let _writer = zip.finish()?;
    Ok(())
}

fn root_rels_xml() -> String {
    format!(
        r#"<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
  <Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="xl/workbook.xml"/>
</Relationships>"#
    )
}

fn workbook_xml(workbook: &Workbook) -> String {
    let mut sheets_xml = String::new();
    for (idx, sheet) in workbook.sheets.iter().enumerate() {
        let sheet_id = idx + 1;
        sheets_xml.push_str(&format!(
            r#"<sheet name="{}" sheetId="{}" r:id="rId{}"/>"#,
            escape_xml(&sheet.name),
            sheet_id,
            sheet_id
        ));
    }

    format!(
        r#"<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<workbook xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">
  <sheets>
    {}
  </sheets>
</workbook>"#,
        sheets_xml
    )
}

fn workbook_rels_xml(workbook: &Workbook, has_shared_strings: bool) -> String {
    let mut rels = String::new();
    for (idx, _sheet) in workbook.sheets.iter().enumerate() {
        let rel_id = idx + 1;
        rels.push_str(&format!(
            r#"<Relationship Id="rId{}" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/worksheet" Target="worksheets/sheet{}.xml"/>"#,
            rel_id,
            rel_id
        ));
    }
    let mut next = workbook.sheets.len() + 1;
    if has_shared_strings {
        rels.push_str(&format!(
            r#"<Relationship Id="rId{}" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/sharedStrings" Target="sharedStrings.xml"/>"#,
            next
        ));
        next += 1;
    }
    rels.push_str(&format!(
        r#"<Relationship Id="rId{}" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/styles" Target="styles.xml"/>"#,
        next
    ));

    format!(
        r#"<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
  {}
</Relationships>"#,
        rels
    )
}

fn sheet_rels_xml(table_parts: &[(String, String)]) -> String {
    let mut rels = String::new();
    for (id, target) in table_parts {
        rels.push_str(&format!(
            r#"<Relationship Id="{}" Type="{}" Target="{}"/>"#,
            escape_xml(id),
            TABLE_REL_TYPE,
            escape_xml(target)
        ));
    }
    format!(
        r#"<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
  {}
</Relationships>"#,
        rels
    )
}

fn sheet_xml(
    sheet: &Worksheet,
    shared_strings: &SharedStrings,
    table_parts: &[(String, String)],
    style_to_xf: &HashMap<u32, u32>,
) -> String {
    // Excel expects rows in ascending order.
    let mut rows: BTreeMap<u32, Vec<(u32, CellRef, &Cell)>> = BTreeMap::new();
    for (cell_ref, cell) in sheet.iter_cells() {
        rows.entry(cell_ref.row)
            .or_default()
            .push((cell_ref.col, cell_ref, cell));
    }
    for row_cells in rows.values_mut() {
        row_cells.sort_by_key(|(col, _, _)| *col);
    }

    let mut sheet_data = String::new();
    for (row_idx, cells) in rows {
        let row_number = row_idx + 1;
        sheet_data.push_str(&format!(r#"<row r="{}">"#, row_number));
        for (_col, cell_ref, cell) in cells {
            sheet_data.push_str(&cell_xml(&cell_ref, cell, shared_strings, style_to_xf));
        }
        sheet_data.push_str("</row>");
    }

    let table_parts_xml = if table_parts.is_empty() {
        String::new()
    } else {
        let parts: String = table_parts
            .iter()
            .map(|(id, _target)| format!(r#"<tablePart r:id="{}"/>"#, escape_xml(id)))
            .collect();
        format!(
            r#"<tableParts count="{}">{}</tableParts>"#,
            table_parts.len(),
            parts
        )
    };

    format!(
        r#"<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">
  <sheetData>
    {}
  </sheetData>
  {}
</worksheet>"#,
        sheet_data, table_parts_xml
    )
}

fn cell_xml(
    cell_ref: &CellRef,
    cell: &Cell,
    shared_strings: &SharedStrings,
    style_to_xf: &HashMap<u32, u32>,
) -> String {
    let a1 = cell_ref.to_a1();
    let mut attrs = format!(r#" r="{}""#, a1);
    let mut value_xml = String::new();

    if cell.style_id != 0 {
        if let Some(xf_index) = style_to_xf.get(&cell.style_id).copied().filter(|xf| *xf != 0) {
            attrs.push_str(&format!(r#" s="{}""#, xf_index));
        }
    }

    if let Some(formula) = &cell.formula {
        let formula = normalize_formula_text(formula);
        if !formula.is_empty() {
            value_xml.push_str(&format!(r#"<f>{}</f>"#, escape_xml(&formula)));
        }
    }

    match &cell.value {
        CellValue::Empty => {}
        CellValue::Number(n) => {
            value_xml.push_str(&format!(r#"<v>{}</v>"#, n));
        }
        CellValue::Boolean(b) => {
            attrs.push_str(r#" t="b""#);
            value_xml.push_str(&format!(r#"<v>{}</v>"#, if *b { 1 } else { 0 }));
        }
        CellValue::String(s) => {
            attrs.push_str(r#" t="s""#);
            let idx = shared_strings
                .index
                .get(s)
                .copied()
                .unwrap_or_default();
            value_xml.push_str(&format!(r#"<v>{}</v>"#, idx));
        }
        CellValue::Error(e) => {
            attrs.push_str(r#" t="e""#);
            value_xml.push_str(&format!(r#"<v>{}</v>"#, escape_xml(e.as_str())));
        }
        CellValue::RichText(r) => {
            attrs.push_str(r#" t="s""#);
            let idx = shared_strings
                .index
                .get(&r.text)
                .copied()
                .unwrap_or_default();
            value_xml.push_str(&format!(r#"<v>{}</v>"#, idx));
        }
        CellValue::Array(_) | CellValue::Spill(_) => {}
    }

    format!(r#"<c{}>{}</c>"#, attrs, value_xml)
}

#[derive(Debug, Clone)]
struct SharedStrings {
    values: Vec<String>,
    index: HashMap<String, usize>,
}

fn build_shared_strings(workbook: &Workbook) -> SharedStrings {
    let mut values: Vec<String> = Vec::new();
    let mut index: HashMap<String, usize> = HashMap::new();

    for sheet in &workbook.sheets {
        for (_cell_ref, cell) in sheet.iter_cells() {
            match &cell.value {
                CellValue::String(s) => {
                    if !index.contains_key(s) {
                        let idx = values.len();
                        values.push(s.clone());
                        index.insert(s.clone(), idx);
                    }
                }
                CellValue::RichText(r) => {
                    if !index.contains_key(&r.text) {
                        let idx = values.len();
                        values.push(r.text.clone());
                        index.insert(r.text.clone(), idx);
                    }
                }
                _ => {}
            }
        }
    }

    SharedStrings { values, index }
}

fn shared_strings_xml(shared: &SharedStrings) -> String {
    let count = shared.values.len();
    let mut si = String::new();
    for v in &shared.values {
        si.push_str(&format!(r#"<si><t>{}</t></si>"#, escape_xml(v)));
    }
    format!(
        r#"<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<sst xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" count="{count}" uniqueCount="{count}">
  {si}
</sst>"#
    )
}

fn content_types_xml(workbook: &Workbook, shared_strings: &SharedStrings) -> String {
    let mut overrides = String::new();
    overrides.push_str(
        r#"<Override PartName="/xl/workbook.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet.main+xml"/>"#,
    );
    overrides.push_str(
        r#"<Override PartName="/xl/styles.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.styles+xml"/>"#,
    );
    if !shared_strings.values.is_empty() {
        overrides.push_str(
            r#"<Override PartName="/xl/sharedStrings.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.sharedStrings+xml"/>"#,
        );
    }
    for (idx, _) in workbook.sheets.iter().enumerate() {
        let sheet_number = idx + 1;
        overrides.push_str(&format!(
            r#"<Override PartName="/xl/worksheets/sheet{sheet_number}.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.worksheet+xml"/>"#
        ));
        if !workbook.sheets[idx].tables.is_empty() {
            overrides.push_str(&format!(
                r#"<Override PartName="/xl/worksheets/_rels/sheet{sheet_number}.xml.rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/>"#
            ));
        }
    }

    let mut table_count = 1usize;
    for sheet in &workbook.sheets {
        for _ in &sheet.tables {
            overrides.push_str(&format!(
                r#"<Override PartName="/xl/tables/table{table_count}.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.table+xml"/>"#
            ));
            table_count += 1;
        }
    }

    format!(
        r#"<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">
  <Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/>
  <Default Extension="xml" ContentType="application/xml"/>
  {}
</Types>"#,
        overrides
    )
}

fn escape_xml(s: &str) -> String {
    s.replace('&', "&amp;")
        .replace('<', "&lt;")
        .replace('>', "&gt;")
        .replace('"', "&quot;")
        .replace('\'', "&apos;")
}
