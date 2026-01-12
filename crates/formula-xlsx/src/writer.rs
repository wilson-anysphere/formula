use crate::styles::StylesPart;
use crate::tables::{write_table_xml, TABLE_REL_TYPE};
use crate::WorkbookKind;
use formula_columnar::{ColumnType as ColumnarType, Value as ColumnarValue};
use formula_model::{
    normalize_formula_text, Cell, CellRef, CellValue, DateSystem, DefinedNameScope, Hyperlink,
    HyperlinkTarget, Range, SheetVisibility, Workbook, Worksheet,
};
use std::collections::{BTreeMap, HashMap, HashSet};
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
    let path = path.as_ref();
    let file = File::create(path)?;
    let ext = path
        .extension()
        .and_then(|s| s.to_str())
        .unwrap_or_default()
        .to_ascii_lowercase();
    let kind = WorkbookKind::from_extension(&ext).unwrap_or(WorkbookKind::Workbook);
    write_workbook_to_writer_with_kind(workbook, file, kind)
}

pub fn write_workbook_to_writer<W: Write + Seek>(
    workbook: &Workbook,
    writer: W,
) -> Result<(), XlsxWriteError> {
    write_workbook_to_writer_with_kind(workbook, writer, WorkbookKind::Workbook)
}

pub fn write_workbook_to_writer_with_kind<W: Write + Seek>(
    workbook: &Workbook,
    writer: W,
    kind: WorkbookKind,
) -> Result<(), XlsxWriteError> {
    let mut zip = ZipWriter::new(writer);
    let options = zip::write::FileOptions::<()>::default()
        .compression_method(zip::CompressionMethod::Deflated);

    let shared_strings = build_shared_strings(workbook);
    let mut style_table = workbook.styles.clone();
    let mut styles_part = StylesPart::parse_or_default(None, &mut style_table)
        .map_err(|e| XlsxWriteError::Invalid(e.to_string()))?;
    let style_ids = workbook
        .sheets
        .iter()
        .flat_map(|sheet| sheet.iter_cells().map(|(_, cell)| cell.style_id))
        .filter(|style_id| *style_id != 0);
    let style_to_xf = styles_part
        .xf_indices_for_style_ids(style_ids, &style_table)
        .map_err(|e| XlsxWriteError::Invalid(e.to_string()))?;
    let styles_xml = styles_part.to_xml_bytes();

    // Root relationships
    zip.start_file("_rels/.rels", options)?;
    zip.write_all(root_rels_xml().as_bytes())?;

    // Content types
    zip.start_file("[Content_Types].xml", options)?;
    zip.write_all(content_types_xml(workbook, &shared_strings, kind).as_bytes())?;

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
        let (sheet_xml, sheet_rels) = sheet_xml(
            sheet,
            &shared_strings,
            &table_parts_by_sheet[idx],
            &style_to_xf,
        )?;
        zip.start_file(&sheet_path, options)?;
        zip.write_all(sheet_xml.as_bytes())?;

        let rels_path = format!("xl/worksheets/_rels/sheet{sheet_number}.xml.rels");
        zip.start_file(&rels_path, options)?;
        zip.write_all(sheet_rels.as_bytes())?;
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
    let workbook_pr = match workbook.date_system {
        DateSystem::Excel1900 => r#"<workbookPr/>"#.to_string(),
        DateSystem::Excel1904 => r#"<workbookPr date1904="1"/>"#.to_string(),
    };
    let calc_pr = calc_pr_xml(workbook);

    let mut sheets_xml = String::new();
    for (idx, sheet) in workbook.sheets.iter().enumerate() {
        let sheet_id = idx + 1;
        let state = match sheet.visibility {
            SheetVisibility::Visible => "",
            SheetVisibility::Hidden => r#" state="hidden""#,
            SheetVisibility::VeryHidden => r#" state="veryHidden""#,
        };
        sheets_xml.push_str(&format!(
            r#"<sheet name="{}" sheetId="{}" r:id="rId{}"{} />"#,
            escape_xml(&sheet.name),
            sheet_id,
            sheet_id,
            state
        ));
    }

    let defined_names_xml = workbook_defined_names_xml(workbook);

    format!(
        r#"<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<workbook xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">
  {}
  <sheets>
    {}
  </sheets>
  {}
  {}
</workbook>"#,
        workbook_pr, sheets_xml, defined_names_xml, calc_pr
    )
}

fn calc_pr_xml(workbook: &Workbook) -> String {
    let settings = &workbook.calc_settings;
    format!(
        r#"<calcPr calcMode="{}" calcOnSave="{}" fullCalcOnLoad="{}" iterative="{}" iterateCount="{}" iterateDelta="{}" fullPrecision="{}"/>"#,
        settings.calculation_mode.as_calc_mode_attr(),
        bool_attr(settings.calculate_before_save),
        bool_attr(settings.full_calc_on_load),
        bool_attr(settings.iterative.enabled),
        settings.iterative.max_iterations,
        trim_float(settings.iterative.max_change),
        bool_attr(settings.full_precision),
    )
}

fn bool_attr(value: bool) -> &'static str {
    if value {
        "1"
    } else {
        "0"
    }
}

fn trim_float(value: f64) -> String {
    let s = format!("{value:.15}");
    let s = s.trim_end_matches('0').trim_end_matches('.');
    if s.is_empty() {
        "0".to_string()
    } else {
        s.to_string()
    }
}

fn workbook_defined_names_xml(workbook: &Workbook) -> String {
    if workbook.defined_names.is_empty() {
        return String::new();
    }

    let mut sheet_index_by_id = HashMap::new();
    for (idx, sheet) in workbook.sheets.iter().enumerate() {
        sheet_index_by_id.insert(sheet.id, idx as u32);
    }

    let mut out = String::new();
    out.push_str("<definedNames>");
    for defined in &workbook.defined_names {
        // Defined name `refersTo` values are stored in workbook.xml without a leading '=' but still
        // use the same `_xlfn.`-prefixed function naming as cell formulas for forward-compatible
        // functions.
        let refers_to = crate::formula_text::add_xlfn_prefixes(&defined.refers_to);
        out.push_str(r#"<definedName"#);
        out.push_str(&format!(r#" name="{}""#, escape_xml(&defined.name)));
        if let Some(comment) = &defined.comment {
            out.push_str(&format!(r#" comment="{}""#, escape_xml(comment)));
        }
        if defined.hidden {
            out.push_str(r#" hidden="1""#);
        }
        if let DefinedNameScope::Sheet(sheet_id) = defined.scope {
            if let Some(idx) = sheet_index_by_id.get(&sheet_id) {
                out.push_str(&format!(r#" localSheetId="{}""#, idx));
            }
        }
        out.push('>');
        out.push_str(&escape_xml(&refers_to));
        out.push_str("</definedName>");
    }
    out.push_str("</definedNames>");
    out
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
) -> Result<(String, String), XlsxWriteError> {
    // Dimension should include both the columnar table extent and any sparse overlay cells.
    let mut dim: Option<Range> = sheet.used_range();
    if dim.is_none() {
        // Some sheet sources may not maintain used_range; fall back to scanning.
        let mut min: Option<CellRef> = None;
        let mut max: Option<CellRef> = None;
        for (cell_ref, _) in sheet.iter_cells() {
            min = Some(match min {
                Some(m) => CellRef::new(m.row.min(cell_ref.row), m.col.min(cell_ref.col)),
                None => cell_ref,
            });
            max = Some(match max {
                Some(m) => CellRef::new(m.row.max(cell_ref.row), m.col.max(cell_ref.col)),
                None => cell_ref,
            });
        }
        dim = match (min, max) {
            (Some(start), Some(end)) => Some(Range::new(start, end)),
            _ => None,
        };
    }
    if let Some(columnar_range) = sheet.columnar_range() {
        dim = Some(match dim {
            Some(existing) => existing.bounding_box(&columnar_range),
            None => columnar_range,
        });
    }
    let dimension_ref = dim
        .unwrap_or_else(|| Range::new(CellRef::new(0, 0), CellRef::new(0, 0)))
        .to_string();

    struct ColumnarInfo<'a> {
        origin: CellRef,
        rows: usize,
        cols: usize,
        table: &'a formula_columnar::ColumnarTable,
    }

    let columnar = sheet
        .columnar_table_extent()
        .and_then(|(origin, rows, cols)| {
            sheet.columnar_table().map(|t| ColumnarInfo {
                origin,
                rows,
                cols,
                table: t.as_ref(),
            })
        });

    // Group overlay cells by row for streaming output.
    let mut overlay_by_row: BTreeMap<u32, Vec<(u32, CellRef, &Cell)>> = BTreeMap::new();
    for (cell_ref, cell) in sheet.iter_cells() {
        overlay_by_row
            .entry(cell_ref.row)
            .or_default()
            .push((cell_ref.col, cell_ref, cell));
    }
    for row_cells in overlay_by_row.values_mut() {
        row_cells.sort_by_key(|(col, _, _)| *col);
    }
    let overlay_rows: Vec<u32> = overlay_by_row.keys().copied().collect();

    // Emit rows in ascending order, streaming through the columnar table rows if present.
    let mut sheet_data = String::new();
    let mut overlay_row_idx: usize = 0;
    let mut table_row: Option<u32> = columnar.as_ref().map(|c| c.origin.row);
    let table_end_row: Option<u32> = columnar
        .as_ref()
        .map(|c| c.origin.row.saturating_add(c.rows.saturating_sub(1) as u32));

    loop {
        let next_overlay_row = overlay_rows.get(overlay_row_idx).copied();
        let next_table_row = match (table_row, table_end_row) {
            (Some(r), Some(end)) if r <= end => Some(r),
            _ => None,
        };

        let Some(row_idx) = (match (next_table_row, next_overlay_row) {
            (Some(t), Some(o)) => Some(t.min(o)),
            (Some(t), None) => Some(t),
            (None, Some(o)) => Some(o),
            (None, None) => None,
        }) else {
            break;
        };

        if next_overlay_row == Some(row_idx) {
            overlay_row_idx += 1;
        }
        if next_table_row == Some(row_idx) {
            table_row = Some(row_idx + 1);
        }

        let overlay_cells: &[(u32, CellRef, &Cell)] = overlay_by_row
            .get(&row_idx)
            .map(Vec::as_slice)
            .unwrap_or(&[]);

        let mut row_cells_xml = String::new();
        let mut wrote_any_cell = false;

        if let Some(columnar) = columnar.as_ref() {
            let in_table_row = row_idx >= columnar.origin.row
                && row_idx < columnar.origin.row.saturating_add(columnar.rows as u32);
            if in_table_row {
                let row_off = (row_idx - columnar.origin.row) as usize;
                let mut overlay_cell_idx = 0usize;

                // Overlay cells left of the table.
                while overlay_cell_idx < overlay_cells.len()
                    && overlay_cells[overlay_cell_idx].0 < columnar.origin.col
                {
                    let (_col, cell_ref, cell) = overlay_cells[overlay_cell_idx];
                    row_cells_xml.push_str(&cell_xml(&cell_ref, cell, shared_strings, style_to_xf));
                    overlay_cell_idx += 1;
                    wrote_any_cell = true;
                }

                // Table columns (overlay overrides).
                for col_off in 0..columnar.cols {
                    let col_idx = columnar.origin.col + col_off as u32;
                    if overlay_cell_idx < overlay_cells.len()
                        && overlay_cells[overlay_cell_idx].0 == col_idx
                    {
                        let (_col, cell_ref, cell) = overlay_cells[overlay_cell_idx];
                        row_cells_xml.push_str(&cell_xml(
                            &cell_ref,
                            cell,
                            shared_strings,
                            style_to_xf,
                        ));
                        overlay_cell_idx += 1;
                        wrote_any_cell = true;
                        continue;
                    }

                    let cell_ref = CellRef::new(row_idx, col_idx);
                    if sheet.merged_regions.resolve_cell(cell_ref) != cell_ref {
                        continue;
                    }

                    let value = columnar.table.get_cell(row_off, col_off);
                    if matches!(value, ColumnarValue::Null) {
                        continue;
                    }
                    let column_type = columnar
                        .table
                        .schema()
                        .get(col_off)
                        .map(|s| s.column_type)
                        .unwrap_or(ColumnarType::String);
                    if let Some(xml) =
                        columnar_cell_xml(&cell_ref, value, column_type, shared_strings)
                    {
                        row_cells_xml.push_str(&xml);
                        wrote_any_cell = true;
                    }
                }

                // Overlay cells right of the table.
                while overlay_cell_idx < overlay_cells.len() {
                    let (_col, cell_ref, cell) = overlay_cells[overlay_cell_idx];
                    row_cells_xml.push_str(&cell_xml(&cell_ref, cell, shared_strings, style_to_xf));
                    overlay_cell_idx += 1;
                    wrote_any_cell = true;
                }
            } else {
                // Row outside the columnar table; only overlay cells apply.
                for (_col, cell_ref, cell) in overlay_cells {
                    row_cells_xml.push_str(&cell_xml(cell_ref, cell, shared_strings, style_to_xf));
                    wrote_any_cell = true;
                }
            }
        } else {
            // No columnar table; only overlay cells apply.
            for (_col, cell_ref, cell) in overlay_cells {
                row_cells_xml.push_str(&cell_xml(cell_ref, cell, shared_strings, style_to_xf));
                wrote_any_cell = true;
            }
        }

        if !wrote_any_cell {
            continue;
        }

        let row_number = row_idx + 1;
        sheet_data.push_str(&format!(r#"<row r="{}">"#, row_number));
        sheet_data.push_str(&row_cells_xml);
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

    let auto_filter_xml = if let Some(filter) = sheet.auto_filter.as_ref() {
        crate::autofilter::write_autofilter(filter)
            .map_err(|e| XlsxWriteError::Invalid(e.to_string()))?
    } else {
        String::new()
    };

    let mut xml = String::new();
    xml.push_str(r#"<?xml version="1.0" encoding="UTF-8" standalone="yes"?>"#);
    xml.push('\n');
    xml.push_str(r#"<worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">"#);
    xml.push('\n');
    xml.push_str(&format!(r#"  <dimension ref="{dimension_ref}"/>"#));
    xml.push('\n');
    xml.push_str("  <sheetData>\n");
    if !sheet_data.is_empty() {
        xml.push_str("    ");
        xml.push_str(&sheet_data);
        xml.push('\n');
    }
    xml.push_str("  </sheetData>\n");
    if !auto_filter_xml.is_empty() {
        xml.push_str("  ");
        xml.push_str(&auto_filter_xml);
        xml.push('\n');
    }
    if !table_parts_xml.is_empty() {
        xml.push_str("  ");
        xml.push_str(&table_parts_xml);
        xml.push('\n');
    }
    xml.push_str("</worksheet>");

    if sheet.tab_color.is_some() {
        xml = crate::sheet_metadata::write_sheet_tab_color(&xml, sheet.tab_color.as_ref())
            .map_err(|e| XlsxWriteError::Invalid(e.to_string()))?;
    }

    let mut merges: Vec<Range> = sheet
        .merged_regions
        .iter()
        .map(|region| region.range)
        .filter(|range| !range.is_single_cell())
        .collect();
    merges.sort_by_key(|range| {
        (
            range.start.row,
            range.start.col,
            range.end.row,
            range.end.col,
        )
    });
    if !merges.is_empty() {
        xml = crate::merge_cells::update_worksheet_xml(&xml, &merges)
            .map_err(|e| XlsxWriteError::Invalid(e.to_string()))?;
    }

    // Generate a safe set of hyperlink relationship IDs for this sheet.
    let mut used_rel_ids: HashSet<String> = table_parts.iter().map(|(id, _)| id.clone()).collect();
    let mut next_rel_id = used_rel_ids
        .iter()
        .filter_map(|id| id.strip_prefix("rId")?.parse::<u32>().ok())
        .max()
        .unwrap_or(0)
        + 1;

    let mut links: Vec<Hyperlink> = sheet.hyperlinks.clone();
    let mut target_by_rel_id: HashMap<String, String> = HashMap::new();
    for link in &mut links {
        let target = match &link.target {
            HyperlinkTarget::ExternalUrl { uri } => Some(uri.as_str()),
            HyperlinkTarget::Email { uri } => Some(uri.as_str()),
            HyperlinkTarget::Internal { .. } => None,
        };
        let Some(target) = target else {
            continue;
        };

        let mut rel_id = link.rel_id.clone();
        let needs_new = match rel_id.as_deref() {
            None => true,
            Some(id) if used_rel_ids.contains(id) && !target_by_rel_id.contains_key(id) => true,
            Some(id) => target_by_rel_id
                .get(id)
                .is_some_and(|existing| existing != target),
        };
        if needs_new {
            loop {
                let candidate = format!("rId{next_rel_id}");
                next_rel_id += 1;
                if used_rel_ids.insert(candidate.clone()) {
                    rel_id = Some(candidate);
                    break;
                }
            }
        } else if let Some(id) = rel_id.as_ref() {
            used_rel_ids.insert(id.clone());
        }

        let id = rel_id.expect("rel id ensured for external hyperlinks");
        link.rel_id = Some(id.clone());
        target_by_rel_id
            .entry(id)
            .or_insert_with(|| target.to_string());
    }

    if !links.is_empty() {
        xml = crate::update_worksheet_xml(&xml, &links)
            .map_err(|e| XlsxWriteError::Invalid(e.to_string()))?;
    }

    let rels_xml = {
        let base = sheet_rels_xml(table_parts);
        // Only external hyperlinks need relationships; internal hyperlinks are stored as `location=`.
        if links.iter().any(|link| {
            matches!(
                link.target,
                HyperlinkTarget::ExternalUrl { .. } | HyperlinkTarget::Email { .. }
            )
        }) {
            crate::update_worksheet_relationships(Some(&base), &links)
                .map_err(|e| XlsxWriteError::Invalid(e.to_string()))?
                .unwrap_or_else(|| sheet_rels_xml(&[]))
        } else {
            base
        }
    };

    Ok((xml, rels_xml))
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
        if let Some(xf_index) = style_to_xf
            .get(&cell.style_id)
            .copied()
            .filter(|xf| *xf != 0)
        {
            attrs.push_str(&format!(r#" s="{}""#, xf_index));
        }
    }

    if let Some(formula) = &cell.formula {
        if let Some(formula) = normalize_formula_text(formula) {
            let file_formula = crate::formula_text::add_xlfn_prefixes(&formula);
            value_xml.push_str(&format!(r#"<f>{}</f>"#, escape_xml(&file_formula)));
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
            let idx = shared_strings.index.get(s).copied().unwrap_or_default();
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
        CellValue::Entity(entity) => {
            attrs.push_str(r#" t="s""#);
            let idx = shared_strings
                .index
                .get(&entity.display_value)
                .copied()
                .unwrap_or_default();
            value_xml.push_str(&format!(r#"<v>{}</v>"#, idx));
        }
        CellValue::Record(record) => {
            let display = record_display_string(record);
            attrs.push_str(r#" t="s""#);
            let idx = shared_strings.index.get(&display).copied().unwrap_or_default();
            value_xml.push_str(&format!(r#"<v>{}</v>"#, idx));
        }
        CellValue::Image(image) => {
            // In-cell images are not yet exported as first-class XLSX rich values. Degrade to
            // plain text when alt text is available; otherwise omit the cached value.
            if let Some(alt) = image.alt_text.as_deref().filter(|s| !s.is_empty()) {
                attrs.push_str(r#" t="s""#);
                let idx = shared_strings.index.get(alt).copied().unwrap_or_default();
                value_xml.push_str(&format!(r#"<v>{}</v>"#, idx));
            }
        }
        CellValue::Array(_) | CellValue::Spill(_) => {}
    }

    format!(r#"<c{}>{}</c>"#, attrs, value_xml)
}

fn columnar_cell_xml(
    cell_ref: &CellRef,
    value: ColumnarValue,
    column_type: ColumnarType,
    shared_strings: &SharedStrings,
) -> Option<String> {
    let a1 = cell_ref.to_a1();
    let mut attrs = format!(r#" r="{}""#, a1);
    let mut value_xml = String::new();

    match value {
        ColumnarValue::Null => return None,
        ColumnarValue::Number(n) => {
            value_xml.push_str(&format!(r#"<v>{}</v>"#, n));
        }
        ColumnarValue::Boolean(b) => {
            attrs.push_str(r#" t="b""#);
            value_xml.push_str(&format!(r#"<v>{}</v>"#, if b { 1 } else { 0 }));
        }
        ColumnarValue::String(s) => {
            attrs.push_str(r#" t="s""#);
            let idx = shared_strings
                .index
                .get(s.as_ref())
                .copied()
                .unwrap_or_default();
            value_xml.push_str(&format!(r#"<v>{}</v>"#, idx));
        }
        ColumnarValue::DateTime(v) => {
            value_xml.push_str(&format!(r#"<v>{}</v>"#, v as f64));
        }
        ColumnarValue::Currency(v) => {
            let n = match column_type {
                ColumnarType::Currency { scale } => {
                    let denom = 10f64.powi(scale as i32);
                    v as f64 / denom
                }
                _ => v as f64,
            };
            value_xml.push_str(&format!(r#"<v>{}</v>"#, n));
        }
        ColumnarValue::Percentage(v) => {
            let n = match column_type {
                ColumnarType::Percentage { scale } => {
                    let denom = 10f64.powi(scale as i32);
                    v as f64 / denom
                }
                _ => v as f64,
            };
            value_xml.push_str(&format!(r#"<v>{}</v>"#, n));
        }
    }

    Some(format!(r#"<c{}>{}</c>"#, attrs, value_xml))
}

#[derive(Debug, Clone)]
struct SharedStrings {
    values: Vec<String>,
    index: HashMap<String, usize>,
}

fn record_display_string(record: &formula_model::RecordValue) -> String {
    record.to_string()
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
                CellValue::Entity(entity) => {
                    let s = entity.display_value.clone();
                    if !index.contains_key(&s) {
                        let idx = values.len();
                        values.push(s.clone());
                        index.insert(s, idx);
                    }
                }
                CellValue::Record(record) => {
                    let s = record.to_string();
                    if !index.contains_key(&s) {
                        let idx = values.len();
                        values.push(s.clone());
                        index.insert(s, idx);
                    }
                }
                CellValue::Image(image) => {
                    if let Some(alt) = image.alt_text.as_deref().filter(|s| !s.is_empty()) {
                        if !index.contains_key(alt) {
                            let idx = values.len();
                            values.push(alt.to_string());
                            index.insert(alt.to_string(), idx);
                        }
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

        if let Some((_, rows, cols)) = sheet.columnar_table_extent() {
            if let Some(table) = sheet.columnar_table() {
                let table = table.as_ref();
                for row in 0..rows {
                    for col in 0..cols {
                        if let ColumnarValue::String(s) = table.get_cell(row, col) {
                            let text = s.as_ref();
                            if !index.contains_key(text) {
                                let idx = values.len();
                                let owned = text.to_string();
                                values.push(owned.clone());
                                index.insert(owned, idx);
                            }
                        }
                    }
                }
            }
        }
    }

    SharedStrings { values, index }
}

fn shared_strings_xml(shared: &SharedStrings) -> String {
    let count = shared.values.len();
    let mut si = String::new();
    for v in &shared.values {
        let preserve = v
            .chars()
            .next()
            .map(|c| c.is_whitespace())
            .unwrap_or(false)
            || v.chars().last().map(|c| c.is_whitespace()).unwrap_or(false);
        if preserve {
            si.push_str(&format!(
                r#"<si><t xml:space="preserve">{}</t></si>"#,
                escape_xml(v)
            ));
        } else {
            si.push_str(&format!(r#"<si><t>{}</t></si>"#, escape_xml(v)));
        }
    }
    format!(
        r#"<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<sst xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" count="{count}" uniqueCount="{count}">
  {si}
</sst>"#
    )
}

fn content_types_xml(
    workbook: &Workbook,
    shared_strings: &SharedStrings,
    kind: WorkbookKind,
) -> String {
    let mut overrides = String::new();
    overrides.push_str(&format!(
        r#"<Override PartName="/xl/workbook.xml" ContentType="{}"/>"#,
        kind.workbook_content_type()
    ));
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
