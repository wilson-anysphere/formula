use crate::styles::StylesPart;
use crate::tables::{write_table_xml, TABLE_REL_TYPE};
use crate::WorkbookKind;
use crate::ConditionalFormattingDxfAggregation;
use formula_columnar::{ColumnType as ColumnarType, Value as ColumnarValue};
use formula_model::{
    normalize_formula_text, Cell, CellIsOperator, CellRef, CellValue, CfRule, CfRuleKind,
    DataValidationErrorStyle, DataValidationKind, DataValidationOperator, DateSystem,
    DefinedNameScope, Hyperlink, HyperlinkTarget, ManualPageBreaks, Outline, PageMargins,
    PageSetup, Range, Scaling, SheetPrintSettings, SheetVisibility, Workbook, WorkbookWindowState,
    Worksheet,
};
use formula_fs::{atomic_write_with_path, AtomicWriteError};
use std::collections::{BTreeMap, HashMap, HashSet};
use std::fs::File;
use std::io::{Cursor, Seek, Write};
use std::path::Path;
use thiserror::Error;
use zip::ZipWriter;
use formula_model::rich_text::{RichText, Underline};

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
    let ext = path
        .extension()
        .and_then(|s| s.to_str())
        .unwrap_or_default()
        .to_ascii_lowercase();
    let kind = WorkbookKind::from_extension(&ext).unwrap_or(WorkbookKind::Workbook);
    atomic_write_with_path(path, |tmp_path| {
        let file = File::create(tmp_path)?;
        write_workbook_to_writer_with_kind(workbook, file, kind)
    })
    .map_err(|err| match err {
        AtomicWriteError::Io(err) => XlsxWriteError::Io(err),
        AtomicWriteError::Writer(err) => err,
    })
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
    let style_ids = workbook.sheets.iter().flat_map(|sheet| {
        sheet
            .iter_cells()
            .map(|(_, cell)| cell.style_id)
            .filter(|style_id| *style_id != 0)
            .chain(
                sheet
                    .row_properties
                    .values()
                    .filter_map(|props| props.style_id),
            )
            .chain(
                sheet
                    .col_properties
                    .values()
                    .filter_map(|props| props.style_id),
            )
    });
    let style_to_xf = styles_part
        .xf_indices_for_style_ids(style_ids, &style_table)
        .map_err(|e| XlsxWriteError::Invalid(e.to_string()))?;

    // Conditional formatting dxfs live in a single global `<dxfs>` table inside styles.xml, but the
    // in-memory model stores them per-sheet. Aggregate and deduplicate deterministically, then
    // remap per-sheet `cfRule/@dxfId` values during worksheet writing.
    let cf_dxfs = ConditionalFormattingDxfAggregation::from_worksheets(&workbook.sheets);
    styles_part.set_conditional_formatting_dxfs(&cf_dxfs.global_dxfs);

    let styles_xml = styles_part.to_xml_bytes();

    // Root relationships
    zip.start_file("_rels/.rels", options)?;
    zip.write_all(root_rels_xml().as_bytes())?;

    // Content types
    zip.start_file("[Content_Types].xml", options)?;
    zip.write_all(content_types_xml(workbook, &shared_strings, kind).as_bytes())?;

    // Document properties
    zip.start_file("docProps/core.xml", options)?;
    zip.write_all(core_properties_xml().as_bytes())?;
    zip.start_file("docProps/app.xml", options)?;
    zip.write_all(app_properties_xml(workbook).as_bytes())?;

    // Workbook
    zip.start_file("xl/workbook.xml", options)?;
    zip.write_all(workbook_xml(workbook).as_bytes())?;

    // Workbook relationships
    zip.start_file("xl/_rels/workbook.xml.rels", options)?;
    zip.write_all(workbook_rels_xml(workbook, !shared_strings.values.is_empty()).as_bytes())?;

    // Theme
    zip.start_file("xl/theme/theme1.xml", options)?;
    zip.write_all(theme_xml(workbook).as_bytes())?;

    // Styles
    zip.start_file("xl/styles.xml", options)?;
    zip.write_all(&styles_xml)?;

    // Shared strings
    if !shared_strings.values.is_empty() {
        zip.start_file("xl/sharedStrings.xml", options)?;
        let xml = crate::shared_strings::write_shared_strings_xml(&shared_strings.values)
            .map_err(|e| XlsxWriteError::Invalid(e.to_string()))?;
        zip.write_all(xml.as_bytes())?;
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
    let mut print_settings_by_sheet_name: HashMap<String, &SheetPrintSettings> = HashMap::new();
    for sheet_settings in &workbook.print_settings.sheets {
        if sheet_settings.sheet_name.is_empty() {
            continue;
        }
        print_settings_by_sheet_name.insert(
            sheet_settings.sheet_name.to_ascii_uppercase(),
            sheet_settings,
        );
    }

    for (idx, sheet) in workbook.sheets.iter().enumerate() {
        let sheet_number = idx + 1;
        let sheet_path = format!("xl/worksheets/sheet{sheet_number}.xml");
        let sheet_print_settings = print_settings_by_sheet_name
            .get(&sheet.name.to_ascii_uppercase())
            .copied();
        let (sheet_xml, sheet_rels) = sheet_xml(
            sheet,
            sheet_print_settings,
            &shared_strings,
            &table_parts_by_sheet[idx],
            &style_to_xf,
            cf_dxfs.local_to_global_by_sheet.get(&sheet.id).map(|v| v.as_slice()),
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

pub fn write_workbook_to_writer_encrypted<W: Write>(
    workbook: &Workbook,
    mut writer: W,
    kind: WorkbookKind,
    password: &str,
) -> Result<(), XlsxWriteError> {
    let mut cursor = Cursor::new(Vec::new());
    write_workbook_to_writer_with_kind(workbook, &mut cursor, kind)?;
    let zip_bytes = cursor.into_inner();

    let ole_bytes = crate::office_crypto::encrypt_package_to_ole(&zip_bytes, password)
        .map_err(|err| XlsxWriteError::Invalid(format!("office encryption error: {err}")))?;
    writer.write_all(&ole_bytes)?;
    Ok(())
}

fn root_rels_xml() -> String {
    format!(
        r#"<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
  <Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="xl/workbook.xml"/>
  <Relationship Id="rId2" Type="http://schemas.openxmlformats.org/package/2006/relationships/metadata/core-properties" Target="docProps/core.xml"/>
  <Relationship Id="rId3" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/extended-properties" Target="docProps/app.xml"/>
</Relationships>"#
    )
}

fn core_properties_xml() -> String {
    // We don't currently have timestamps/author information in the workbook model, but some
    // downstream consumers expect the docProps part to exist. Use deterministic placeholders to
    // keep output stable between runs.
    let timestamp = "2000-01-01T00:00:00Z";
    let author = "formula-xlsx";
    format!(
        r#"<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<cp:coreProperties xmlns:cp="http://schemas.openxmlformats.org/package/2006/metadata/core-properties" xmlns:dc="http://purl.org/dc/elements/1.1/" xmlns:dcterms="http://purl.org/dc/terms/" xmlns:dcmitype="http://purl.org/dc/dcmitype/" xmlns:xsi="http://www.w3.org/2001/XMLSchema-instance">
  <dc:creator>{author}</dc:creator>
  <cp:lastModifiedBy>{author}</cp:lastModifiedBy>
  <dcterms:created xsi:type="dcterms:W3CDTF">{timestamp}</dcterms:created>
  <dcterms:modified xsi:type="dcterms:W3CDTF">{timestamp}</dcterms:modified>
</cp:coreProperties>"#
    )
}

fn app_properties_xml(workbook: &Workbook) -> String {
    let sheet_count = workbook.sheets.len();
    let mut sheet_names = String::new();
    for sheet in &workbook.sheets {
        sheet_names.push_str(&format!(
            r#"<vt:lpstr>{}</vt:lpstr>"#,
            escape_xml(&sheet.name)
        ));
    }
    format!(
        r#"<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Properties xmlns="http://schemas.openxmlformats.org/officeDocument/2006/extended-properties" xmlns:vt="http://schemas.openxmlformats.org/officeDocument/2006/docPropsVTypes">
  <Application>Microsoft Excel</Application>
  <DocSecurity>0</DocSecurity>
  <ScaleCrop>false</ScaleCrop>
  <HeadingPairs>
    <vt:vector size="2" baseType="variant">
      <vt:variant>
        <vt:lpstr>Worksheets</vt:lpstr>
      </vt:variant>
      <vt:variant>
        <vt:i4>{sheet_count}</vt:i4>
      </vt:variant>
    </vt:vector>
  </HeadingPairs>
  <TitlesOfParts>
    <vt:vector size="{sheet_count}" baseType="lpstr">
      {sheet_names}
    </vt:vector>
  </TitlesOfParts>
  <Company></Company>
  <LinksUpToDate>false</LinksUpToDate>
  <SharedDoc>false</SharedDoc>
  <HyperlinksChanged>false</HyperlinksChanged>
  <AppVersion>16.0300</AppVersion>
</Properties>"#
    )
}

fn workbook_xml(workbook: &Workbook) -> String {
    let workbook_pr = match workbook.date_system {
        DateSystem::Excel1900 => r#"<workbookPr/>"#.to_string(),
        DateSystem::Excel1904 => r#"<workbookPr date1904="1"/>"#.to_string(),
    };
    let workbook_protection = workbook_protection_xml(workbook);
    let book_views = workbook_view_xml(workbook);
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
  {}
  {}
  <sheets>
    {}
  </sheets>
  {}
  {}
</workbook>"#,
        workbook_pr, workbook_protection, book_views, sheets_xml, defined_names_xml, calc_pr
    )
}

fn workbook_protection_xml(workbook: &Workbook) -> String {
    let prot = &workbook.workbook_protection;
    if formula_model::WorkbookProtection::is_default(prot) {
        return String::new();
    }
    let mut attrs = String::new();
    if prot.lock_structure {
        attrs.push_str(r#" lockStructure="1""#);
    }
    if prot.lock_windows {
        attrs.push_str(r#" lockWindows="1""#);
    }
    if let Some(hash) = prot.password_hash {
        attrs.push_str(&format!(r#" workbookPassword="{:04X}""#, hash));
    }
    format!(r#"<workbookProtection{attrs}/>"#)
}

fn workbook_view_xml(workbook: &Workbook) -> String {
    let mut attrs = String::new();

    if let Some(active_sheet_id) = workbook.view.active_sheet_id {
        if let Some(idx) = workbook.sheets.iter().position(|s| s.id == active_sheet_id) {
            if idx != 0 {
                attrs.push_str(&format!(r#" activeTab="{idx}""#));
            }
        }
    }

    if let Some(window) = workbook.view.window.as_ref().filter(|window| {
        window.x.is_some()
            || window.y.is_some()
            || window.width.is_some()
            || window.height.is_some()
            || window.state.is_some()
    }) {
        if let Some(x) = window.x {
            attrs.push_str(&format!(r#" xWindow="{x}""#));
        }
        if let Some(y) = window.y {
            attrs.push_str(&format!(r#" yWindow="{y}""#));
        }
        if let Some(width) = window.width {
            attrs.push_str(&format!(r#" windowWidth="{width}""#));
        }
        if let Some(height) = window.height {
            attrs.push_str(&format!(r#" windowHeight="{height}""#));
        }
        if let Some(state) = window.state {
            let v = match state {
                WorkbookWindowState::Normal => None,
                WorkbookWindowState::Minimized => Some("minimized"),
                WorkbookWindowState::Maximized => Some("maximized"),
            };
            if let Some(v) = v {
                attrs.push_str(&format!(r#" windowState="{v}""#));
            }
        }
    }

    if attrs.is_empty() {
        return String::new();
    }

    format!(r#"<bookViews><workbookView{attrs}/></bookViews>"#)
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
    let mut settings_by_sheet_name: HashMap<String, &SheetPrintSettings> = HashMap::new();
    for sheet_settings in &workbook.print_settings.sheets {
        if sheet_settings.sheet_name.is_empty() {
            continue;
        }
        settings_by_sheet_name.insert(
            sheet_settings.sheet_name.to_ascii_uppercase(),
            sheet_settings,
        );
    }

    let mut print_defined_names: Vec<(String, u32, String)> = Vec::new();
    for (sheet_index, sheet) in workbook.sheets.iter().enumerate() {
        let Some(settings) = settings_by_sheet_name
            .get(&sheet.name.to_ascii_uppercase())
            .copied()
        else {
            continue;
        };

        if let Some(areas) = settings.print_area.as_ref().filter(|a| !a.is_empty()) {
            let ranges: Vec<crate::print::CellRange> = areas
                .iter()
                .map(|range| crate::print::CellRange {
                    start_row: range.start.row.saturating_add(1),
                    end_row: range.end.row.saturating_add(1),
                    start_col: range.start.col.saturating_add(1),
                    end_col: range.end.col.saturating_add(1),
                })
                .collect();
            let value = crate::print::format_print_area_defined_name(&sheet.name, &ranges);
            if !value.is_empty() {
                print_defined_names.push((
                    "_xlnm.Print_Area".to_string(),
                    sheet_index as u32,
                    value,
                ));
            }
        }

        if let Some(titles) = settings
            .print_titles
            .as_ref()
            .filter(|t| t.repeat_rows.is_some() || t.repeat_cols.is_some())
        {
            let titles = crate::print::PrintTitles {
                repeat_rows: titles.repeat_rows.map(|rows| crate::print::RowRange {
                    start: rows.start.saturating_add(1),
                    end: rows.end.saturating_add(1),
                }),
                repeat_cols: titles.repeat_cols.map(|cols| crate::print::ColRange {
                    start: cols.start.saturating_add(1),
                    end: cols.end.saturating_add(1),
                }),
            };
            let value = crate::print::format_print_titles_defined_name(&sheet.name, &titles);
            if !value.is_empty() {
                print_defined_names.push((
                    "_xlnm.Print_Titles".to_string(),
                    sheet_index as u32,
                    value,
                ));
            }
        }
    }

    if workbook.defined_names.is_empty() && print_defined_names.is_empty() {
        return String::new();
    }

    let mut sheet_index_by_id = HashMap::new();
    for (idx, sheet) in workbook.sheets.iter().enumerate() {
        sheet_index_by_id.insert(sheet.id, idx as u32);
    }

    // Excel defined names are case-insensitive; normalize keys so we reliably suppress
    // duplicate built-in print names regardless of casing in `workbook.defined_names`.
    let print_keys: HashSet<(String, u32)> = print_defined_names
        .iter()
        .map(|(name, local_sheet_id, _)| (name.to_ascii_uppercase(), *local_sheet_id))
        .collect();

    let mut out = String::new();
    out.push_str("<definedNames>");
    for defined in &workbook.defined_names {
        let local_sheet_id = match defined.scope {
            DefinedNameScope::Sheet(sheet_id) => sheet_index_by_id.get(&sheet_id).copied(),
            DefinedNameScope::Workbook => None,
        };
        if let Some(local_sheet_id) = local_sheet_id {
            if print_keys.contains(&(defined.name.to_ascii_uppercase(), local_sheet_id)) {
                // Avoid emitting duplicate built-in print names; `Workbook::print_settings`
                // is the canonical representation for these.
                continue;
            }
        }

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
        if let Some(local_sheet_id) = local_sheet_id {
            out.push_str(&format!(r#" localSheetId="{}""#, local_sheet_id));
        }
        out.push('>');
        out.push_str(&escape_xml(&refers_to));
        out.push_str("</definedName>");
    }

    for (name, local_sheet_id, value) in print_defined_names {
        out.push_str(r#"<definedName"#);
        out.push_str(&format!(r#" name="{}""#, escape_xml(&name)));
        out.push_str(&format!(r#" localSheetId="{}">"#, local_sheet_id));
        out.push_str(&escape_xml(&value));
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
    next += 1;
    rels.push_str(&format!(
        r#"<Relationship Id="rId{}" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/theme" Target="theme/theme1.xml"/>"#,
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

#[derive(Clone, Debug, PartialEq)]
struct ColXmlProps {
    width: Option<f32>,
    hidden: bool,
    outline_level: u8,
    collapsed: bool,
    style_xf: Option<u32>,
}

fn render_cols(sheet: &Worksheet, outline: &Outline, style_to_xf: &HashMap<u32, u32>) -> String {
    let mut col_xml_props: BTreeMap<u32, ColXmlProps> = BTreeMap::new();

    // Column properties are stored 0-based in the model; OOXML uses 1-based indices.
    for (col0, props) in sheet.col_properties.iter() {
        let col_1 = col0.saturating_add(1);
        if col_1 == 0 || col_1 > formula_model::EXCEL_MAX_COLS {
            continue;
        }
        let style_xf = props
            .style_id
            .map(|style_id| style_to_xf.get(&style_id).copied().unwrap_or(0));
        col_xml_props.insert(
            col_1,
            ColXmlProps {
                width: props.width,
                hidden: props.hidden,
                outline_level: 0,
                collapsed: false,
                style_xf,
            },
        );
    }

    // Merge in outline metadata (levels/collapsed/hidden). Outline indices are already 1-based.
    for (col_1, entry) in outline.cols.iter() {
        if col_1 == 0 || col_1 > formula_model::EXCEL_MAX_COLS {
            continue;
        }
        if entry.level == 0 && !entry.hidden.is_hidden() && !entry.collapsed {
            continue;
        }
        col_xml_props
            .entry(col_1)
            .and_modify(|props| {
                props.outline_level = entry.level;
                props.collapsed = entry.collapsed;
                props.hidden |= entry.hidden.is_hidden();
            })
            .or_insert_with(|| ColXmlProps {
                width: None,
                hidden: entry.hidden.is_hidden(),
                outline_level: entry.level,
                collapsed: entry.collapsed,
                style_xf: None,
            });
    }

    if col_xml_props.is_empty() {
        return String::new();
    }

    let mut out = String::new();
    out.push_str("<cols>");

    let mut current: Option<(u32, u32, ColXmlProps)> = None;
    for (&col_1, props) in col_xml_props.iter() {
        let props = props.clone();
        match current.take() {
            None => current = Some((col_1, col_1, props)),
            Some((start, end, cur)) if col_1 == end + 1 && props == cur => {
                current = Some((start, col_1, cur));
            }
            Some((start, end, cur)) => {
                out.push_str(&render_col_range(start, end, &cur));
                current = Some((col_1, col_1, props));
            }
        }
    }
    if let Some((start, end, cur)) = current {
        out.push_str(&render_col_range(start, end, &cur));
    }

    out.push_str("</cols>");
    out
}

fn render_col_range(start_col_1: u32, end_col_1: u32, props: &ColXmlProps) -> String {
    let mut s = String::new();
    s.push_str(&format!(r#"<col min="{start_col_1}" max="{end_col_1}""#));
    if let Some(width) = props.width {
        s.push_str(&format!(r#" width="{width}""#));
        s.push_str(r#" customWidth="1""#);
    }
    if let Some(style_xf) = props.style_xf {
        s.push_str(&format!(r#" style="{style_xf}" customFormat="1""#));
    }
    if props.hidden {
        s.push_str(r#" hidden="1""#);
    }
    if props.outline_level > 0 {
        s.push_str(&format!(r#" outlineLevel="{}""#, props.outline_level));
    }
    if props.collapsed {
        s.push_str(r#" collapsed="1""#);
    }
    s.push_str("/>");
    s
}

fn render_conditional_formatting(sheet: &Worksheet, local_to_global_dxf: Option<&[u32]>) -> String {
    if sheet.conditional_formatting_rules.is_empty() {
        return String::new();
    }

    let mut out = String::new();
    for rule in &sheet.conditional_formatting_rules {
        let Some(cf_rule_xml) = render_cf_rule(rule, local_to_global_dxf) else {
            continue;
        };

        let sqref = rule
            .applies_to
            .iter()
            .map(|r| r.to_string())
            .collect::<Vec<_>>()
            .join(" ");
        if sqref.is_empty() {
            continue;
        }

        out.push_str(&format!(
            r#"<conditionalFormatting sqref="{}">{}</conditionalFormatting>"#,
            escape_xml(&sqref),
            cf_rule_xml
        ));
    }
    out
}

fn render_cf_rule(rule: &CfRule, local_to_global_dxf: Option<&[u32]>) -> Option<String> {
    let mut attrs = String::new();

    if let Some(id) = rule.id.as_deref() {
        attrs.push_str(&format!(r#" id="{}""#, escape_xml(id)));
    }

    attrs.push_str(&format!(r#" priority="{}""#, rule.priority));

    if rule.stop_if_true {
        attrs.push_str(r#" stopIfTrue="1""#);
    }

    // Remap per-sheet `dxf_id` to the workbook-global `dxfs` index table. Best-effort:
    // out-of-bounds indices are emitted as no `dxfId` attribute.
    let global_dxf_id = rule
        .dxf_id
        .and_then(|local| local_to_global_dxf?.get(local as usize).copied());
    if let Some(global) = global_dxf_id {
        attrs.push_str(&format!(r#" dxfId="{}""#, global));
    }

    let (type_attr, body) = match &rule.kind {
        CfRuleKind::Expression { formula } => (
            "expression",
            format!(r#"<formula>{}</formula>"#, escape_xml(formula)),
        ),
        CfRuleKind::CellIs { operator, formulas } => {
            let op = cell_is_operator_attr(*operator);
            let mut inner = String::new();
            for f in formulas {
                inner.push_str(&format!(r#"<formula>{}</formula>"#, escape_xml(f)));
            }
            attrs.push_str(&format!(r#" operator="{op}""#));
            ("cellIs", inner)
        }
        // Best-effort: skip rules we can't currently serialize.
        _ => return None,
    };

    Some(format!(r#"<cfRule type="{type_attr}"{attrs}>{body}</cfRule>"#))
}

fn cell_is_operator_attr(op: CellIsOperator) -> &'static str {
    match op {
        CellIsOperator::GreaterThan => "greaterThan",
        CellIsOperator::GreaterThanOrEqual => "greaterThanOrEqual",
        CellIsOperator::LessThan => "lessThan",
        CellIsOperator::LessThanOrEqual => "lessThanOrEqual",
        CellIsOperator::Equal => "equal",
        CellIsOperator::NotEqual => "notEqual",
        CellIsOperator::Between => "between",
        CellIsOperator::NotBetween => "notBetween",
    }
}

fn sheet_xml(
    sheet: &Worksheet,
    print_settings: Option<&SheetPrintSettings>,
    shared_strings: &SharedStrings,
    table_parts: &[(String, String)],
    style_to_xf: &HashMap<u32, u32>,
    local_to_global_dxf: Option<&[u32]>,
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

    let outline = {
        // Ensure outline-hidden flags are up to date before emitting `hidden="1"` for collapsed
        // detail rows/columns.
        let mut outline = sheet.outline.clone();
        outline.recompute_outline_hidden_rows();
        outline.recompute_outline_hidden_cols();
        outline
    };

    let fit_to_page = print_settings
        .as_ref()
        .is_some_and(|s| matches!(s.page_setup.scaling, Scaling::FitTo { .. }));

    let sheet_pr_xml = if outline != Outline::default() || fit_to_page {
        let mut out = String::new();
        out.push_str("<sheetPr>");
        if outline != Outline::default() {
            out.push_str(&format!(
                r#"<outlinePr summaryBelow="{}" summaryRight="{}" showOutlineSymbols="{}"/>"#,
                bool_attr(outline.pr.summary_below),
                bool_attr(outline.pr.summary_right),
                bool_attr(outline.pr.show_outline_symbols),
            ));
        }
        if fit_to_page {
            out.push_str(r#"<pageSetUpPr fitToPage="1"/>"#);
        }
        out.push_str("</sheetPr>");
        out
    } else {
        String::new()
    };

    let sheet_format_pr_xml = {
        let base = sheet.base_col_width;
        let default_col_width = sheet.default_col_width;
        let default_row_height = sheet.default_row_height;

        if base.is_none() && default_col_width.is_none() && default_row_height.is_none() {
            String::new()
        } else {
            let mut out = String::new();
            out.push_str("<sheetFormatPr");

            if let Some(base) = base {
                out.push_str(&format!(r#" baseColWidth="{base}""#));
            }
            if let Some(width) = default_col_width {
                out.push_str(&format!(r#" defaultColWidth="{width}""#));
            }
            if let Some(height) = default_row_height {
                out.push_str(&format!(r#" defaultRowHeight="{height}""#));
            }

            out.push_str("/>");
            out
        }
    };
    let cols_xml = render_cols(sheet, &outline, style_to_xf);

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

    let row_props_rows: Vec<u32> = sheet
        .row_properties
        .iter()
        .filter_map(|(&row, props)| {
            (props.height.is_some() || props.hidden || props.style_id.is_some()).then_some(row)
        })
        .collect();

    let outline_rows: Vec<u32> = outline
        .rows
        .iter()
        .filter_map(|(row_1, entry)| {
            if entry.level > 0 || entry.hidden.is_hidden() || entry.collapsed {
                Some(row_1.saturating_sub(1))
            } else {
                None
            }
        })
        .collect();

    // Emit rows in ascending order, streaming through the columnar table rows if present.
    let mut sheet_data = String::new();
    let mut overlay_row_idx: usize = 0;
    let mut table_row: Option<u32> = columnar.as_ref().map(|c| c.origin.row);
    let table_end_row: Option<u32> = columnar
        .as_ref()
        .map(|c| c.origin.row.saturating_add(c.rows.saturating_sub(1) as u32));
    let mut outline_row_idx: usize = 0;
    let mut row_props_row_idx: usize = 0;

    loop {
        let next_overlay_row = overlay_rows.get(overlay_row_idx).copied();
        let next_table_row = match (table_row, table_end_row) {
            (Some(r), Some(end)) if r <= end => Some(r),
            _ => None,
        };
        let next_row_props_row = row_props_rows.get(row_props_row_idx).copied();
        let next_outline_row = outline_rows.get(outline_row_idx).copied();

        let mut row_idx: Option<u32> = None;
        for candidate in [
            next_table_row,
            next_overlay_row,
            next_outline_row,
            next_row_props_row,
        ] {
            if let Some(candidate) = candidate {
                row_idx = Some(match row_idx {
                    Some(existing) => existing.min(candidate),
                    None => candidate,
                });
            }
        }
        let Some(row_idx) = row_idx else { break };

        if next_overlay_row == Some(row_idx) {
            overlay_row_idx += 1;
        }
        if next_table_row == Some(row_idx) {
            table_row = Some(row_idx + 1);
        }
        if next_outline_row == Some(row_idx) {
            outline_row_idx += 1;
        }
        if next_row_props_row == Some(row_idx) {
            row_props_row_idx += 1;
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

        let row_number = row_idx + 1;
        let outline_entry = outline.rows.entry(row_number);
        let row_props = sheet.row_properties.get(&row_idx);
        let has_row_height = row_props.is_some_and(|props| props.height.is_some());
        let row_style_id = row_props.and_then(|props| props.style_id);
        let is_row_hidden =
            outline_entry.hidden.is_hidden() || row_props.is_some_and(|props| props.hidden);
        let needs_row = wrote_any_cell
            || outline_entry.level > 0
            || is_row_hidden
            || outline_entry.collapsed
            || has_row_height
            || row_style_id.is_some();
        if !needs_row {
            continue;
        }

        let mut row_attrs = format!(r#" r="{}""#, row_number);
        if let Some(height) = row_props.and_then(|props| props.height) {
            let ht = trim_float(height as f64);
            row_attrs.push_str(&format!(r#" ht="{ht}" customHeight="1""#));
        }
        if let Some(style_id) = row_style_id {
            let xf_index = style_to_xf.get(&style_id).copied().unwrap_or(0);
            row_attrs.push_str(&format!(r#" s="{xf_index}" customFormat="1""#));
        }
        if outline_entry.level > 0 {
            row_attrs.push_str(&format!(r#" outlineLevel="{}""#, outline_entry.level));
        }
        if is_row_hidden {
            row_attrs.push_str(r#" hidden="1""#);
        }
        if outline_entry.collapsed {
            row_attrs.push_str(r#" collapsed="1""#);
        }

        if wrote_any_cell {
            sheet_data.push_str(&format!(r#"<row{row_attrs}>"#));
            sheet_data.push_str(&row_cells_xml);
            sheet_data.push_str("</row>");
        } else {
            sheet_data.push_str(&format!(r#"<row{row_attrs}/>"#));
        }
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

    let conditional_formatting_xml = render_conditional_formatting(sheet, local_to_global_dxf);

    let sheet_protection_xml = sheet_protection_xml(sheet);
    let data_validations_xml = sheet_data_validations_xml(sheet);

    let mut page_margins_xml = String::new();
    let mut page_setup_xml = String::new();
    let mut row_breaks_xml = String::new();
    let mut col_breaks_xml = String::new();
    if let Some(settings) = print_settings {
        if settings.page_setup.margins != PageMargins::default() {
            page_margins_xml = format!(
                r#"<pageMargins left="{left}" right="{right}" top="{top}" bottom="{bottom}" header="{header}" footer="{footer}"/>"#,
                left = settings.page_setup.margins.left,
                right = settings.page_setup.margins.right,
                top = settings.page_setup.margins.top,
                bottom = settings.page_setup.margins.bottom,
                header = settings.page_setup.margins.header,
                footer = settings.page_setup.margins.footer,
            );
        }

        let default_page_setup = PageSetup::default();
        if settings.page_setup.orientation != default_page_setup.orientation
            || settings.page_setup.paper_size != default_page_setup.paper_size
            || settings.page_setup.scaling != default_page_setup.scaling
        {
            let orientation = match settings.page_setup.orientation {
                formula_model::Orientation::Portrait => "portrait",
                formula_model::Orientation::Landscape => "landscape",
            };

            let mut attrs = format!(
                r#" paperSize="{}" orientation="{}""#,
                settings.page_setup.paper_size.code, orientation
            );
            match settings.page_setup.scaling {
                Scaling::Percent(pct) => {
                    attrs.push_str(&format!(r#" scale="{}""#, pct));
                }
                Scaling::FitTo { width, height } => {
                    attrs.push_str(&format!(
                        r#" fitToWidth="{}" fitToHeight="{}""#,
                        width, height
                    ));
                }
            }
            page_setup_xml = format!(r#"<pageSetup{attrs}/>"#);
        }

        row_breaks_xml = render_row_breaks_xml(&settings.manual_page_breaks);
        col_breaks_xml = render_col_breaks_xml(&settings.manual_page_breaks);
    }

    let mut xml = String::new();
    xml.push_str(r#"<?xml version="1.0" encoding="UTF-8" standalone="yes"?>"#);
    xml.push('\n');
    xml.push_str(r#"<worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">"#);
    xml.push('\n');
    if !sheet_pr_xml.is_empty() {
        xml.push_str("  ");
        xml.push_str(&sheet_pr_xml);
        xml.push('\n');
    }
    xml.push_str(&format!(r#"  <dimension ref="{dimension_ref}"/>"#));
    xml.push('\n');
    if !sheet_format_pr_xml.is_empty() {
        xml.push_str("  ");
        xml.push_str(&sheet_format_pr_xml);
        xml.push('\n');
    }
    if !cols_xml.is_empty() {
        xml.push_str("  ");
        xml.push_str(&cols_xml);
        xml.push('\n');
    }
    xml.push_str("  <sheetData>\n");
    if !sheet_data.is_empty() {
        xml.push_str("    ");
        xml.push_str(&sheet_data);
        xml.push('\n');
    }
    xml.push_str("  </sheetData>\n");
    if !sheet_protection_xml.is_empty() {
        xml.push_str("  ");
        xml.push_str(&sheet_protection_xml);
        xml.push('\n');
    }
    if !auto_filter_xml.is_empty() {
        xml.push_str("  ");
        xml.push_str(&auto_filter_xml);
        xml.push('\n');
    }
    if !conditional_formatting_xml.is_empty() {
        xml.push_str("  ");
        xml.push_str(&conditional_formatting_xml.replace('\n', "\n  "));
        xml.push('\n');
    }
    if !data_validations_xml.is_empty() {
        xml.push_str("  ");
        xml.push_str(&data_validations_xml);
        xml.push('\n');
    }
    if !page_margins_xml.is_empty() {
        xml.push_str("  ");
        xml.push_str(&page_margins_xml);
        xml.push('\n');
    }
    if !page_setup_xml.is_empty() {
        xml.push_str("  ");
        xml.push_str(&page_setup_xml);
        xml.push('\n');
    }
    if !row_breaks_xml.is_empty() {
        xml.push_str("  ");
        xml.push_str(&row_breaks_xml);
        xml.push('\n');
    }
    if !col_breaks_xml.is_empty() {
        xml.push_str("  ");
        xml.push_str(&col_breaks_xml);
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

fn render_row_breaks_xml(breaks: &ManualPageBreaks) -> String {
    if breaks.row_breaks_after.is_empty() {
        return String::new();
    }
    let count = breaks.row_breaks_after.len();
    let mut out = String::new();
    out.push_str(&format!(
        r#"<rowBreaks count="{count}" manualBreakCount="{count}">"#
    ));
    for row0 in &breaks.row_breaks_after {
        let id = row0.saturating_add(1);
        out.push_str(&format!(r#"<brk id="{id}" max="16383" man="1"/>"#));
    }
    out.push_str("</rowBreaks>");
    out
}

fn render_col_breaks_xml(breaks: &ManualPageBreaks) -> String {
    if breaks.col_breaks_after.is_empty() {
        return String::new();
    }
    let count = breaks.col_breaks_after.len();
    let mut out = String::new();
    out.push_str(&format!(
        r#"<colBreaks count="{count}" manualBreakCount="{count}">"#
    ));
    for col0 in &breaks.col_breaks_after {
        let id = col0.saturating_add(1);
        out.push_str(&format!(r#"<brk id="{id}" max="1048575" man="1"/>"#));
    }
    out.push_str("</colBreaks>");
    out
}

fn sheet_data_validation_kind_attr(kind: DataValidationKind) -> &'static str {
    match kind {
        DataValidationKind::Whole => "whole",
        DataValidationKind::Decimal => "decimal",
        DataValidationKind::List => "list",
        DataValidationKind::Date => "date",
        DataValidationKind::Time => "time",
        DataValidationKind::TextLength => "textLength",
        DataValidationKind::Custom => "custom",
    }
}

fn sheet_data_validation_operator_attr(op: DataValidationOperator) -> &'static str {
    match op {
        DataValidationOperator::Between => "between",
        DataValidationOperator::NotBetween => "notBetween",
        DataValidationOperator::Equal => "equal",
        DataValidationOperator::NotEqual => "notEqual",
        DataValidationOperator::GreaterThan => "greaterThan",
        DataValidationOperator::GreaterThanOrEqual => "greaterThanOrEqual",
        DataValidationOperator::LessThan => "lessThan",
        DataValidationOperator::LessThanOrEqual => "lessThanOrEqual",
    }
}

fn sheet_data_validation_error_style_attr(style: DataValidationErrorStyle) -> &'static str {
    match style {
        DataValidationErrorStyle::Stop => "stop",
        DataValidationErrorStyle::Warning => "warning",
        DataValidationErrorStyle::Information => "information",
    }
}

fn sheet_data_validations_xml(sheet: &Worksheet) -> String {
    if sheet.data_validations.is_empty() {
        return String::new();
    }

    let mut out = String::new();
    out.push_str(&format!(
        r#"<dataValidations count="{}">"#,
        sheet.data_validations.len()
    ));

    for assignment in &sheet.data_validations {
        let dv = &assignment.validation;

        let sqref = assignment
            .ranges
            .iter()
            .map(|r| r.to_string())
            .collect::<Vec<_>>()
            .join(" ");

        let mut attrs = String::new();
        attrs.push_str(&format!(
            r#" type="{}""#,
            sheet_data_validation_kind_attr(dv.kind)
        ));
        if let Some(op) = dv.operator {
            attrs.push_str(&format!(
                r#" operator="{}""#,
                sheet_data_validation_operator_attr(op)
            ));
        }
        if dv.allow_blank {
            attrs.push_str(r#" allowBlank="1""#);
        }
        if dv.show_input_message {
            attrs.push_str(r#" showInputMessage="1""#);
        }
        if dv.show_error_message {
            attrs.push_str(r#" showErrorMessage="1""#);
        }
        if dv.show_drop_down {
            attrs.push_str(r#" showDropDown="1""#);
        }

        if let Some(msg) = dv.input_message.as_ref() {
            if let Some(title) = msg.title.as_deref().filter(|s| !s.is_empty()) {
                attrs.push_str(&format!(r#" promptTitle="{}""#, escape_xml(title)));
            }
            if let Some(body) = msg.body.as_deref().filter(|s| !s.is_empty()) {
                attrs.push_str(&format!(r#" prompt="{}""#, escape_xml(body)));
            }
        }

        if let Some(alert) = dv.error_alert.as_ref() {
            attrs.push_str(&format!(
                r#" errorStyle="{}""#,
                sheet_data_validation_error_style_attr(alert.style)
            ));
            if let Some(title) = alert.title.as_deref().filter(|s| !s.is_empty()) {
                attrs.push_str(&format!(r#" errorTitle="{}""#, escape_xml(title)));
            }
            if let Some(body) = alert.body.as_deref().filter(|s| !s.is_empty()) {
                attrs.push_str(&format!(r#" error="{}""#, escape_xml(body)));
            }
        }

        attrs.push_str(&format!(r#" sqref="{}""#, escape_xml(&sqref)));

        out.push_str(&format!(r#"<dataValidation{attrs}>"#));

        if let Some(formula1) = normalize_formula_text(&dv.formula1) {
            let file_formula = crate::formula_text::add_xlfn_prefixes(&formula1);
            out.push_str(&format!(
                r#"<formula1>{}</formula1>"#,
                escape_xml(&file_formula)
            ));
        }
        if let Some(formula2) = dv
            .formula2
            .as_deref()
            .and_then(|s| normalize_formula_text(s))
        {
            let file_formula = crate::formula_text::add_xlfn_prefixes(&formula2);
            out.push_str(&format!(
                r#"<formula2>{}</formula2>"#,
                escape_xml(&file_formula)
            ));
        }

        out.push_str("</dataValidation>");
    }

    out.push_str("</dataValidations>");
    out
}

fn sheet_format_pr_xml(sheet: &Worksheet) -> String {
    if sheet.default_row_height.is_none()
        && sheet.default_col_width.is_none()
        && sheet.base_col_width.is_none()
    {
        return String::new();
    }

    let mut attrs = String::new();
    if let Some(base) = sheet.base_col_width {
        attrs.push_str(&format!(r#" baseColWidth="{base}""#));
    }
    if let Some(w) = sheet.default_col_width {
        attrs.push_str(&format!(r#" defaultColWidth="{w}""#));
    }
    if let Some(ht) = sheet.default_row_height {
        attrs.push_str(&format!(r#" defaultRowHeight="{ht}""#));
    }

    format!(r#"<sheetFormatPr{attrs}/>"#)
}

fn sheet_protection_xml(sheet: &Worksheet) -> String {
    let prot = &sheet.sheet_protection;
    if !prot.enabled {
        return String::new();
    }

    let mut attrs = String::new();
    attrs.push_str(r#" sheet="1""#);

    // `objects` / `scenarios` are inverse semantics: 1 means "protected".
    attrs.push_str(&format!(
        r#" objects="{}""#,
        if prot.edit_objects { 0 } else { 1 }
    ));
    attrs.push_str(&format!(
        r#" scenarios="{}""#,
        if prot.edit_scenarios { 0 } else { 1 }
    ));

    // For selectLocked/Unlocked, Excel defaults to allowing selection; set explicit 0 when disallowed.
    if !prot.select_locked_cells {
        attrs.push_str(r#" selectLockedCells="0""#);
    }
    if !prot.select_unlocked_cells {
        attrs.push_str(r#" selectUnlockedCells="0""#);
    }

    if prot.format_cells {
        attrs.push_str(r#" formatCells="1""#);
    }
    if prot.format_columns {
        attrs.push_str(r#" formatColumns="1""#);
    }
    if prot.format_rows {
        attrs.push_str(r#" formatRows="1""#);
    }
    if prot.insert_columns {
        attrs.push_str(r#" insertColumns="1""#);
    }
    if prot.insert_rows {
        attrs.push_str(r#" insertRows="1""#);
    }
    if prot.insert_hyperlinks {
        attrs.push_str(r#" insertHyperlinks="1""#);
    }
    if prot.delete_columns {
        attrs.push_str(r#" deleteColumns="1""#);
    }
    if prot.delete_rows {
        attrs.push_str(r#" deleteRows="1""#);
    }
    if prot.sort {
        attrs.push_str(r#" sort="1""#);
    }
    if prot.auto_filter {
        attrs.push_str(r#" autoFilter="1""#);
    }
    if prot.pivot_tables {
        attrs.push_str(r#" pivotTables="1""#);
    }

    if let Some(hash) = prot.password_hash {
        attrs.push_str(&format!(r#" password="{:04X}""#, hash));
    }

    format!(r#"<sheetProtection{attrs}/>"#)
}

fn sheet_format_pr_xml(sheet: &Worksheet) -> String {
    if sheet.default_row_height.is_none()
        && sheet.default_col_width.is_none()
        && sheet.base_col_width.is_none()
    {
        return String::new();
    }

    let mut attrs = String::new();
    if let Some(base) = sheet.base_col_width {
        attrs.push_str(&format!(r#" baseColWidth="{base}""#));
    }
    if let Some(width) = sheet.default_col_width {
        attrs.push_str(&format!(r#" defaultColWidth="{width}""#));
    }
    if let Some(height) = sheet.default_row_height {
        attrs.push_str(&format!(r#" defaultRowHeight="{height}""#));
    }

    format!(r#"<sheetFormatPr{attrs}/>"#)
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

    // Best-effort support for phonetic text (ruby) annotations.
    //
    // SpreadsheetML encodes phonetic strings as inline strings with `<rPh>` runs. The simple
    // exporter does not attempt to preserve full sharedStrings.xml structure, so emit inlineStr
    // when `Cell.phonetic` is present.
    if let Some(phonetic) = cell.phonetic.as_deref() {
        let base_text: Option<String> = match &cell.value {
            CellValue::String(s) => Some(s.clone()),
            CellValue::Entity(entity) => Some(entity.display_value.clone()),
            CellValue::Record(record) => Some(record_display_string(record)),
            CellValue::RichText(r) => Some(r.text.clone()),
            CellValue::Image(image) => image
                .alt_text
                .as_deref()
                .filter(|s| !s.is_empty())
                .map(|s| s.to_string()),
            _ => None,
        };

        if let Some(base_text) = base_text {
            attrs.push_str(r#" t="inlineStr""#);
            value_xml.push_str(&inline_string_with_phonetic_xml(&base_text, phonetic));
            return format!(r#"<c{}>{}</c>"#, attrs, value_xml);
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
            let key = SharedStringKey::plain(s);
            let idx = shared_strings.index.get(&key).copied().unwrap_or_default();
            value_xml.push_str(&format!(r#"<v>{}</v>"#, idx));
        }
        CellValue::Entity(entity) => {
            attrs.push_str(r#" t="s""#);
            let key = SharedStringKey::plain(&entity.display_value);
            let idx = shared_strings.index.get(&key).copied().unwrap_or_default();
            value_xml.push_str(&format!(r#"<v>{}</v>"#, idx));
        }
        CellValue::Record(record) => {
            let display = record_display_string(record);
            attrs.push_str(r#" t="s""#);
            let key = SharedStringKey::plain(&display);
            let idx = shared_strings.index.get(&key).copied().unwrap_or_default();
            value_xml.push_str(&format!(r#"<v>{}</v>"#, idx));
        }
        CellValue::Error(e) => {
            attrs.push_str(r#" t="e""#);
            value_xml.push_str(&format!(r#"<v>{}</v>"#, escape_xml(e.as_str())));
        }
        CellValue::RichText(r) => {
            attrs.push_str(r#" t="s""#);
            let key = SharedStringKey::from_rich_text(r);
            let idx = shared_strings.index.get(&key).copied().unwrap_or_default();
            value_xml.push_str(&format!(r#"<v>{}</v>"#, idx));
        }
        CellValue::Image(image) => {
            // In-cell images are not yet exported as first-class XLSX rich values. Degrade to
            // plain text when alt text is available; otherwise omit the cached value.
            if let Some(alt) = image.alt_text.as_deref().filter(|s| !s.is_empty()) {
                attrs.push_str(r#" t="s""#);
                let key = SharedStringKey::plain(alt);
                let idx = shared_strings.index.get(&key).copied().unwrap_or_default();
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
            let key = SharedStringKey::plain(s.as_ref());
            let idx = shared_strings.index.get(&key).copied().unwrap_or_default();
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
    values: crate::shared_strings::SharedStrings,
    index: HashMap<SharedStringKey, usize>,
}

fn record_display_string(record: &formula_model::RecordValue) -> String {
    record.to_string()
}

fn build_shared_strings(workbook: &Workbook) -> SharedStrings {
    let mut values = crate::shared_strings::SharedStrings::default();
    let mut index: HashMap<SharedStringKey, usize> = HashMap::new();

    for sheet in &workbook.sheets {
        for (_cell_ref, cell) in sheet.iter_cells() {
            match &cell.value {
                CellValue::String(s) => {
                    let key = SharedStringKey::plain(s);
                    if !index.contains_key(&key) {
                        let idx = values.items.len();
                        values.items.push(RichText::new(s.clone()));
                        index.insert(key, idx);
                    }
                }
                CellValue::Entity(entity) => {
                    let s = entity.display_value.clone();
                    let key = SharedStringKey::plain(&s);
                    if !index.contains_key(&key) {
                        let idx = values.items.len();
                        values.items.push(RichText::new(s.clone()));
                        index.insert(key, idx);
                    }
                }
                CellValue::Record(record) => {
                    let s = record_display_string(record);
                    let key = SharedStringKey::plain(&s);
                    if !index.contains_key(&key) {
                        let idx = values.items.len();
                        values.items.push(RichText::new(s.clone()));
                        index.insert(key, idx);
                    }
                }
                CellValue::Image(image) => {
                    if let Some(alt) = image.alt_text.as_deref().filter(|s| !s.is_empty()) {
                        let key = SharedStringKey::plain(alt);
                        if !index.contains_key(&key) {
                            let idx = values.items.len();
                            values.items.push(RichText::new(alt.to_string()));
                            index.insert(key, idx);
                        }
                    }
                }
                CellValue::RichText(r) => {
                    let key = SharedStringKey::from_rich_text(r);
                    if !index.contains_key(&key) {
                        let idx = values.items.len();
                        values.items.push(r.clone());
                        index.insert(key, idx);
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
                            let key = SharedStringKey::plain(text);
                            if !index.contains_key(&key) {
                                let idx = values.items.len();
                                values.items.push(RichText::new(text.to_string()));
                                index.insert(key, idx);
                            }
                        }
                    }
                }
            }
        }
    }

    SharedStrings { values, index }
}

#[derive(Clone, Debug, PartialEq, Eq, Hash)]
struct SharedStringKey {
    text: String,
    runs: Vec<SharedStringRunKey>,
}

#[derive(Clone, Debug, PartialEq, Eq, Hash)]
struct SharedStringRunKey {
    start: usize,
    end: usize,
    style: SharedStringRunStyleKey,
}

#[derive(Clone, Debug, PartialEq, Eq, Hash)]
struct SharedStringRunStyleKey {
    bold: Option<bool>,
    italic: Option<bool>,
    underline: Option<u8>,
    color: Option<u32>,
    font: Option<String>,
    size_100pt: Option<u16>,
}

impl SharedStringKey {
    fn plain(text: &str) -> Self {
        Self {
            text: text.to_string(),
            runs: Vec::new(),
        }
    }

    fn from_rich_text(rich: &RichText) -> Self {
        let runs = rich
            .runs
            .iter()
            .map(|run| SharedStringRunKey {
                start: run.start,
                end: run.end,
                style: SharedStringRunStyleKey {
                    bold: run.style.bold,
                    italic: run.style.italic,
                    underline: run.style.underline.map(underline_key),
                    color: run.style.color.and_then(|c| c.argb()),
                    font: run.style.font.clone(),
                    size_100pt: run.style.size_100pt,
                },
            })
            .collect();
        Self {
            text: rich.text.clone(),
            runs,
        }
    }
}

fn underline_key(underline: Underline) -> u8 {
    match underline {
        Underline::None => 0,
        Underline::Single => 1,
        Underline::Double => 2,
        Underline::SingleAccounting => 3,
        Underline::DoubleAccounting => 4,
    }
}

fn content_types_xml(
    workbook: &Workbook,
    shared_strings: &SharedStrings,
    kind: WorkbookKind,
) -> String {
    let mut overrides = String::new();
    overrides.push_str(
        r#"<Override PartName="/docProps/app.xml" ContentType="application/vnd.openxmlformats-officedocument.extended-properties+xml"/>"#,
    );
    overrides.push_str(
        r#"<Override PartName="/docProps/core.xml" ContentType="application/vnd.openxmlformats-package.core-properties+xml"/>"#,
    );
    overrides.push_str(&format!(
        r#"<Override PartName="/xl/workbook.xml" ContentType="{}"/>"#,
        kind.workbook_content_type()
    ));
    overrides.push_str(
        r#"<Override PartName="/xl/styles.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.styles+xml"/>"#,
    );
    overrides.push_str(
        r#"<Override PartName="/xl/theme/theme1.xml" ContentType="application/vnd.openxmlformats-officedocument.theme+xml"/>"#,
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

fn theme_xml(workbook: &Workbook) -> String {
    // The theme part primarily exists so downstream consumers can resolve theme-based colors used
    // in `styles.xml` (`<color theme="N" .../>`). Keep this deterministic and minimal: just the
    // `a:clrScheme` plus required placeholder font/format scheme elements.
    let theme = &workbook.theme;

    fn rgb_hex(argb: u32) -> String {
        // Theme XML uses RGB (`RRGGBB`); ignore alpha.
        format!("{:06X}", argb & 0x00FF_FFFF)
    }

    let dk1 = rgb_hex(theme.dk1.argb());
    let lt1 = rgb_hex(theme.lt1.argb());
    let dk2 = rgb_hex(theme.dk2.argb());
    let lt2 = rgb_hex(theme.lt2.argb());
    let accent1 = rgb_hex(theme.accent1.argb());
    let accent2 = rgb_hex(theme.accent2.argb());
    let accent3 = rgb_hex(theme.accent3.argb());
    let accent4 = rgb_hex(theme.accent4.argb());
    let accent5 = rgb_hex(theme.accent5.argb());
    let accent6 = rgb_hex(theme.accent6.argb());
    let hlink = rgb_hex(theme.hlink.argb());
    let fol_hlink = rgb_hex(theme.fol_hlink.argb());

    format!(
        r#"<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<a:theme xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main" name="Office Theme">
  <a:themeElements>
    <a:clrScheme name="Office">
      <a:dk1><a:srgbClr val="{dk1}"/></a:dk1>
      <a:lt1><a:srgbClr val="{lt1}"/></a:lt1>
      <a:dk2><a:srgbClr val="{dk2}"/></a:dk2>
      <a:lt2><a:srgbClr val="{lt2}"/></a:lt2>
      <a:accent1><a:srgbClr val="{accent1}"/></a:accent1>
      <a:accent2><a:srgbClr val="{accent2}"/></a:accent2>
      <a:accent3><a:srgbClr val="{accent3}"/></a:accent3>
      <a:accent4><a:srgbClr val="{accent4}"/></a:accent4>
      <a:accent5><a:srgbClr val="{accent5}"/></a:accent5>
      <a:accent6><a:srgbClr val="{accent6}"/></a:accent6>
      <a:hlink><a:srgbClr val="{hlink}"/></a:hlink>
      <a:folHlink><a:srgbClr val="{fol_hlink}"/></a:folHlink>
    </a:clrScheme>
    <a:fontScheme name="Office">
      <a:majorFont>
        <a:latin typeface="Calibri Light"/>
        <a:ea typeface=""/>
        <a:cs typeface=""/>
      </a:majorFont>
      <a:minorFont>
        <a:latin typeface="Calibri"/>
        <a:ea typeface=""/>
        <a:cs typeface=""/>
      </a:minorFont>
    </a:fontScheme>
    <a:fmtScheme name="Office">
      <a:fillStyleLst>
        <a:solidFill><a:schemeClr val="phClr"/></a:solidFill>
      </a:fillStyleLst>
      <a:lnStyleLst>
        <a:ln w="6350">
          <a:solidFill><a:schemeClr val="phClr"/></a:solidFill>
          <a:prstDash val="solid"/>
        </a:ln>
      </a:lnStyleLst>
      <a:effectStyleLst>
        <a:effectStyle><a:effectLst/></a:effectStyle>
      </a:effectStyleLst>
      <a:bgFillStyleLst>
        <a:solidFill><a:schemeClr val="phClr"/></a:solidFill>
      </a:bgFillStyleLst>
    </a:fmtScheme>
  </a:themeElements>
  <a:objectDefaults/>
  <a:extraClrSchemeLst/>
</a:theme>"#
    )
}

fn inline_string_with_phonetic_xml(base: &str, phonetic: &str) -> String {
    let len = base.chars().count();
    format!(
        r#"<is>{}<rPh sb="0" eb="{}">{}</rPh></is>"#,
        inline_string_t(base),
        len,
        inline_string_t(phonetic)
    )
}

fn inline_string_t(text: &str) -> String {
    if needs_xml_space_preserve(text) {
        format!(
            r#"<t xml:space="preserve">{}</t>"#,
            escape_xml(text)
        )
    } else {
        format!(r#"<t>{}</t>"#, escape_xml(text))
    }
}

fn needs_xml_space_preserve(text: &str) -> bool {
    text.chars()
        .next()
        .map(|c| c.is_whitespace())
        .unwrap_or(false)
        || text
            .chars()
            .last()
            .map(|c| c.is_whitespace())
            .unwrap_or(false)
}
