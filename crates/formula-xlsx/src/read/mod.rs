use std::borrow::Cow;
use std::collections::{BTreeMap, HashMap, HashSet};
#[cfg(not(target_arch = "wasm32"))]
use std::fs::File;
use std::io::{Cursor, Read, Seek, SeekFrom};
#[cfg(not(target_arch = "wasm32"))]
use std::path::Path;

use formula_engine::{parse_formula, CellAddr, ParseOptions, SerializeOptions};
use formula_model::drawings::{DrawingObject, ImageData, ImageId};
use formula_model::pivots::{PivotChartId, PivotChartModel, PivotTableId};
use formula_model::rich_text::{RichText, RichTextRunStyle, Underline};
use formula_model::Color;
use formula_model::{
    normalize_formula_text, Cell, CellRef, CellValue, Comment, CommentKind, DataValidation,
    DataValidationErrorAlert, DataValidationErrorStyle, DataValidationInputMessage,
    DataValidationKind, DataValidationOperator, DefinedNameScope, ErrorValue, Range,
    SheetProtection, SheetSelection, SheetView, SheetVisibility, Workbook, WorkbookProtection,
    WorkbookWindow, WorkbookWindowState,
};
use quick_xml::events::attributes::AttrError;
use quick_xml::events::{BytesStart, Event};
use quick_xml::Reader;
use thiserror::Error;
use zip::ZipArchive;

use crate::autofilter::{parse_worksheet_autofilter, AutoFilterParseError};
use crate::calc_settings::read_calc_settings_from_workbook_xml;
use crate::conditional_formatting::parse_worksheet_conditional_formatting_streaming;
use crate::drawings::DrawingPart;
use crate::path::{rels_for_part, resolve_target, resolve_target_candidates};
use crate::shared_strings::parse_shared_strings_xml;
use crate::sheet_metadata::parse_sheet_tab_color;
use crate::styles::StylesPart;
use crate::tables::{parse_table, TABLE_REL_TYPE};
use crate::theme::convert::to_model_theme_palette;
use crate::theme::parse_theme_palette;
use crate::WorkbookKind;
use crate::WorksheetCommentPartNames;
use crate::{parse_worksheet_hyperlinks, XlsxError};
use crate::{
    CalcPr, CellMeta, CellValueKind, DateSystem, FormulaMeta, SheetMeta, XlsxDocument, XlsxMeta,
};

mod rich_values;

const WORKBOOK_PART: &str = "xl/workbook.xml";
const WORKBOOK_RELS_PART: &str = "xl/_rels/workbook.xml.rels";
const REL_TYPE_STYLES: &str =
    "http://schemas.openxmlformats.org/officeDocument/2006/relationships/styles";
const REL_TYPE_SHARED_STRINGS: &str =
    "http://schemas.openxmlformats.org/officeDocument/2006/relationships/sharedStrings";
const REL_TYPE_METADATA: &str =
    "http://schemas.openxmlformats.org/officeDocument/2006/relationships/metadata";
const REL_TYPE_DRAWING: &str =
    "http://schemas.openxmlformats.org/officeDocument/2006/relationships/drawing";
const REL_TYPE_SHEET_METADATA: &str =
    "http://schemas.openxmlformats.org/officeDocument/2006/relationships/sheetMetadata";
const REL_TYPE_THEME: &str =
    "http://schemas.openxmlformats.org/officeDocument/2006/relationships/theme";

// Worksheet-level comment relationship types.
const REL_TYPE_COMMENTS: &str =
    "http://schemas.openxmlformats.org/officeDocument/2006/relationships/comments";
const REL_TYPE_THREADED_COMMENTS: &str =
    "http://schemas.microsoft.com/office/2017/10/relationships/threadedComment";

fn is_threaded_comment_rel_type(type_uri: &str) -> bool {
    // Excel has emitted a few variants over time; accept the canonical URI and tolerate
    // other future variants that contain "threadedComment" (mirrors the importer).
    type_uri == REL_TYPE_THREADED_COMMENTS || type_uri.contains("threadedComment")
}

// Namespace used for deterministic `Uuid::new_v5` IDs when importing pivot-related objects from an
// XLSX package.
//
// Pivot tables / charts do not have stable IDs in OOXML, so we derive them from stable part names.
const PIVOT_BINDING_NAMESPACE: PivotTableId =
    PivotTableId::from_u128(0xaa5f186245314193be90229689c7d364);

fn pivot_table_id_from_part_name(part_name: &str) -> PivotTableId {
    let part_name = part_name.strip_prefix('/').unwrap_or(part_name);
    let mut key = String::new();
    if let Some(cap) = "pivotTable:".len().checked_add(part_name.len()) {
        let _ = key.try_reserve(cap);
    }
    key.push_str("pivotTable:");
    key.push_str(part_name);
    PivotTableId::new_v5(&PIVOT_BINDING_NAMESPACE, key.as_bytes())
}

fn pivot_chart_id_from_part_name(part_name: &str) -> PivotChartId {
    let part_name = part_name.strip_prefix('/').unwrap_or(part_name);
    let mut key = String::new();
    if let Some(cap) = "pivotChart:".len().checked_add(part_name.len()) {
        let _ = key.try_reserve(cap);
    }
    key.push_str("pivotChart:");
    key.push_str(part_name);
    PivotChartId::new_v5(&PIVOT_BINDING_NAMESPACE, key.as_bytes())
}

#[derive(Debug, Error)]
pub enum ReadError {
    #[error("io error: {0}")]
    Io(#[from] std::io::Error),
    #[error("zip error: {0}")]
    Zip(#[from] zip::result::ZipError),
    #[error("xml error: {0}")]
    Xml(#[from] quick_xml::Error),
    #[error("xml attribute error: {0}")]
    XmlAttr(#[from] AttrError),
    #[error("utf-8 error: {0}")]
    Utf8(#[from] std::str::Utf8Error),
    #[error("sharedStrings.xml parse error: {0}")]
    SharedStrings(#[from] crate::shared_strings::SharedStringsError),
    #[error(transparent)]
    Styles(#[from] crate::styles::StylesPartError),
    #[error("invalid worksheet name: {0}")]
    InvalidSheetName(#[from] formula_model::SheetNameError),
    #[error(transparent)]
    Xlsx(#[from] XlsxError),
    #[error("missing required part: {0}")]
    MissingPart(&'static str),
    #[error("invalid cell reference: {0}")]
    InvalidCellRef(String),
    #[error("invalid range reference: {0}")]
    InvalidRangeRef(String),
    #[error("invalid password for encrypted workbook")]
    InvalidPassword,
    #[error("unsupported encryption: {0}")]
    UnsupportedEncryption(String),
    #[error("invalid encrypted workbook: {0}")]
    InvalidEncryptedWorkbook(String),
}

impl From<crate::calc_settings::CalcSettingsError> for ReadError {
    fn from(err: crate::calc_settings::CalcSettingsError) -> Self {
        match err {
            crate::calc_settings::CalcSettingsError::MissingPart(part) => Self::MissingPart(part),
            crate::calc_settings::CalcSettingsError::Io(err) => Self::Io(err),
            crate::calc_settings::CalcSettingsError::Xml(err) => Self::Xml(err),
            crate::calc_settings::CalcSettingsError::XmlAttr(err) => Self::XmlAttr(err),
            crate::calc_settings::CalcSettingsError::Utf8(err) => Self::Utf8(err),
        }
    }
}

impl From<crate::encrypted::EncryptedOoxmlError> for ReadError {
    fn from(err: crate::encrypted::EncryptedOoxmlError) -> Self {
        match err {
            crate::encrypted::EncryptedOoxmlError::InvalidPassword => Self::InvalidPassword,
            crate::encrypted::EncryptedOoxmlError::UnsupportedEncryption(msg) => {
                Self::UnsupportedEncryption(msg)
            }
            crate::encrypted::EncryptedOoxmlError::InvalidEncryptedWorkbook(msg) => {
                Self::InvalidEncryptedWorkbook(msg)
            }
            crate::encrypted::EncryptedOoxmlError::Io(err) => Self::Io(err),
        }
    }
}

#[cfg(not(target_arch = "wasm32"))]
pub fn load_from_path(path: impl AsRef<Path>) -> Result<XlsxDocument, ReadError> {
    // Avoid buffering the entire ZIP container into memory (the output `XlsxDocument` already
    // materializes all *inflated* part bytes with explicit limits).
    let file = File::open(path)?;
    load_from_reader(file)
}

/// Read an XLSX workbook model from in-memory bytes without materializing every
/// ZIP part into memory.
///
/// This is a lightweight alternative to [`load_from_bytes`] that only inflates
/// the parts required to build a [`formula_model::Workbook`] (workbook metadata,
/// styles, shared strings, and referenced worksheets).
pub fn read_workbook_model_from_bytes(bytes: &[u8]) -> Result<Workbook, ReadError> {
    read_workbook_model_from_reader(Cursor::new(bytes))
}

/// Read an XLSX workbook model from in-memory bytes, transparently decrypting Office
/// `EncryptedPackage` OLE wrappers when the input bytes are password-protected.
pub fn read_workbook_model_from_bytes_with_password(
    bytes: &[u8],
    password: &str,
) -> Result<Workbook, ReadError> {
    let bytes = crate::encrypted::maybe_decrypt_office_encrypted_package(bytes, password)?;
    read_workbook_model_from_bytes(bytes.as_ref())
}

/// Read an XLSX workbook model directly from a seekable reader without inflating
/// the entire XLSX package (or every ZIP part) into memory.
pub fn read_workbook_model_from_reader<R: Read + Seek>(
    mut reader: R,
) -> Result<Workbook, ReadError> {
    // Ensure we start from the beginning; callers may pass a reused reader.
    reader.seek(SeekFrom::Start(0))?;
    let mut archive = ZipArchive::new(reader)?;
    read_workbook_model_from_zip(&mut archive)
}

fn read_workbook_model_from_zip<R: Read + Seek>(
    archive: &mut ZipArchive<R>,
) -> Result<Workbook, ReadError> {
    let workbook_xml = read_zip_part_required(archive, WORKBOOK_PART)?;
    let workbook_rels = read_zip_part_required(archive, WORKBOOK_RELS_PART)?;

    let rels_info = parse_relationships(&workbook_rels)?;
    let (date_system, _calc_pr, sheets, defined_names, workbook_protection, workbook_view) =
        parse_workbook_metadata(&workbook_xml, &rels_info.id_to_target)?;
    let calc_settings = read_calc_settings_from_workbook_xml(&workbook_xml)?;

    let mut workbook = Workbook::new();
    workbook.calc_settings = calc_settings;
    workbook.date_system = match date_system {
        DateSystem::V1900 => formula_model::DateSystem::Excel1900,
        DateSystem::V1904 => formula_model::DateSystem::Excel1904,
    };
    workbook.workbook_protection = workbook_protection;

    // Best-effort: load theme palette to enable resolving theme-based colors (e.g. in styles.xml).
    if let Some(theme) = read_theme_palette_from_zip(archive, &rels_info) {
        workbook.theme = theme;
    }
    let mut worksheet_ids_by_index: Vec<formula_model::WorksheetId> = Vec::new();
    let _ = worksheet_ids_by_index.try_reserve(sheets.len());

    let styles_bytes = if let Some(target) = rels_info.styles_target.as_deref() {
        let mut found = None;
        for candidate in resolve_target_candidates(WORKBOOK_PART, target) {
            if let Some(bytes) = read_zip_part_optional(archive, &candidate)? {
                found = Some(bytes);
                break;
            }
        }
        found
    } else {
        read_zip_part_optional(archive, "xl/styles.xml")?
    };
    let styles_part = StylesPart::parse_or_default(styles_bytes.as_deref(), &mut workbook.styles)?;
    // Conditional formatting dxfs are only needed if a worksheet contains conditional
    // formatting rules. Parse them lazily to avoid unnecessary DOM parsing for workbooks
    // without conditional formatting.
    let mut conditional_formatting_dxfs: Option<Vec<formula_model::CfStyleOverride>> = None;

    let shared_strings_bytes = if let Some(target) = rels_info.shared_strings_target.as_deref() {
        let mut found = None;
        for candidate in resolve_target_candidates(WORKBOOK_PART, target) {
            if let Some(bytes) = read_zip_part_optional(archive, &candidate)? {
                found = Some(bytes);
                break;
            }
        }
        found
    } else {
        read_zip_part_optional(archive, "xl/sharedStrings.xml")?
    };
    let shared_strings = match shared_strings_bytes {
        Some(bytes) => parse_shared_strings(&bytes)?,
        None => Vec::new(),
    };

    let metadata_part_bytes = if let Some(target) = rels_info.metadata_target.as_deref() {
        let mut found = None;
        for candidate in resolve_target_candidates(WORKBOOK_PART, target) {
            if let Some(bytes) = read_zip_part_optional(archive, &candidate)? {
                found = Some(bytes);
                break;
            }
        }
        found
    } else {
        read_zip_part_optional(archive, "xl/metadata.xml")?
    };
    let metadata_part = metadata_part_bytes
        .as_deref()
        .and_then(|bytes| MetadataPart::parse(bytes).ok());

    fn person_part_name_if_matches(raw: &str) -> Option<String> {
        let name = raw.strip_prefix('/').unwrap_or(raw);
        let name: Cow<'_, str> = if name.contains('\\') {
            Cow::Owned(name.replace('\\', "/"))
        } else {
            Cow::Borrowed(name)
        };
        let name = name.as_ref();
        if !name
            .get(.."xl/persons/".len())
            .is_some_and(|p| p.eq_ignore_ascii_case("xl/persons/"))
        {
            return None;
        }
        let suffix_start = name.len().checked_sub(".xml".len())?;
        if !name
            .get(suffix_start..)
            .is_some_and(|s| s.eq_ignore_ascii_case(".xml"))
        {
            return None;
        }
        Some(name.to_string())
    }

    // Best-effort threaded comment personId -> displayName mapping. Missing/invalid parts should
    // not fail workbook load.
    let person_part_names: Vec<String> = archive
        .file_names()
        .filter_map(|name| {
            let name = name.strip_prefix('/').unwrap_or(name);
            person_part_name_if_matches(name)
        })
        .collect();
    let persons = crate::comments::import::collect_persons(
        WORKBOOK_PART,
        &workbook_rels,
        person_part_names,
        |target| {
            read_zip_part_optional(archive, target)
                .ok()
                .flatten()
                .map(Cow::Owned)
        },
    );

    for (sheet_index, sheet) in sheets.into_iter().enumerate() {
        let ws_id = workbook.add_sheet(sheet.name.clone())?;
        worksheet_ids_by_index.push(ws_id);

        {
            let Some(ws) = workbook.sheet_mut(ws_id) else {
                debug_assert!(
                    false,
                    "sheet id {ws_id:?} missing immediately after add_sheet({:?})",
                    sheet.name
                );
                return Err(ReadError::Xlsx(XlsxError::Invalid(
                    "internal error: sheet id missing immediately after add_sheet".to_string(),
                )));
            };
            ws.xlsx_sheet_id = Some(sheet.sheet_id);
            ws.xlsx_rel_id = Some(sheet.relationship_id.clone());
            ws.visibility = match sheet.state.as_deref() {
                Some("hidden") => SheetVisibility::Hidden,
                Some("veryHidden") => SheetVisibility::VeryHidden,
                _ => SheetVisibility::Visible,
            };
        }

        let sheet_part_candidates = rels_info
            .id_to_target
            .get(&sheet.relationship_id)
            .map(|target| resolve_target_candidates(WORKBOOK_PART, target))
            .unwrap_or_else(|| vec![sheet.path.clone()]);

        let mut worksheet_part = None;
        let mut sheet_xml = None;
        for candidate in sheet_part_candidates {
            if let Some(bytes) = read_zip_part_optional(archive, &candidate)? {
                worksheet_part = Some(candidate);
                sheet_xml = Some(bytes);
                break;
            }
        }
        let worksheet_part = worksheet_part.ok_or(ReadError::MissingPart(
            "worksheet part referenced from workbook.xml.rels",
        ))?;
        let Some(sheet_xml) = sheet_xml else {
            debug_assert!(
                false,
                "worksheet_part is set ({worksheet_part:?}) but sheet_xml bytes are missing"
            );
            return Err(ReadError::Xlsx(XlsxError::Invalid(
                "internal error: worksheet bytes missing after resolving worksheet part".to_string(),
            )));
        };

        // Worksheet print settings (page setup/margins, manual page breaks) live in the worksheet XML.
        // This is parsed via a streaming extractor (quick-xml) to avoid DOM parsing.
        let (page_setup, manual_breaks) = crate::print::parse_worksheet_print_settings(&sheet_xml)
            .unwrap_or_else(|_| {
                (
                    crate::print::PageSetup::default(),
                    crate::print::ManualPageBreaks::default(),
                )
            });
        let page_setup = page_setup_to_model(page_setup);
        if page_setup != formula_model::PageSetup::default() {
            workbook.set_sheet_page_setup(ws_id, page_setup);
        }
        let manual_breaks = manual_page_breaks_to_model(manual_breaks);
        if !manual_breaks.is_empty() {
            workbook.set_manual_page_breaks(ws_id, manual_breaks);
        }

        // Worksheet-level metadata lives inside the worksheet part (and sometimes its .rels).
        let sheet_xml_str = std::str::from_utf8(&sheet_xml)?;

        // Optional metadata: best-effort.
        let tab_color = parse_sheet_tab_color(sheet_xml_str).unwrap_or(None);

        // Conditional formatting: best-effort. This is parsed via a streaming extractor so we
        // don't DOM-parse the entire worksheet XML.
        let parsed_conditional_formatting =
            parse_worksheet_conditional_formatting_streaming(sheet_xml_str).ok();

        // Merged cells (must be parsed before cell content so we don't treat interior
        // cells as value-bearing).
        let merged_cells =
            crate::merge_cells::read_merge_cells_from_worksheet_xml(sheet_xml_str).ok();

        // Worksheet relationships are needed to resolve drawings, external hyperlink targets, and
        // table parts.
        let rels_part = rels_for_part(&worksheet_part);
        let rels_xml_bytes = read_zip_part_optional(archive, &rels_part)?;
        let rels_xml = rels_xml_bytes
            .as_deref()
            .and_then(|bytes| std::str::from_utf8(bytes).ok());

        // Floating drawings (images, shapes, charts) anchored to the sheet.
        let drawing_objects = load_sheet_drawings_from_archive(
            sheet_index,
            &worksheet_part,
            &sheet_xml,
            rels_xml_bytes.as_deref(),
            archive,
            &mut workbook,
        );

        let Some(ws) = workbook.sheet_mut(ws_id) else {
            debug_assert!(
                false,
                "sheet id {ws_id:?} missing immediately after add_sheet({:?})",
                sheet.name
            );
            return Err(ReadError::Xlsx(XlsxError::Invalid(
                "internal error: sheet id missing immediately after add_sheet".to_string(),
            )));
        };

        ws.tab_color = tab_color;

        if let Some(parsed) = parsed_conditional_formatting {
            if !parsed.rules.is_empty() {
                ws.conditional_formatting_rules = parsed.rules;
                let dxfs = conditional_formatting_dxfs
                    .get_or_insert_with(|| styles_part.conditional_formatting_dxfs());
                ws.conditional_formatting_dxfs = dxfs.clone();
            }
        }

        if let Some(merges) = merged_cells {
            for range in merges {
                let _ = ws.merged_regions.add(range);
            }
        }

        ws.drawings.extend(drawing_objects);

        ws.hyperlinks = parse_worksheet_hyperlinks(sheet_xml_str, rels_xml).unwrap_or_default();

        // Best-effort: comments.
        crate::comments::import::import_sheet_comments(
            ws,
            &worksheet_part,
            rels_xml_bytes.as_deref(),
            &persons,
            |target| {
                read_zip_part_optional(archive, target)
                    .ok()
                    .flatten()
                    .map(Cow::Owned)
            },
        );

        ws.auto_filter = parse_worksheet_autofilter(sheet_xml_str).ok().flatten();

        attach_tables_from_parts(
            ws,
            &worksheet_part,
            &sheet_xml,
            rels_xml_bytes.as_deref(),
            archive,
        );

        parse_worksheet_into_model(
            ws,
            ws_id,
            &sheet_xml,
            &shared_strings,
            &styles_part,
            None,
            None,
            metadata_part.as_ref(),
        )?;

        // Best-effort outline/grouping metadata.
        apply_outline_from_worksheet_xml_str(ws, sheet_xml_str);
    }

    if let Some(active_tab) = workbook_view.active_tab {
        if let Some(sheet_id) = worksheet_ids_by_index.get(active_tab).copied() {
            workbook.view.active_sheet_id = Some(sheet_id);
        }
    }
    workbook.view.window = workbook_view.window;

    for defined in defined_names {
        let scope = match defined
            .local_sheet_id
            .and_then(|idx| worksheet_ids_by_index.get(idx as usize).copied())
        {
            Some(sheet_id) => DefinedNameScope::Sheet(sheet_id),
            None => DefinedNameScope::Workbook,
        };
        // Best-effort: ignore invalid/duplicate names so we can still load the workbook.
        let _ = workbook.create_defined_name(
            scope,
            defined.name,
            defined.value,
            defined.comment,
            defined.hidden,
            defined.local_sheet_id,
        );
    }

    // Import Excel print settings (print area + print titles) from the canonical XLSX encoding:
    // workbook-scoped defined names with a `localSheetId` pointing at the sheet index.
    //
    // Best-effort: ignore malformed references so workbook load does not fail.
    populate_workbook_print_settings_from_xlsx_defined_names(&mut workbook);

    Ok(workbook)
}

fn populate_workbook_print_settings_from_xlsx_defined_names(workbook: &mut Workbook) {
    // Snapshot relevant defined names up-front so we can mutably update print settings while
    // iterating.
    let builtins: Vec<(u32, String, String)> = workbook
        .defined_names
        .iter()
        .filter(|n| {
            n.name.eq_ignore_ascii_case(formula_model::XLNM_PRINT_AREA)
                || n.name
                    .eq_ignore_ascii_case(formula_model::XLNM_PRINT_TITLES)
        })
        .filter_map(|n| {
            n.xlsx_local_sheet_id
                .map(|idx| (idx, n.name.clone(), n.refers_to.clone()))
        })
        .collect();

    for (sheet_index, name, refers_to) in builtins {
        let Some(sheet_name) = workbook
            .sheets
            .get(sheet_index as usize)
            .map(|s| s.name.clone())
        else {
            continue;
        };

        if name.eq_ignore_ascii_case(formula_model::XLNM_PRINT_AREA) {
            if workbook
                .sheet_print_settings_by_name(&sheet_name)
                .print_area
                .is_some()
            {
                continue;
            }
            let Ok(ranges) = crate::print::parse_print_area_defined_name(&sheet_name, &refers_to)
            else {
                continue;
            };
            let converted: Vec<Range> = ranges
                .into_iter()
                .filter_map(xlsx_cell_range_to_model_range)
                .collect();
            if !converted.is_empty() {
                workbook.set_sheet_print_area_by_name(&sheet_name, Some(converted));
            }
        } else if name.eq_ignore_ascii_case(formula_model::XLNM_PRINT_TITLES) {
            if workbook
                .sheet_print_settings_by_name(&sheet_name)
                .print_titles
                .is_some()
            {
                continue;
            }
            let Ok(titles) = crate::print::parse_print_titles_defined_name(&sheet_name, &refers_to)
            else {
                continue;
            };
            let Some(converted) = xlsx_print_titles_to_model_titles(titles) else {
                continue;
            };
            workbook.set_sheet_print_titles_by_name(&sheet_name, Some(converted));
        }
    }
}

fn xlsx_cell_range_to_model_range(range: crate::print::CellRange) -> Option<Range> {
    let start_row = range.start_row.checked_sub(1)?;
    let end_row = range.end_row.checked_sub(1)?;
    let start_col = range.start_col.checked_sub(1)?;
    let end_col = range.end_col.checked_sub(1)?;
    Some(Range::new(
        CellRef::new(start_row, start_col),
        CellRef::new(end_row, end_col),
    ))
}

fn xlsx_print_titles_to_model_titles(
    titles: crate::print::PrintTitles,
) -> Option<formula_model::PrintTitles> {
    let repeat_rows = titles.repeat_rows.and_then(|r| {
        Some(formula_model::RowRange {
            start: r.start.checked_sub(1)?,
            end: r.end.checked_sub(1)?,
        })
    });
    let repeat_cols = titles.repeat_cols.and_then(|c| {
        Some(formula_model::ColRange {
            start: c.start.checked_sub(1)?,
            end: c.end.checked_sub(1)?,
        })
    });

    (repeat_rows.is_some() || repeat_cols.is_some()).then_some(formula_model::PrintTitles {
        repeat_rows,
        repeat_cols,
    })
}

fn attach_tables_from_parts<R: Read + Seek>(
    worksheet: &mut formula_model::Worksheet,
    worksheet_part: &str,
    worksheet_xml: &[u8],
    worksheet_rels_xml: Option<&[u8]>,
    archive: &mut ZipArchive<R>,
) {
    attach_tables_from_part_getter(
        worksheet,
        worksheet_part,
        worksheet_xml,
        worksheet_rels_xml,
        |target| {
            read_zip_part_optional(archive, target)
                .ok()
                .flatten()
                .map(Cow::Owned)
        },
    );
}

fn attach_tables_from_part_getter<'a, F>(
    worksheet: &mut formula_model::Worksheet,
    worksheet_part: &str,
    worksheet_xml: &[u8],
    worksheet_rels_xml: Option<&[u8]>,
    mut get_part: F,
) where
    F: FnMut(&str) -> Option<Cow<'a, [u8]>>,
{
    let table_rel_ids = match parse_table_part_ids(worksheet_xml) {
        Ok(ids) => ids,
        Err(_) => Vec::new(),
    };
    if table_rel_ids.is_empty() {
        return;
    }

    let Some(rels_xml) = worksheet_rels_xml else {
        return;
    };

    let relationships = match crate::openxml::parse_relationships(rels_xml) {
        Ok(rels) => rels,
        Err(_) => return,
    };

    let mut rels_by_id: HashMap<String, crate::openxml::Relationship> = HashMap::new();
    for rel in relationships {
        rels_by_id.insert(rel.id.clone(), rel);
    }

    let mut seen_rel_ids: std::collections::HashSet<String> = worksheet
        .tables
        .iter()
        .filter_map(|t| t.relationship_id.clone())
        .collect();

    for r_id in table_rel_ids {
        if !seen_rel_ids.insert(r_id.clone()) {
            continue;
        }

        let Some(rel) = rels_by_id.get(&r_id) else {
            continue;
        };
        if rel.type_uri != TABLE_REL_TYPE {
            continue;
        }

        let mut resolved_target = None;
        let mut table_bytes = None;
        for candidate in resolve_target_candidates(worksheet_part, &rel.target) {
            if let Some(bytes) = get_part(&candidate) {
                resolved_target = Some(candidate);
                table_bytes = Some(bytes);
                break;
            }
        }
        let (target, table_bytes) = match (resolved_target, table_bytes) {
            (Some(target), Some(bytes)) => (target, bytes),
            _ => continue,
        };

        let table_xml = match std::str::from_utf8(table_bytes.as_ref()) {
            Ok(xml) => xml,
            Err(_) => continue,
        };

        let mut table = match parse_table(table_xml) {
            Ok(t) => t,
            Err(_) => continue,
        };
        table.relationship_id = Some(r_id);
        table.part_path = Some(target);
        worksheet.tables.push(table);
    }
}

fn load_sheet_drawings_from_archive<R: Read + Seek>(
    sheet_index: usize,
    sheet_part: &str,
    sheet_xml: &[u8],
    sheet_rels_xml: Option<&[u8]>,
    archive: &mut ZipArchive<R>,
    workbook: &mut Workbook,
) -> Vec<DrawingObject> {
    let drawing_rel_ids = match parse_sheet_drawing_part_ids(sheet_xml) {
        Ok(ids) => ids,
        Err(_) => Vec::new(),
    };
    if drawing_rel_ids.is_empty() {
        return Vec::new();
    }

    let Some(rels_xml) = sheet_rels_xml else {
        return Vec::new();
    };

    let relationships = match crate::openxml::parse_relationships(rels_xml) {
        Ok(rels) => rels,
        Err(_) => return Vec::new(),
    };

    let mut rels_by_id: HashMap<String, crate::openxml::Relationship> = HashMap::new();
    for rel in relationships {
        let rel_id = rel.id.trim();
        if rel_id.is_empty() {
            continue;
        }
        // Be tolerant of whitespace around `Id` (seen in some synthetic fixtures).
        if !rels_by_id.contains_key(rel_id) {
            rels_by_id.insert(rel_id.to_string(), rel);
        }
    }

    let mut objects = Vec::new();
    let mut seen_drawing_parts: HashSet<String> = HashSet::new();

    for rel_id in drawing_rel_ids {
        let rel_id = rel_id.trim();
        if rel_id.is_empty() {
            continue;
        }
        let Some(rel) = rels_by_id.get(rel_id) else {
            continue;
        };
        if rel.type_uri.trim() != REL_TYPE_DRAWING {
            continue;
        }
        if rel
            .target_mode
            .as_deref()
            .is_some_and(|mode| mode.trim().eq_ignore_ascii_case("External"))
        {
            continue;
        }

        let drawing_part = resolve_target(sheet_part, &rel.target);

        // Best-effort tolerance for invalid relationship targets:
        // - Some producers emit `Target="drawings/drawing1.xml"` (relative to `xl/`) instead of the
        //   more common `Target="../drawings/drawing1.xml"` (relative to `xl/worksheets/`).
        // - Some producers emit `Target="xl/drawings/drawing1.xml"` without a leading `/`, which
        //   incorrectly resolves under `xl/worksheets/xl/...`.
        let mut candidates: Vec<String> = Vec::new();
        candidates.push(drawing_part.clone());
        if let Some(rest) = drawing_part.strip_prefix("xl/worksheets/") {
            if rest.starts_with("drawings/") {
                candidates.push(format!("xl/{rest}"));
            } else if rest.starts_with("xl/") {
                candidates.push(rest.to_string());
            }
        }

        for candidate in candidates {
            if !seen_drawing_parts.insert(candidate.clone()) {
                continue;
            }
            if let Ok(part) = DrawingPart::parse_from_archive(sheet_index, &candidate, archive, workbook)
            {
                objects.extend(part.objects);
                break;
            }
        }
    }

    objects
}

fn load_sheet_drawings_from_parts(
    sheet_index: usize,
    sheet_part: &str,
    sheet_xml: &[u8],
    sheet_rels_xml: Option<&[u8]>,
    parts: &BTreeMap<String, Vec<u8>>,
    workbook: &mut Workbook,
) -> Vec<DrawingObject> {
    let drawing_rel_ids = match parse_sheet_drawing_part_ids(sheet_xml) {
        Ok(ids) => ids,
        Err(_) => Vec::new(),
    };
    if drawing_rel_ids.is_empty() {
        return Vec::new();
    }

    let Some(rels_xml) = sheet_rels_xml else {
        return Vec::new();
    };

    let relationships = match crate::openxml::parse_relationships(rels_xml) {
        Ok(rels) => rels,
        Err(_) => return Vec::new(),
    };

    let mut rels_by_id: HashMap<String, crate::openxml::Relationship> = HashMap::new();
    for rel in relationships {
        let rel_id = rel.id.trim();
        if rel_id.is_empty() {
            continue;
        }
        // Be tolerant of whitespace around `Id` (seen in some synthetic fixtures).
        if !rels_by_id.contains_key(rel_id) {
            rels_by_id.insert(rel_id.to_string(), rel);
        }
    }

    let mut objects = Vec::new();
    let mut seen_drawing_parts: HashSet<String> = HashSet::new();

    for rel_id in drawing_rel_ids {
        let rel_id = rel_id.trim();
        if rel_id.is_empty() {
            continue;
        }
        let Some(rel) = rels_by_id.get(rel_id) else {
            continue;
        };
        if rel.type_uri.trim() != REL_TYPE_DRAWING {
            continue;
        }
        if rel
            .target_mode
            .as_deref()
            .is_some_and(|mode| mode.trim().eq_ignore_ascii_case("External"))
        {
            continue;
        }

        let drawing_part = resolve_target(sheet_part, &rel.target);

        // Relationship targets are URIs and may be percent-encoded differently than the underlying
        // ZIP entry name (e.g. `drawing%201.xml` vs `drawing 1.xml`). Use the tolerant part-name
        // lookup so we can still locate the drawing XML in `parts`.
        let mut resolved = part_name_tolerant(parts, &drawing_part);

        // Best-effort tolerance for invalid relationship targets:
        // - `Target="drawings/drawing1.xml"` (relative to `xl/`) instead of `../drawings/...`.
        // - `Target="xl/drawings/drawing1.xml"` without a leading `/`.
        if resolved.is_none() {
            if let Some(rest) = drawing_part.strip_prefix("xl/worksheets/") {
                if rest.starts_with("drawings/") {
                    resolved = part_name_tolerant(parts, &format!("xl/{rest}"));
                } else if rest.starts_with("xl/") {
                    resolved = part_name_tolerant(parts, rest);
                }
            }
        }

        let Some(drawing_part) = resolved else {
            continue;
        };

        if !seen_drawing_parts.insert(drawing_part.clone()) {
            continue;
        }

        let parsed =
            match DrawingPart::parse_from_parts(sheet_index, &drawing_part, parts, workbook) {
                Ok(part) => part,
                Err(_) => continue,
            };
        objects.extend(parsed.objects);
    }

    objects
}

fn read_zip_part_required<R: Read + std::io::Seek>(
    archive: &mut ZipArchive<R>,
    name: &'static str,
) -> Result<Vec<u8>, ReadError> {
    read_zip_part_optional(archive, name)?.ok_or(ReadError::MissingPart(name))
}

fn read_zip_part_optional<R: Read + std::io::Seek>(
    archive: &mut ZipArchive<R>,
    name: &str,
) -> Result<Option<Vec<u8>>, ReadError> {
    crate::zip_util::read_zip_part_optional_with_limit(
        archive,
        name,
        crate::zip_util::DEFAULT_MAX_ZIP_PART_BYTES,
    )
    .map_err(ReadError::from)
}

fn part_name_tolerant(parts: &BTreeMap<String, Vec<u8>>, name: &str) -> Option<String> {
    if parts.contains_key(name) {
        return Some(name.to_string());
    }

    // Tolerate leading `/` and Windows-style separators.
    let normalized = name.trim_start_matches(|c| c == '/' || c == '\\');
    if !normalized.contains('\\') {
        if parts.contains_key(normalized) {
            return Some(normalized.to_string());
        }
        // Some producers may include a leading `/` despite this loader normalizing entries.
        let with_slash = format!("/{normalized}");
        if parts.contains_key(&with_slash) {
            return Some(with_slash);
        }
    } else {
        let normalized = normalized.replace('\\', "/");
        if parts.contains_key(normalized.as_str()) {
            return Some(normalized);
        }
        // Some producers may include a leading `/` despite this loader normalizing entries.
        let with_slash = format!("/{normalized}");
        if parts.contains_key(&with_slash) {
            return Some(with_slash);
        }
    }

    parts
        .keys()
        .find(|key| crate::zip_util::zip_part_names_equivalent(key.as_str(), name))
        .cloned()
}

fn part_bytes_tolerant<'a>(parts: &'a BTreeMap<String, Vec<u8>>, name: &str) -> Option<&'a [u8]> {
    // Fast path: exact match.
    if let Some(bytes) = parts.get(name) {
        return Some(bytes.as_slice());
    }

    // Tolerate leading `/` and Windows-style separators.
    let normalized = name.trim_start_matches(|c| c == '/' || c == '\\');
    if !normalized.contains('\\') {
        if let Some(bytes) = parts.get(normalized) {
            return Some(bytes.as_slice());
        }
        // Some producers may include a leading `/` despite this loader normalizing entries.
        let with_slash = format!("/{normalized}");
        if let Some(bytes) = parts.get(&with_slash) {
            return Some(bytes.as_slice());
        }
    } else {
        let normalized = normalized.replace('\\', "/");
        if let Some(bytes) = parts.get(normalized.as_str()) {
            return Some(bytes.as_slice());
        }
        // Some producers may include a leading `/` despite this loader normalizing entries.
        let with_slash = format!("/{normalized}");
        if let Some(bytes) = parts.get(&with_slash) {
            return Some(bytes.as_slice());
        }
    }

    // Best-effort fallback: tolerate case, path separators, leading `/`, and percent-encoded names
    // (relationship targets are URIs and may be encoded differently than the underlying ZIP entry).
    parts.iter().find_map(|(key, bytes)| {
        crate::zip_util::zip_part_names_equivalent(key, name).then_some(bytes.as_slice())
    })
}

fn detect_workbook_kind_from_content_types(xml: &[u8]) -> Option<WorkbookKind> {
    let mut reader = Reader::from_reader(xml);
    reader.config_mut().trim_text(false);
    let mut buf = Vec::new();
    loop {
        match reader.read_event_into(&mut buf).ok()? {
            Event::Start(e) | Event::Empty(e)
                if crate::openxml::local_name(e.name().as_ref())
                    .eq_ignore_ascii_case(b"Override") =>
            {
                let mut part_name = None;
                let mut content_type = None;
                for attr in e.attributes().with_checks(false).flatten() {
                    match crate::openxml::local_name(attr.key.as_ref()) {
                        b"PartName" => {
                            part_name = attr.unescape_value().ok().map(|v| v.into_owned())
                        }
                        b"ContentType" => {
                            content_type = attr.unescape_value().ok().map(|v| v.into_owned())
                        }
                        _ => {}
                    }
                }
                if part_name.as_deref() == Some("/xl/workbook.xml") {
                    return content_type
                        .as_deref()
                        .and_then(WorkbookKind::from_workbook_main_content_type);
                }
            }
            Event::Eof => break,
            _ => {}
        }
        buf.clear();
    }
    None
}

pub fn load_from_bytes(bytes: &[u8]) -> Result<XlsxDocument, ReadError> {
    let cursor = Cursor::new(bytes);
    let mut archive = ZipArchive::new(cursor)?;
    load_from_zip_archive(&mut archive)
}

/// Load an [`XlsxDocument`] from a seekable reader without first buffering the entire ZIP file.
///
/// This is equivalent to [`load_from_bytes`] but avoids allocating a `Vec<u8>` for the full
/// container (useful when opening from disk).
pub fn load_from_reader<R: Read + Seek>(mut reader: R) -> Result<XlsxDocument, ReadError> {
    // Ensure we start from the beginning; callers may pass a reused reader.
    reader.seek(SeekFrom::Start(0))?;
    let mut archive = ZipArchive::new(reader)?;
    load_from_zip_archive(&mut archive)
}

fn load_from_zip_archive<R: Read + Seek>(
    archive: &mut ZipArchive<R>,
) -> Result<XlsxDocument, ReadError> {
    let mut budget =
        crate::zip_util::ZipInflateBudget::new(crate::zip_util::DEFAULT_MAX_ZIP_TOTAL_BYTES);
    let mut parts: BTreeMap<String, Vec<u8>> = BTreeMap::new();
    for i in 0..archive.len() {
        let mut file = archive.by_index(i)?;
        if file.is_dir() {
            continue;
        }
        // ZIP entry names in valid XLSX/XLSM packages should not start with `/` (or `\`), but
        // tolerate producers that include leading separators or Windows-style path separators by
        // normalizing to canonical part names. This keeps all downstream lookups (which assume
        // `xl/...`) working.
        let name = file.name();
        let name = name.trim_start_matches(|c| c == '/' || c == '\\');
        let name = if name.contains('\\') {
            name.replace('\\', "/")
        } else {
            name.to_string()
        };
        let buf = crate::zip_util::read_zip_file_bytes_with_budget(
            &mut file,
            &name,
            crate::zip_util::DEFAULT_MAX_ZIP_PART_BYTES,
            &mut budget,
        )?;
        parts.insert(name, buf);
    }

    let workbook_kind = part_bytes_tolerant(&parts, "[Content_Types].xml")
        .and_then(detect_workbook_kind_from_content_types)
        .unwrap_or(WorkbookKind::Workbook);

    let workbook_xml =
        part_bytes_tolerant(&parts, WORKBOOK_PART).ok_or(ReadError::MissingPart(WORKBOOK_PART))?;
    let workbook_rels = part_bytes_tolerant(&parts, WORKBOOK_RELS_PART)
        .ok_or(ReadError::MissingPart(WORKBOOK_RELS_PART))?;

    let rels_info = parse_relationships(workbook_rels)?;
    let (date_system, calc_pr, sheets, defined_names, workbook_protection, workbook_view) =
        parse_workbook_metadata(workbook_xml, &rels_info.id_to_target)?;
    let calc_settings = read_calc_settings_from_workbook_xml(workbook_xml)?;

    let mut workbook = Workbook::new();
    workbook.calc_settings = calc_settings;
    workbook.date_system = match date_system {
        DateSystem::V1900 => formula_model::DateSystem::Excel1900,
        DateSystem::V1904 => formula_model::DateSystem::Excel1904,
    };
    workbook.workbook_protection = workbook_protection;

    // Best-effort: load theme palette to enable resolving theme-based colors (e.g. in styles.xml).
    if let Some(theme) = read_theme_palette_from_parts(&parts, &rels_info) {
        workbook.theme = theme;
    }
    // Conditional formatting dxfs are only needed if a worksheet contains conditional
    // formatting rules. Parse them lazily to avoid unnecessary DOM parsing for workbooks
    // without conditional formatting.
    let mut conditional_formatting_dxfs: Option<Vec<formula_model::CfStyleOverride>> = None;
    let styles_bytes = if let Some(target) = rels_info.styles_target.as_deref() {
        resolve_target_candidates(WORKBOOK_PART, target)
            .into_iter()
            .find_map(|candidate| part_bytes_tolerant(&parts, &candidate))
    } else {
        part_bytes_tolerant(&parts, "xl/styles.xml")
    };
    let styles_part = StylesPart::parse_or_default(styles_bytes, &mut workbook.styles)?;

    let shared_strings_bytes = if let Some(target) = rels_info.shared_strings_target.as_deref() {
        resolve_target_candidates(WORKBOOK_PART, target)
            .into_iter()
            .find_map(|candidate| part_bytes_tolerant(&parts, &candidate))
    } else {
        part_bytes_tolerant(&parts, "xl/sharedStrings.xml")
    };
    let shared_strings = if let Some(bytes) = shared_strings_bytes {
        parse_shared_strings(bytes)?
    } else {
        Vec::new()
    };

    let metadata_bytes = if let Some(target) = rels_info.metadata_target.as_deref() {
        resolve_target_candidates(WORKBOOK_PART, target)
            .into_iter()
            .find_map(|candidate| part_bytes_tolerant(&parts, &candidate))
    } else {
        part_bytes_tolerant(&parts, "xl/metadata.xml")
    };
    let mut metadata_part = metadata_bytes.and_then(|bytes| MetadataPart::parse(bytes).ok());
    if let Some(metadata_part) = metadata_part.as_mut() {
        metadata_part.vm_index_base = infer_vm_index_base_for_workbook(&parts, &sheets);
    }
    let mut sheet_meta: Vec<SheetMeta> = Vec::new();
    let _ = sheet_meta.try_reserve(sheets.len());
    let mut cell_meta = std::collections::HashMap::new();
    let mut rich_value_cells = std::collections::HashMap::new();
    let mut conditional_formatting = std::collections::HashMap::new();
    let mut comment_part_names: std::collections::HashMap<
        formula_model::WorksheetId,
        WorksheetCommentPartNames,
    > = std::collections::HashMap::new();
    let mut comment_snapshot: std::collections::HashMap<formula_model::WorksheetId, Vec<Comment>> =
        std::collections::HashMap::new();
    let mut drawings_snapshot: std::collections::HashMap<
        formula_model::WorksheetId,
        Vec<DrawingObject>,
    > = std::collections::HashMap::new();

    fn person_part_name_if_matches(raw: &str) -> Option<String> {
        let name = raw.strip_prefix('/').unwrap_or(raw);
        let name: Cow<'_, str> = if name.contains('\\') {
            Cow::Owned(name.replace('\\', "/"))
        } else {
            Cow::Borrowed(name)
        };
        let name = name.as_ref();
        if !name
            .get(.."xl/persons/".len())
            .is_some_and(|p| p.eq_ignore_ascii_case("xl/persons/"))
        {
            return None;
        }
        let suffix_start = name.len().checked_sub(".xml".len())?;
        if !name
            .get(suffix_start..)
            .is_some_and(|s| s.eq_ignore_ascii_case(".xml"))
        {
            return None;
        }
        Some(name.to_string())
    }

    // Best-effort threaded comment personId -> displayName mapping. Missing/invalid parts should
    // not fail workbook load.
    let person_part_names: Vec<String> = parts
        .keys()
        .filter_map(|name| {
            person_part_name_if_matches(name.as_str())
        })
        .collect();
    let persons = crate::comments::import::collect_persons(
        WORKBOOK_PART,
        workbook_rels,
        person_part_names,
        |target| part_bytes_tolerant(&parts, target).map(Cow::Borrowed),
    );

    let mut worksheet_ids_by_index: Vec<formula_model::WorksheetId> = Vec::new();
    for (sheet_index, sheet) in sheets.into_iter().enumerate() {
        let ws_id = workbook.add_sheet(sheet.name.clone())?;
        worksheet_ids_by_index.push(ws_id);

        let sheet_part_candidates = rels_info
            .id_to_target
            .get(&sheet.relationship_id)
            .map(|target| resolve_target_candidates(WORKBOOK_PART, target))
            .unwrap_or_else(|| vec![sheet.path.clone()]);
        let worksheet_part = sheet_part_candidates
            .into_iter()
            .find_map(|candidate| part_name_tolerant(&parts, &candidate))
            .ok_or(ReadError::MissingPart(
                "worksheet part referenced from workbook.xml.rels",
            ))?;
        let sheet_xml = parts
            .get(&worksheet_part)
            .map(|bytes| bytes.as_slice())
            .ok_or(ReadError::MissingPart(
                "worksheet part referenced from workbook.xml.rels",
            ))?;

        // Worksheet print settings (page setup/margins, manual page breaks) live in the worksheet
        // XML. Parse them via a streaming extractor to avoid DOM parsing.
        let (page_setup, manual_breaks) = crate::print::parse_worksheet_print_settings(sheet_xml)
            .unwrap_or_else(|_| {
                (
                    crate::print::PageSetup::default(),
                    crate::print::ManualPageBreaks::default(),
                )
            });
        let page_setup = page_setup_to_model(page_setup);
        if page_setup != formula_model::PageSetup::default() {
            workbook.set_sheet_page_setup(ws_id, page_setup);
        }
        let manual_breaks = manual_page_breaks_to_model(manual_breaks);
        if !manual_breaks.is_empty() {
            workbook.set_manual_page_breaks(ws_id, manual_breaks);
        }

        // Worksheet-level metadata lives inside the worksheet part (and sometimes its .rels).
        let sheet_xml_str = std::str::from_utf8(sheet_xml)?;

        // Worksheet relationships are needed to resolve table parts, hyperlinks, and drawings.
        let rels_part = rels_for_part(&worksheet_part);
        let rels_xml_bytes = part_bytes_tolerant(&parts, &rels_part);
        let rels_xml = rels_xml_bytes.map(std::str::from_utf8).transpose()?;

        // Comment part mapping discovered via the worksheet's `.rels`.
        // We only support *editing* existing comment infrastructure for now.
        let mut comment_parts_for_sheet =
            discover_worksheet_comment_part_names(&worksheet_part, rels_xml_bytes);
        let has_comment_parts = comment_parts_for_sheet.legacy_comments.is_some()
            || comment_parts_for_sheet.threaded_comments.is_some();
        if has_comment_parts {
            if let Some(path) = comment_parts_for_sheet.legacy_comments.clone() {
                if let Some(actual) = part_name_tolerant(&parts, &path) {
                    comment_parts_for_sheet.legacy_comments = Some(actual);
                }
            }
            if let Some(path) = comment_parts_for_sheet.threaded_comments.clone() {
                if let Some(actual) = part_name_tolerant(&parts, &path) {
                    comment_parts_for_sheet.threaded_comments = Some(actual);
                }
            }
            comment_part_names.insert(ws_id, comment_parts_for_sheet.clone());
        }

        {
            let Some(ws) = workbook.sheet_mut(ws_id) else {
                debug_assert!(
                    false,
                    "sheet id {ws_id:?} missing immediately after add_sheet({:?})",
                    sheet.name
                );
                return Err(ReadError::Xlsx(XlsxError::Invalid(
                    "internal error: sheet id missing immediately after add_sheet".to_string(),
                )));
            };

            ws.xlsx_sheet_id = Some(sheet.sheet_id);
            ws.xlsx_rel_id = Some(sheet.relationship_id.clone());
            ws.visibility = match sheet.state.as_deref() {
                Some("hidden") => SheetVisibility::Hidden,
                Some("veryHidden") => SheetVisibility::VeryHidden,
                _ => SheetVisibility::Visible,
            };

            ws.tab_color = parse_sheet_tab_color(sheet_xml_str)?;

            // Conditional formatting. Parsed via a streaming extractor so we don't DOM-parse the
            // full worksheet XML.
            let parsed_cf =
                parse_worksheet_conditional_formatting_streaming(sheet_xml_str).unwrap_or_default();
            let crate::ParsedConditionalFormatting { rules, raw_blocks } = parsed_cf;
            if !raw_blocks.is_empty() {
                conditional_formatting.insert(ws_id, raw_blocks);
            }
            if !rules.is_empty() {
                ws.conditional_formatting_rules = rules;
                let dxfs = conditional_formatting_dxfs
                    .get_or_insert_with(|| styles_part.conditional_formatting_dxfs());
                ws.conditional_formatting_dxfs = dxfs.clone();
            }
            // Merged cells must be parsed before cell content so we don't treat interior cells as
            // value-bearing, and before comments so comment anchors can normalize to the merged
            // region's top-left cell.
            let merges = crate::merge_cells::read_merge_cells_from_worksheet_xml(sheet_xml_str)
                .map_err(|err| match err {
                    crate::merge_cells::MergeCellsError::Xml(e) => ReadError::Xml(e),
                    crate::merge_cells::MergeCellsError::Attr(e) => ReadError::XmlAttr(e),
                    crate::merge_cells::MergeCellsError::Utf8(e) => ReadError::Utf8(e),
                    crate::merge_cells::MergeCellsError::InvalidRef(r) => {
                        ReadError::InvalidRangeRef(r)
                    }
                    crate::merge_cells::MergeCellsError::Zip(e) => ReadError::Zip(e),
                    crate::merge_cells::MergeCellsError::Io(e) => ReadError::Io(e),
                })?;
            for range in merges {
                ws.merged_regions
                    .add(range)
                    .map_err(|e| ReadError::InvalidRangeRef(e.to_string()))?;
            }

            ws.hyperlinks = parse_worksheet_hyperlinks(sheet_xml_str, rels_xml)?;

            // Best-effort: comments.
            crate::comments::import::import_sheet_comments(
                ws,
                &worksheet_part,
                rels_xml_bytes,
                &persons,
                |target| part_bytes_tolerant(&parts, target).map(Cow::Borrowed),
            );

            ws.auto_filter =
                parse_worksheet_autofilter(sheet_xml_str).map_err(|err| match err {
                    AutoFilterParseError::Xml(e) => ReadError::Xml(e),
                    AutoFilterParseError::Attr(e) => ReadError::XmlAttr(e),
                    AutoFilterParseError::MissingRef => ReadError::InvalidRangeRef(
                        "missing worksheet autoFilter ref attribute".to_string(),
                    ),
                    AutoFilterParseError::InvalidRef(e) => {
                        ReadError::InvalidRangeRef(e.to_string())
                    }
                })?;

            attach_tables_from_part_getter(
                ws,
                &worksheet_part,
                sheet_xml,
                rels_xml_bytes,
                |target| part_bytes_tolerant(&parts, target).map(Cow::Borrowed),
            );

            parse_worksheet_into_model(
                ws,
                ws_id,
                sheet_xml,
                &shared_strings,
                &styles_part,
                Some(&mut cell_meta),
                Some(&mut rich_value_cells),
                metadata_part.as_ref(),
            )?;

            // Best-effort outline/grouping metadata.
            apply_outline_from_worksheet_xml_str(ws, sheet_xml_str);

            expand_shared_formulas(ws, ws_id, &cell_meta);

            // Capture a normalized comment snapshot after load so the writer can detect edits.
            if has_comment_parts {
                comment_snapshot.insert(ws_id, normalize_worksheet_comments(ws));
            }
        }

        let drawing_objects = load_sheet_drawings_from_parts(
            sheet_index,
            &worksheet_part,
            sheet_xml,
            rels_xml_bytes,
            &parts,
            &mut workbook,
        );
        if !drawing_objects.is_empty() {
            if let Some(ws) = workbook.sheet_mut(ws_id) {
                ws.drawings.extend(drawing_objects);
                drawings_snapshot.insert(ws_id, ws.drawings.clone());
            }
        }

        sheet_meta.push(SheetMeta {
            worksheet_id: ws_id,
            sheet_id: sheet.sheet_id,
            relationship_id: sheet.relationship_id,
            state: sheet.state,
            path: worksheet_part,
        });
    }

    if let Some(active_tab) = workbook_view.active_tab {
        if let Some(sheet_id) = worksheet_ids_by_index.get(active_tab).copied() {
            workbook.view.active_sheet_id = Some(sheet_id);
        }
    }
    workbook.view.window = workbook_view.window;

    for defined in defined_names {
        let scope = match defined
            .local_sheet_id
            .and_then(|idx| worksheet_ids_by_index.get(idx as usize).copied())
        {
            Some(sheet_id) => DefinedNameScope::Sheet(sheet_id),
            None => DefinedNameScope::Workbook,
        };
        // Best-effort: ignore invalid/duplicate names so we can still load the workbook.
        let _ = workbook.create_defined_name(
            scope,
            defined.name,
            defined.value,
            defined.comment,
            defined.hidden,
            defined.local_sheet_id,
        );
    }

    // Import Excel print settings (print area + print titles) from workbook defined names.
    populate_workbook_print_settings_from_xlsx_defined_names(&mut workbook);

    // Best-effort in-cell image loader (`xl/cellimages*.xml`). Missing parts or media should not
    // prevent the workbook from loading.
    crate::cell_images::load_cell_images_from_parts(&parts, &mut workbook);

    // Best-effort loader for rich-value-backed images-in-cells that point directly at `xl/media/*`
    // (without a `xl/cellimages.xml` store part).
    //
    // This keeps `workbook.images` populated for real Excel workbooks that use the RichData
    // pipeline (`xl/metadata.xml` + `xl/richData/*`), even when no DrawingML cell image store part
    // exists.
    load_rich_value_images_from_parts(&parts, &mut workbook);

    // Best-effort entity/record rich value decoding. This only affects the in-memory model; the
    // underlying parts are preserved verbatim for round-trip.
    rich_values::apply_rich_values_to_workbook(&mut workbook, &rich_value_cells, &parts);

    // Best-effort pivot chart binding import. Pivot charts are stored as regular chart parts with a
    // `<c:pivotSource>` element binding them to a pivot table part.
    //
    // Missing/invalid chart parts should not prevent the workbook from loading, so this is
    // intentionally tolerant.
    let package = crate::XlsxPackage::from_parts_map(parts);
    if let Ok(chart_parts) = package.pivot_chart_parts() {
        let mut pivot_charts: Vec<PivotChartModel> = chart_parts
            .into_iter()
            .filter_map(|chart_part| {
                let pivot_source_part = chart_part.pivot_source_part?;
                let pivot_source_part = pivot_source_part
                    .strip_prefix('/')
                    .unwrap_or(pivot_source_part.as_str())
                    .to_string();
                let chart_part_name = chart_part
                    .part_name
                    .strip_prefix('/')
                    .unwrap_or(chart_part.part_name.as_str())
                    .to_string();

                Some(PivotChartModel {
                    id: pivot_chart_id_from_part_name(&chart_part_name),
                    name: chart_part
                        .pivot_source_name
                        .unwrap_or_else(|| chart_part_name.clone()),
                    pivot_table_id: pivot_table_id_from_part_name(&pivot_source_part),
                    sheet_id: None,
                    chart_part: Some(chart_part_name),
                })
            })
            .collect();

        // Best-effort: resolve which worksheet a pivot chart is placed on.
        //
        // This is relatively expensive (it scans drawing + sheet relationship parts), so only do
        // it when we actually discovered pivot chart bindings.
        if !pivot_charts.is_empty() {
            let worksheet_id_by_part: HashMap<&str, formula_model::WorksheetId> = sheet_meta
                .iter()
                .map(|meta| (meta.path.as_str(), meta.worksheet_id))
                .collect();

            if let Ok(with_placement) = package.pivot_chart_parts_with_placement() {
                let mut sheet_id_by_chart_part: HashMap<String, formula_model::WorksheetId> =
                    HashMap::new();
                for chart in with_placement {
                    if chart.chart.pivot_source_part.is_none() {
                        continue;
                    }

                    let chart_part_name = chart
                        .chart
                        .part_name
                        .strip_prefix('/')
                        .unwrap_or(chart.chart.part_name.as_str())
                        .to_string();

                    let mut candidates: HashSet<formula_model::WorksheetId> = HashSet::new();
                    for sheet_part in &chart.placed_on_sheets {
                        if let Some(id) = worksheet_id_by_part.get(sheet_part.as_str()) {
                            candidates.insert(*id);
                        }
                    }
                    if candidates.len() == 1 {
                        if let Some(sheet_id) = candidates.into_iter().next() {
                            sheet_id_by_chart_part.insert(chart_part_name, sheet_id);
                        }
                    }
                }

                for chart in &mut pivot_charts {
                    let Some(part) = chart.chart_part.as_deref() else {
                        continue;
                    };
                    if let Some(sheet_id) = sheet_id_by_chart_part.get(part) {
                        chart.sheet_id = Some(*sheet_id);
                    }
                }
            }
        }

        workbook.pivot_charts = pivot_charts;
    }
    let parts = package.into_parts_map();

    let print_settings_snapshot = workbook.print_settings.clone();

    Ok(XlsxDocument {
        workbook,
        parts,
        shared_strings,
        meta: XlsxMeta {
            date_system,
            calc_pr,
            sheets: sheet_meta,
            cell_meta,
            conditional_formatting,
            rich_value_cells,
            comment_part_names,
            comment_snapshot,
            drawings_snapshot,
            print_settings_snapshot,
        },
        calc_affecting_edits: false,
        workbook_kind,
    })
}

/// Load an [`XlsxDocument`] from in-memory bytes, transparently decrypting Office
/// `EncryptedPackage` OLE wrappers when the input bytes are password-protected.
pub fn load_from_bytes_with_password(
    bytes: &[u8],
    password: &str,
) -> Result<XlsxDocument, ReadError> {
    let bytes = crate::encrypted::maybe_decrypt_office_encrypted_package(bytes, password)?;
    load_from_bytes(bytes.as_ref())
}
fn discover_worksheet_comment_part_names(
    worksheet_part: &str,
    worksheet_rels: Option<&[u8]>,
) -> WorksheetCommentPartNames {
    let mut out = WorksheetCommentPartNames::default();
    let Some(rels_bytes) = worksheet_rels else {
        return out;
    };

    let rels = match crate::openxml::parse_relationships(rels_bytes) {
        Ok(r) => r,
        Err(_) => return out,
    };

    for rel in rels {
        if rel
            .target_mode
            .as_deref()
            .is_some_and(|mode| mode.trim().eq_ignore_ascii_case("External"))
        {
            continue;
        }

        if rel.type_uri == REL_TYPE_COMMENTS {
            if out.legacy_comments.is_none() {
                let target = resolve_target(worksheet_part, &rel.target);
                if !target.is_empty() {
                    out.legacy_comments = Some(target);
                }
            }
        } else if is_threaded_comment_rel_type(&rel.type_uri) {
            if out.threaded_comments.is_none() {
                let target = resolve_target(worksheet_part, &rel.target);
                if !target.is_empty() {
                    out.threaded_comments = Some(target);
                }
            }
        }
    }

    out
}

fn normalize_worksheet_comments(worksheet: &formula_model::Worksheet) -> Vec<Comment> {
    let mut out: Vec<Comment> = worksheet
        .iter_comments()
        .map(|(_, comment)| comment.clone())
        .collect();
    out.sort_by(|a, b| {
        (
            a.cell_ref.row,
            a.cell_ref.col,
            comment_kind_rank(a.kind),
            &a.id,
        )
            .cmp(&(
                b.cell_ref.row,
                b.cell_ref.col,
                comment_kind_rank(b.kind),
                &b.id,
            ))
    });
    out
}

fn comment_kind_rank(kind: CommentKind) -> u8 {
    match kind {
        CommentKind::Note => 0,
        CommentKind::Threaded => 1,
    }
}

fn apply_outline_from_worksheet_xml_str(worksheet: &mut formula_model::Worksheet, sheet_xml: &str) {
    let Ok(outline) = crate::outline::read_outline_from_worksheet_xml(sheet_xml) else {
        return;
    };

    worksheet.outline = outline;

    // `Worksheet::set_row_hidden` / `set_col_hidden` keep the persisted "user hidden" flag
    // (`row_properties.hidden` / `col_properties.hidden`) in sync with
    // `Worksheet::outline.{rows,cols}[*].hidden.user`.
    //
    // `parse_worksheet_into_model` currently treats `row/@hidden` and `col/@hidden` as
    // user-hidden. However, Excel also uses those attributes for rows/columns hidden by outline
    // collapse. After parsing the full outline, reconcile the persisted flags so outline-hidden
    // rows/cols are *not* treated as user-hidden.
    reconcile_user_hidden_with_outline(worksheet);
}

fn reconcile_user_hidden_with_outline(worksheet: &mut formula_model::Worksheet) {
    // Outline indices are 1-based, while row/col properties are 0-based.
    let mut rows_to_check: HashSet<u32> = HashSet::new();
    for (&row0, props) in &worksheet.row_properties {
        if props.hidden {
            rows_to_check.insert(row0);
        }
    }
    for (row1, entry) in worksheet.outline.rows.iter() {
        if entry.hidden.user && row1 > 0 {
            rows_to_check.insert(row1 - 1);
        }
    }

    let mut rows_to_check: Vec<u32> = rows_to_check.into_iter().collect();
    rows_to_check.sort_unstable();
    for row0 in rows_to_check {
        let desired = worksheet
            .outline
            .rows
            .entry(row0.saturating_add(1))
            .hidden
            .user;
        let current = worksheet
            .row_properties
            .get(&row0)
            .map(|p| p.hidden)
            .unwrap_or(false);
        if current != desired {
            worksheet.set_row_hidden(row0, desired);
        }
    }

    let mut cols_to_check: HashSet<u32> = HashSet::new();
    for (&col0, props) in &worksheet.col_properties {
        if props.hidden {
            cols_to_check.insert(col0);
        }
    }
    for (col1, entry) in worksheet.outline.cols.iter() {
        if !entry.hidden.user || col1 == 0 {
            continue;
        }
        let col0 = col1 - 1;
        if col0 < formula_model::EXCEL_MAX_COLS {
            cols_to_check.insert(col0);
        }
    }

    let mut cols_to_check: Vec<u32> = cols_to_check.into_iter().collect();
    cols_to_check.sort_unstable();
    for col0 in cols_to_check {
        let desired = worksheet
            .outline
            .cols
            .entry(col0.saturating_add(1))
            .hidden
            .user;
        let current = worksheet
            .col_properties
            .get(&col0)
            .map(|p| p.hidden)
            .unwrap_or(false);
        if current != desired {
            worksheet.set_col_hidden(col0, desired);
        }
    }
}

fn load_rich_value_images_from_parts(parts: &BTreeMap<String, Vec<u8>>, workbook: &mut Workbook) {
    let targets = match crate::rich_data::resolve_rich_value_image_targets(parts) {
        Ok(v) => v,
        Err(_) => return,
    };

    for target in targets.into_iter().flatten() {
        let Some(bytes) = part_bytes_tolerant(parts, &target) else {
            continue;
        };

        let image_id = image_id_from_target_path(&target);
        if workbook.images.get(&image_id).is_some() {
            continue;
        }

        let ext = image_id
            .as_str()
            .rsplit_once('.')
            .map(|(_, ext)| ext)
            .unwrap_or("");
        let content_type = crate::drawings::content_type_for_extension(ext).to_string();

        workbook.images.insert(
            image_id,
            ImageData {
                bytes: bytes.to_vec(),
                content_type: Some(content_type),
            },
        );
    }
}

fn image_id_from_target_path(target_path: &str) -> ImageId {
    let file_name = target_path
        .strip_prefix("xl/media/")
        .or_else(|| target_path.strip_prefix("media/"))
        .unwrap_or(target_path)
        .to_string();
    ImageId::new(file_name)
}

fn parse_relationships(bytes: &[u8]) -> Result<RelationshipsInfo, ReadError> {
    let mut reader = Reader::from_reader(bytes);
    reader.config_mut().trim_text(false);
    let mut buf = Vec::new();
    let mut id_to_target = BTreeMap::new();
    let mut styles_target = None;
    let mut shared_strings_target = None;
    let mut metadata_target = None;
    let mut sheet_metadata_target = None;
    let mut theme_target = None;
    loop {
        match reader.read_event_into(&mut buf)? {
            Event::Start(e) | Event::Empty(e)
                if crate::openxml::local_name(e.name().as_ref())
                    .eq_ignore_ascii_case(b"Relationship") =>
            {
                let mut id = None;
                let mut type_ = None;
                let mut target = None;
                let mut target_mode = None;
                for attr in e.attributes() {
                    let attr = attr?;
                    let key = crate::openxml::local_name(attr.key.as_ref());
                    if key.eq_ignore_ascii_case(b"Id") {
                        id = Some(attr.unescape_value()?.into_owned());
                    } else if key.eq_ignore_ascii_case(b"Type") {
                        type_ = Some(attr.unescape_value()?.into_owned());
                    } else if key.eq_ignore_ascii_case(b"Target") {
                        target = Some(attr.unescape_value()?.into_owned());
                    } else if key.eq_ignore_ascii_case(b"TargetMode") {
                        target_mode = Some(attr.unescape_value()?.into_owned());
                    }
                }
                if let (Some(id), Some(target)) = (id, target) {
                    if target_mode
                        .as_deref()
                        .is_some_and(|mode| mode.trim().eq_ignore_ascii_case("External"))
                    {
                        // Workbook relationship targets can be external URIs. These do not
                        // correspond to OPC part names and should not participate in the workbook
                        // part resolution graph.
                        continue;
                    }

                    if let Some(type_) = &type_ {
                        match type_.as_str() {
                            REL_TYPE_STYLES => {
                                styles_target.get_or_insert_with(|| target.clone());
                            }
                            REL_TYPE_SHARED_STRINGS => {
                                shared_strings_target.get_or_insert_with(|| target.clone());
                            }
                            REL_TYPE_METADATA => {
                                metadata_target.get_or_insert_with(|| target.clone());
                            }
                            REL_TYPE_THEME => {
                                theme_target.get_or_insert_with(|| target.clone());
                            }
                            // Modern Excel emits the metadata part using the `sheetMetadata`
                            // relationship type. Prefer this over the legacy `metadata` relationship
                            // type if both are present, since `sheetMetadata` may point at a
                            // non-default target name.
                            REL_TYPE_SHEET_METADATA => {
                                sheet_metadata_target.get_or_insert_with(|| target.clone());
                            }
                            _ => {}
                        }
                    }
                    id_to_target.insert(id, target);
                }
            }
            Event::Eof => break,
            _ => {}
        }
        buf.clear();
    }
    Ok(RelationshipsInfo {
        id_to_target,
        styles_target,
        shared_strings_target,
        metadata_target: sheet_metadata_target.or(metadata_target),
        theme_target,
    })
}

#[derive(Debug, Clone)]
struct RelationshipsInfo {
    id_to_target: BTreeMap<String, String>,
    styles_target: Option<String>,
    shared_strings_target: Option<String>,
    metadata_target: Option<String>,
    theme_target: Option<String>,
}

fn read_theme_palette_from_zip<R: Read + Seek>(
    archive: &mut ZipArchive<R>,
    rels_info: &RelationshipsInfo,
) -> Option<formula_model::ThemePalette> {
    let candidates = if let Some(target) = rels_info
        .theme_target
        .as_deref()
        .filter(|target| !target.trim().is_empty())
    {
        resolve_target_candidates(WORKBOOK_PART, target)
    } else {
        vec!["xl/theme/theme1.xml".to_string()]
    };

    for candidate in candidates {
        let Some(theme_xml) = read_zip_part_optional(archive, &candidate).ok().flatten() else {
            continue;
        };
        let Ok(palette) = parse_theme_palette(&theme_xml) else {
            continue;
        };
        return Some(to_model_theme_palette(palette));
    }

    None
}

fn read_theme_palette_from_parts(
    parts: &BTreeMap<String, Vec<u8>>,
    rels_info: &RelationshipsInfo,
) -> Option<formula_model::ThemePalette> {
    let candidates = if let Some(target) = rels_info
        .theme_target
        .as_deref()
        .filter(|target| !target.trim().is_empty())
    {
        resolve_target_candidates(WORKBOOK_PART, target)
    } else {
        vec!["xl/theme/theme1.xml".to_string()]
    };

    for candidate in candidates {
        let Some(theme_xml) = part_bytes_tolerant(parts, &candidate) else {
            continue;
        };
        let Ok(palette) = parse_theme_palette(theme_xml) else {
            continue;
        };
        return Some(to_model_theme_palette(palette));
    }

    None
}

/// Parsed representation of `xl/metadata.xml` for resolving cell `vm` attributes to rich value
/// indices (e.g. images-in-cell stored in `xl/richData/richValue.xml`).
///
/// This is intentionally best-effort; Excel occasionally emits `vm` attributes that cannot be
/// resolved. In that case we treat the cell as having no rich value.
#[derive(Debug, Clone, Default)]
struct MetadataPart {
    /// Direct mapping from worksheet `vm` indices to rich value record indices.
    ///
    /// When `xl/metadata.xml` follows the modern `metadataTypes` + `futureMetadata` form, we can
    /// resolve these indices via the richer DOM-based parser in `crate::rich_data::metadata`.
    /// The mapping is stored in this field so worksheet parsing can remain streaming.
    vm_to_rich_value: HashMap<u32, u32>,
    /// Best-effort inference of whether worksheet `c/@vm` attributes are 0-based or 1-based.
    ///
    /// Excel uses 1-based `vm` indices, but some synthetic fixtures and other producers emit
    /// 0-based indices. When we can infer the base from the workbook's worksheets, we use it to
    /// resolve `vm` deterministically without relying on lossy "map both bases into one HashMap"
    /// tricks.
    vm_index_base: VmIndexBase,
    /// Metadata type indices that appear to represent rich values.
    rich_type_indices: HashSet<u32>,
    /// Mapping of `metadataRecords` entry index -> rich value record index.
    rich_value_by_record: HashMap<u32, u32>,
    /// Rich value indices referenced via `<futureMetadata name="XLRICHVALUE">`.
    ///
    /// When present, `valueMetadata` `<rc v="...">` values often index into this list, which then
    /// points to the rich value record index (`rvb/@i`).
    future_rich_value_by_bk: Vec<Option<u32>>,
    /// Whether the document contains a `<futureMetadata name="XLRICHVALUE">` block.
    ///
    /// This is used to avoid misinterpreting `rc/@v` as a direct rich value index when we know
    /// there's an intermediate mapping layer.
    saw_future_rich_value_metadata: bool,
    /// `valueMetadata` `<bk>` blocks (potentially run-length encoded via `bk/@count`).
    ///
    /// These are referenced from worksheet cells via `c/@vm` (with potential 0/1-based ambiguity
    /// handled in [`Self::vm_to_rich_value_index`]).
    value_metadata: Vec<ValueMetadataBlock>,
    /// Best-effort inference of whether `rc/@t` values are 0-based or 1-based indices into
    /// `<metadataTypes>`.
    rc_t_base: RcIndexBase,
}

#[derive(Debug, Clone, Copy, PartialEq, Eq)]
enum VmIndexBase {
    ZeroBased,
    OneBased,
    Unknown,
}

impl Default for VmIndexBase {
    fn default() -> Self {
        Self::Unknown
    }
}

#[derive(Debug, Clone, Copy, PartialEq, Eq)]
enum RcIndexBase {
    ZeroBased,
    OneBased,
    Unknown,
}

impl Default for RcIndexBase {
    fn default() -> Self {
        Self::Unknown
    }
}

#[derive(Debug, Clone)]
struct ValueMetadataBlock {
    count: u32,
    rc_refs: Vec<(u32, u32)>,
}

impl Default for ValueMetadataBlock {
    fn default() -> Self {
        Self {
            count: 1,
            rc_refs: Vec::new(),
        }
    }
}

impl MetadataPart {
    fn parse(xml: &[u8]) -> Result<Self, ReadError> {
        if let Ok(parsed) =
            crate::rich_data::metadata::parse_value_metadata_vm_to_rich_value_index_map(xml)
        {
            if !parsed.is_empty() {
                // The DOM-based rich value metadata parser returns a mapping keyed by 1-based `vm`
                // indices (matching modern Excel). Base ambiguity for worksheet `c/@vm` values is
                // handled later via `vm_index_base` inference, so store the canonical mapping as-is
                // to avoid lossy collisions (e.g. `vm=2` and `vm=1` competing for the same key when
                // we try to represent both 0-based and 1-based schemes in one `HashMap`).
                return Ok(Self {
                    vm_to_rich_value: parsed,
                    ..Default::default()
                });
            }
        }

        let mut reader = Reader::from_reader(xml);
        reader.config_mut().trim_text(true);
        let mut buf = Vec::new();

        let mut rich_type_indices: HashSet<u32> = HashSet::new();
        let mut rich_value_by_record: HashMap<u32, u32> = HashMap::new();
        let mut future_rich_value_by_bk: Vec<Option<u32>> = Vec::new();
        let mut saw_future_rich_value_metadata = false;
        let mut value_metadata: Vec<ValueMetadataBlock> = Vec::new();

        let mut in_metadata_types = false;
        let mut next_metadata_type_idx: u32 = 0;

        let mut in_metadata_records = false;
        let mut in_mdr = false;
        let mut next_mdr_idx: u32 = 0;
        let mut current_mdr_idx: Option<u32> = None;

        let mut in_future_metadata = false;
        let mut current_future_is_rich_value = false;
        let mut in_future_bk = false;
        let mut current_future_rich_idx: Option<u32> = None;

        let mut in_value_metadata = false;
        let mut in_bk = false;
        let mut current_bk_count: u32 = 1;
        let mut current_bk: Vec<(u32, u32)> = Vec::new();

        loop {
            match reader.read_event_into(&mut buf)? {
                Event::Start(e) if e.local_name().as_ref() == b"metadataTypes" => {
                    in_metadata_types = true
                }
                Event::End(e) if e.local_name().as_ref() == b"metadataTypes" => {
                    in_metadata_types = false
                }
                Event::Start(e) | Event::Empty(e)
                    if in_metadata_types && e.local_name().as_ref() == b"metadataType" =>
                {
                    let mut name: Option<String> = None;
                    for attr in e.attributes() {
                        let attr = attr?;
                        if crate::openxml::local_name(attr.key.as_ref()) == b"name" {
                            name = Some(attr.unescape_value()?.into_owned());
                        }
                    }

                    if let Some(name) = name {
                        if crate::ascii::contains_ignore_case(&name, "richvalue")
                            || crate::ascii::contains_ignore_case(&name, "rich_value")
                            || crate::ascii::contains_ignore_case(&name, "richdata")
                            || crate::ascii::contains_ignore_case(&name, "rich")
                        {
                            rich_type_indices.insert(next_metadata_type_idx);
                        }
                    }
                    next_metadata_type_idx = next_metadata_type_idx.saturating_add(1);
                }

                Event::Start(e) if e.local_name().as_ref() == b"metadataRecords" => {
                    in_metadata_records = true
                }
                Event::End(e) if e.local_name().as_ref() == b"metadataRecords" => {
                    in_metadata_records = false;
                    in_mdr = false;
                    current_mdr_idx = None;
                }
                Event::Start(e) if in_metadata_records && e.local_name().as_ref() == b"mdr" => {
                    in_mdr = true;
                    current_mdr_idx = Some(next_mdr_idx);
                    next_mdr_idx = next_mdr_idx.saturating_add(1);
                    drop(e);
                }
                Event::Empty(e) if in_metadata_records && e.local_name().as_ref() == b"mdr" => {
                    // Empty metadata record; still consume an index.
                    next_mdr_idx = next_mdr_idx.saturating_add(1);
                    drop(e);
                }
                Event::End(e) if in_metadata_records && e.local_name().as_ref() == b"mdr" => {
                    in_mdr = false;
                    current_mdr_idx = None;
                    drop(e);
                }
                Event::Start(e) | Event::Empty(e) if in_mdr => {
                    let Some(record_idx) = current_mdr_idx else {
                        continue;
                    };

                    let local_name = e.local_name();
                    let local = local_name.as_ref();
                    let looks_like_rich_value = local == b"rvb"
                        || local == b"richValue"
                        || local == b"richvalue"
                        || local == b"rv";
                    if !looks_like_rich_value {
                        continue;
                    }

                    let mut rich_idx: Option<u32> = None;
                    for attr in e.attributes() {
                        let attr = attr?;
                        let key = crate::openxml::local_name(attr.key.as_ref());
                        if key == b"i" || key == b"idx" || key == b"index" || key == b"v" {
                            rich_idx = attr.unescape_value()?.trim().parse::<u32>().ok();
                            if rich_idx.is_some() {
                                break;
                            }
                        }
                    }

                    if let Some(rich_idx) = rich_idx {
                        rich_value_by_record.entry(record_idx).or_insert(rich_idx);
                    }
                }

                Event::Start(e) if e.local_name().as_ref() == b"futureMetadata" => {
                    in_future_metadata = true;
                    current_future_is_rich_value = false;

                    let mut name: Option<String> = None;
                    for attr in e.attributes() {
                        let attr = attr?;
                        if crate::openxml::local_name(attr.key.as_ref()) == b"name" {
                            name = Some(attr.unescape_value()?.into_owned());
                        }
                    }

                    if let Some(name) = name {
                        if name.eq_ignore_ascii_case("XLRICHVALUE") {
                            current_future_is_rich_value = true;
                            saw_future_rich_value_metadata = true;
                        }
                    }
                }
                Event::End(e) if e.local_name().as_ref() == b"futureMetadata" => {
                    in_future_metadata = false;
                    current_future_is_rich_value = false;
                    in_future_bk = false;
                    current_future_rich_idx = None;
                }
                Event::Start(e)
                    if in_future_metadata
                        && current_future_is_rich_value
                        && e.local_name().as_ref() == b"bk" =>
                {
                    in_future_bk = true;
                    current_future_rich_idx = None;
                    drop(e);
                }
                Event::Empty(e)
                    if in_future_metadata
                        && current_future_is_rich_value
                        && e.local_name().as_ref() == b"bk" =>
                {
                    future_rich_value_by_bk.push(None);
                    drop(e);
                }
                Event::End(e)
                    if in_future_metadata
                        && current_future_is_rich_value
                        && in_future_bk
                        && e.local_name().as_ref() == b"bk" =>
                {
                    future_rich_value_by_bk.push(current_future_rich_idx.take());
                    in_future_bk = false;
                    drop(e);
                }
                Event::Start(e) | Event::Empty(e)
                    if in_future_metadata && current_future_is_rich_value && in_future_bk =>
                {
                    // Look for the rich-value index (`rvb/@i`) inside the future metadata block.
                    //
                    // Prefix/namespace can vary (`xlrd:rvb`, `rvb`, etc.), so match by local name.
                    if e.local_name().as_ref() != b"rvb" {
                        continue;
                    }
                    if current_future_rich_idx.is_some() {
                        continue;
                    }

                    for attr in e.attributes() {
                        let attr = attr?;
                        let key = crate::openxml::local_name(attr.key.as_ref());
                        if key == b"i" || key == b"idx" || key == b"index" || key == b"v" {
                            current_future_rich_idx =
                                attr.unescape_value()?.trim().parse::<u32>().ok();
                            if current_future_rich_idx.is_some() {
                                break;
                            }
                        }
                    }
                }

                Event::Start(e) if e.local_name().as_ref() == b"valueMetadata" => {
                    in_value_metadata = true;
                    drop(e);
                }
                Event::End(e) if e.local_name().as_ref() == b"valueMetadata" => {
                    in_value_metadata = false;
                    in_bk = false;
                    current_bk_count = 1;
                    current_bk.clear();
                    drop(e);
                }

                Event::Start(e) if in_value_metadata && e.local_name().as_ref() == b"bk" => {
                    in_bk = true;
                    current_bk_count = 1;
                    for attr in e.attributes() {
                        let attr = attr?;
                        if crate::openxml::local_name(attr.key.as_ref()) == b"count" {
                            current_bk_count = attr
                                .unescape_value()?
                                .trim()
                                .parse::<u32>()
                                .ok()
                                .filter(|v| *v >= 1)
                                .unwrap_or(1);
                        }
                    }
                    current_bk.clear();
                    drop(e);
                }
                Event::Empty(e) if in_value_metadata && e.local_name().as_ref() == b"bk" => {
                    let mut count: u32 = 1;
                    for attr in e.attributes() {
                        let attr = attr?;
                        if crate::openxml::local_name(attr.key.as_ref()) == b"count" {
                            count = attr
                                .unescape_value()?
                                .trim()
                                .parse::<u32>()
                                .ok()
                                .filter(|v| *v >= 1)
                                .unwrap_or(1);
                        }
                    }
                    value_metadata.push(ValueMetadataBlock {
                        count,
                        rc_refs: Vec::new(),
                    });
                    drop(e);
                }
                Event::End(e) if in_bk && e.local_name().as_ref() == b"bk" => {
                    value_metadata.push(ValueMetadataBlock {
                        count: current_bk_count,
                        rc_refs: std::mem::take(&mut current_bk),
                    });
                    in_bk = false;
                    current_bk_count = 1;
                    drop(e);
                }
                Event::Start(e) | Event::Empty(e) if in_bk && e.local_name().as_ref() == b"rc" => {
                    let mut t: Option<u32> = None;
                    let mut v: Option<u32> = None;
                    for attr in e.attributes() {
                        let attr = attr?;
                        match crate::openxml::local_name(attr.key.as_ref()) {
                            b"t" => {
                                t = attr.unescape_value()?.trim().parse::<u32>().ok();
                            }
                            b"v" => {
                                v = attr.unescape_value()?.trim().parse::<u32>().ok();
                            }
                            _ => {}
                        }
                    }
                    if let (Some(t), Some(v)) = (t, v) {
                        current_bk.push((t, v));
                    }
                }

                Event::Eof => break,
                _ => {}
            }
            buf.clear();
        }

        let metadata_type_count = next_metadata_type_idx;

        let mut saw_t_zero = false;
        let mut saw_t_eq_count = false;
        if metadata_type_count > 0 {
            for bk in &value_metadata {
                for (t, _) in &bk.rc_refs {
                    if *t == 0 {
                        saw_t_zero = true;
                    }
                    if *t == metadata_type_count {
                        saw_t_eq_count = true;
                    }
                }
            }
        }

        let rc_t_base = if saw_t_zero {
            RcIndexBase::ZeroBased
        } else if saw_t_eq_count {
            RcIndexBase::OneBased
        } else {
            RcIndexBase::Unknown
        };

        Ok(Self {
            vm_to_rich_value: HashMap::new(),
            vm_index_base: VmIndexBase::Unknown,
            rich_type_indices,
            rich_value_by_record,
            future_rich_value_by_bk,
            saw_future_rich_value_metadata,
            value_metadata,
            rc_t_base,
        })
    }

    fn vm_to_rich_value_index(&self, vm: u32) -> Option<u32> {
        if !self.vm_to_rich_value.is_empty() {
            // The DOM-based rich value metadata parser returns a mapping keyed by 1-based `vm`
            // indices (matching modern Excel). Some producers emit 0-based `vm` values in worksheet
            // cells. Prefer the inferred base when available, but keep a fallback to the other
            // interpretation for resilience.
            let one_based = self.vm_to_rich_value.get(&vm).copied();
            let zero_based = vm
                .checked_add(1)
                .and_then(|vm1| self.vm_to_rich_value.get(&vm1).copied());
            return match self.vm_index_base {
                VmIndexBase::ZeroBased => zero_based.or(one_based),
                VmIndexBase::OneBased => one_based.or(zero_based),
                VmIndexBase::Unknown => one_based.or(zero_based),
            };
        }

        self.vm_to_rich_value_index_with_candidate(vm).or_else(|| {
            vm.checked_sub(1)
                .and_then(|vm| self.vm_to_rich_value_index_with_candidate(vm))
        })
    }

    fn t_matches_rich_type(&self, t: u32) -> bool {
        if self.rich_type_indices.is_empty() {
            return true;
        }

        match self.rc_t_base {
            RcIndexBase::ZeroBased => self.rich_type_indices.contains(&t),
            RcIndexBase::OneBased => t
                .checked_sub(1)
                .is_some_and(|t0| self.rich_type_indices.contains(&t0)),
            RcIndexBase::Unknown => {
                self.rich_type_indices.contains(&t)
                    || t.checked_sub(1)
                        .is_some_and(|t0| self.rich_type_indices.contains(&t0))
            }
        }
    }

    fn resolve_rich_value_index(&self, v: u32) -> Option<u32> {
        // 1) metadataRecords-style indirection:
        if let Some(idx) = self.rich_value_by_record.get(&v) {
            return Some(*idx);
        }
        // Some files appear to use 1-based indices into `metadataRecords`.
        if v > 0 {
            if let Some(idx) = self.rich_value_by_record.get(&(v - 1)) {
                return Some(*idx);
            }
        }

        // 2) futureMetadata-style indirection:
        if let Some(Some(idx)) = self.future_rich_value_by_bk.get(v as usize) {
            return Some(*idx);
        }
        // Some files appear to use 1-based indices into `futureMetadata` `<bk>` lists.
        if v > 0 {
            if let Some(Some(idx)) = self.future_rich_value_by_bk.get((v - 1) as usize) {
                return Some(*idx);
            }
        }

        None
    }

    fn vm_to_rich_value_index_with_candidate(&self, vm_idx: u32) -> Option<u32> {
        let rc_refs = self.value_metadata_rc_refs(vm_idx)?;

        // Prefer record references whose metadata type looks like a rich value type, but fall back
        // to any record if we don't know the type.
        for pass in 0..2 {
            for (t, v) in rc_refs {
                if pass == 0 && !self.t_matches_rich_type(*t) {
                    continue;
                }

                if let Some(idx) = self.resolve_rich_value_index(*v) {
                    return Some(idx);
                }

                // If we couldn't resolve via `metadataRecords`/`futureMetadata`, avoid guessing
                // unless we have no structured mapping at all.
                if self.rich_value_by_record.is_empty() && !self.saw_future_rich_value_metadata {
                    return Some(*v);
                }
            }

            if self.rich_type_indices.is_empty() {
                break;
            }
        }

        None
    }

    fn value_metadata_rc_refs(&self, vm_idx: u32) -> Option<&[(u32, u32)]> {
        self.value_metadata_block_by_vm_index(vm_idx)
            .map(|block| block.rc_refs.as_slice())
    }

    fn value_metadata_block_by_vm_index(&self, vm_idx: u32) -> Option<&ValueMetadataBlock> {
        // Excel emits `c/@vm` indices as 1-based, but some producers use 0-based indices. Prefer
        // the inferred workbook vm base when available, while still allowing a fallback to the
        // other interpretation for resilience.
        match self.vm_index_base {
            VmIndexBase::ZeroBased => self
                .value_metadata_block_by_vm_index_candidate(vm_idx)
                .or_else(|| {
                    vm_idx
                        .checked_sub(1)
                        .and_then(|idx| self.value_metadata_block_by_vm_index_candidate(idx))
                }),
            VmIndexBase::OneBased | VmIndexBase::Unknown => vm_idx
                .checked_sub(1)
                .and_then(|idx| self.value_metadata_block_by_vm_index_candidate(idx))
                .or_else(|| self.value_metadata_block_by_vm_index_candidate(vm_idx)),
        }
    }

    fn value_metadata_block_by_vm_index_candidate(
        &self,
        vm_idx: u32,
    ) -> Option<&ValueMetadataBlock> {
        // Walk blocks cumulatively to support run-length encoding via `bk/@count`.
        let mut cursor: u32 = 0;
        for block in &self.value_metadata {
            let count = block.count.max(1);
            let end = cursor.saturating_add(count);
            if vm_idx < end {
                return Some(block);
            }
            cursor = end;
        }
        None
    }
}

fn infer_vm_index_base_for_workbook(
    parts: &BTreeMap<String, Vec<u8>>,
    sheets: &[ParsedSheet],
) -> VmIndexBase {
    for sheet in sheets {
        let Some(bytes) = part_bytes_tolerant(parts, &sheet.path) else {
            continue;
        };
        if worksheet_contains_vm_zero(bytes) {
            return VmIndexBase::ZeroBased;
        }
    }
    VmIndexBase::OneBased
}

fn worksheet_contains_vm_zero(xml: &[u8]) -> bool {
    let mut reader = Reader::from_reader(xml);
    reader.config_mut().trim_text(true);
    let mut buf = Vec::new();

    loop {
        match reader.read_event_into(&mut buf) {
            Ok(Event::Start(e) | Event::Empty(e)) if e.local_name().as_ref() == b"c" => {
                for attr in e.attributes() {
                    let Ok(attr) = attr else {
                        continue;
                    };
                    if crate::openxml::local_name(attr.key.as_ref()) != b"vm" {
                        continue;
                    }
                    let Ok(val) = attr.unescape_value() else {
                        continue;
                    };
                    if val.trim() == "0" {
                        return true;
                    }
                }
            }
            Ok(Event::Eof) => break,
            Ok(_) => {}
            Err(_) => break,
        }
        buf.clear();
    }

    false
}

fn parse_shared_strings(bytes: &[u8]) -> Result<Vec<RichText>, ReadError> {
    let xml = std::str::from_utf8(bytes)?;
    let parsed = parse_shared_strings_xml(xml)?;
    Ok(parsed.items)
}

#[derive(Debug, Clone)]
struct ParsedSheet {
    name: String,
    sheet_id: u32,
    relationship_id: String,
    state: Option<String>,
    path: String,
}

#[derive(Debug, Clone)]
struct ParsedDefinedName {
    name: String,
    local_sheet_id: Option<u32>,
    comment: Option<String>,
    hidden: bool,
    value: String,
}

#[derive(Debug, Clone, Default)]
struct ParsedWorkbookView {
    active_tab: Option<usize>,
    window: Option<WorkbookWindow>,
}

fn parse_workbook_metadata(
    workbook_xml: &[u8],
    rels: &BTreeMap<String, String>,
) -> Result<
    (
        DateSystem,
        CalcPr,
        Vec<ParsedSheet>,
        Vec<ParsedDefinedName>,
        WorkbookProtection,
        ParsedWorkbookView,
    ),
    ReadError,
> {
    let mut reader = Reader::from_reader(workbook_xml);
    reader.config_mut().trim_text(false);
    let mut buf = Vec::new();

    let mut date_system = DateSystem::V1900;
    let mut calc_pr = CalcPr::default();
    let mut workbook_protection = WorkbookProtection::default();
    let mut workbook_view = ParsedWorkbookView::default();
    let mut sheets = Vec::new();
    let mut defined_names = Vec::new();
    let mut current_defined: Option<ParsedDefinedName> = None;

    loop {
        match reader.read_event_into(&mut buf)? {
            Event::Start(e) | Event::Empty(e) if e.local_name().as_ref() == b"workbookPr" => {
                for attr in e.attributes() {
                    let attr = attr?;
                    if attr.key.as_ref() == b"date1904" {
                        let val = attr.unescape_value()?;
                        if parse_xml_bool(&val) {
                            date_system = DateSystem::V1904;
                        }
                    }
                }
            }
            Event::Start(e) | Event::Empty(e) if e.local_name().as_ref() == b"calcPr" => {
                for attr in e.attributes() {
                    let attr = attr?;
                    match attr.key.as_ref() {
                        b"calcId" => calc_pr.calc_id = Some(attr.unescape_value()?.into_owned()),
                        b"calcMode" => {
                            calc_pr.calc_mode = Some(attr.unescape_value()?.into_owned())
                        }
                        b"fullCalcOnLoad" => {
                            let v = attr.unescape_value()?;
                            calc_pr.full_calc_on_load = Some(parse_xml_bool(&v))
                        }
                        _ => {}
                    }
                }
            }
            Event::Start(e) | Event::Empty(e)
                if e.local_name().as_ref() == b"workbookProtection" =>
            {
                for attr in e.attributes() {
                    let attr = attr?;
                    let value = attr.unescape_value()?.into_owned();
                    match attr.key.as_ref() {
                        b"lockStructure" => {
                            workbook_protection.lock_structure = parse_xml_bool(&value)
                        }
                        b"lockWindows" => workbook_protection.lock_windows = parse_xml_bool(&value),
                        b"workbookPassword" => {
                            workbook_protection.password_hash =
                                parse_xml_u16_hex(&value).filter(|hash| *hash != 0);
                        }
                        _ => {}
                    }
                }
            }
            Event::Start(e) | Event::Empty(e) if e.local_name().as_ref() == b"workbookView" => {
                let mut saw_window_attr = false;
                let mut window = workbook_view.window.clone().unwrap_or_default();
                for attr in e.attributes() {
                    let attr = attr?;
                    match attr.key.as_ref() {
                        b"activeTab" => {
                            if workbook_view.active_tab.is_none() {
                                workbook_view.active_tab =
                                    attr.unescape_value()?.trim().parse::<usize>().ok();
                            }
                        }
                        b"xWindow" => {
                            window.x = attr.unescape_value()?.trim().parse::<i32>().ok();
                            saw_window_attr = true;
                        }
                        b"yWindow" => {
                            window.y = attr.unescape_value()?.trim().parse::<i32>().ok();
                            saw_window_attr = true;
                        }
                        b"windowWidth" => {
                            window.width = attr.unescape_value()?.trim().parse::<u32>().ok();
                            saw_window_attr = true;
                        }
                        b"windowHeight" => {
                            window.height = attr.unescape_value()?.trim().parse::<u32>().ok();
                            saw_window_attr = true;
                        }
                        b"windowState" => {
                            let state = attr.unescape_value()?;
                            let state = state.trim();
                            window.state = if state.eq_ignore_ascii_case("minimized") {
                                Some(WorkbookWindowState::Minimized)
                            } else if state.eq_ignore_ascii_case("maximized") {
                                Some(WorkbookWindowState::Maximized)
                            } else if state.eq_ignore_ascii_case("normal") {
                                Some(WorkbookWindowState::Normal)
                            } else {
                                None
                            };
                            saw_window_attr = true;
                        }
                        _ => {}
                    }
                }

                if saw_window_attr {
                    // If the workbook view window is entirely default-ish (all zeros + normal),
                    // treat it as missing metadata to avoid persisting meaningless 0x0 geometry.
                    let is_empty = window.x.unwrap_or(0) == 0
                        && window.y.unwrap_or(0) == 0
                        && window.width.unwrap_or(0) == 0
                        && window.height.unwrap_or(0) == 0
                        && matches!(window.state, None | Some(WorkbookWindowState::Normal));
                    if !is_empty {
                        workbook_view.window = Some(window);
                    }
                }
            }
            Event::Start(e) | Event::Empty(e) if e.local_name().as_ref() == b"sheet" => {
                let mut name = None;
                let mut sheet_id = None;
                let mut r_id = None;
                let mut state = None;
                for attr in e.attributes() {
                    let attr = attr?;
                    let key = attr.key.as_ref();
                    match key {
                        b"name" => name = Some(attr.unescape_value()?.into_owned()),
                        b"sheetId" => {
                            sheet_id = Some(attr.unescape_value()?.trim().parse().unwrap_or(0))
                        }
                        b"state" => state = Some(attr.unescape_value()?.into_owned()),
                        _ if crate::openxml::local_name(key) == b"id" => {
                            r_id = Some(attr.unescape_value()?.into_owned())
                        }
                        _ => {}
                    }
                }
                let name = name.unwrap_or_else(|| "Sheet".to_string());
                let sheet_id = sheet_id.unwrap_or(0);
                let relationship_id = r_id.unwrap_or_else(|| "rId1".to_string());
                let target = rels
                    .get(&relationship_id)
                    .cloned()
                    .unwrap_or_else(|| "worksheets/sheet1.xml".to_string());
                let path = resolve_target(WORKBOOK_PART, &target);
                sheets.push(ParsedSheet {
                    name,
                    sheet_id,
                    relationship_id,
                    state,
                    path,
                });
            }
            Event::Start(e) if e.local_name().as_ref() == b"definedName" => {
                let mut name = None;
                let mut local_sheet_id = None;
                let mut comment = None;
                let mut hidden = false;
                for attr in e.attributes() {
                    let attr = attr?;
                    match attr.key.as_ref() {
                        b"name" => name = Some(attr.unescape_value()?.into_owned()),
                        b"localSheetId" => {
                            local_sheet_id = attr.unescape_value()?.trim().parse::<u32>().ok();
                        }
                        b"comment" => comment = Some(attr.unescape_value()?.into_owned()),
                        b"hidden" => {
                            let v = attr.unescape_value()?;
                            hidden = parse_xml_bool(&v);
                        }
                        _ => {}
                    }
                }
                let Some(name) = name else {
                    current_defined = None;
                    continue;
                };
                current_defined = Some(ParsedDefinedName {
                    name,
                    local_sheet_id,
                    comment,
                    hidden,
                    value: String::new(),
                });
            }
            Event::Empty(e) if e.local_name().as_ref() == b"definedName" => {
                let mut name = None;
                let mut local_sheet_id = None;
                let mut comment = None;
                let mut hidden = false;
                for attr in e.attributes() {
                    let attr = attr?;
                    match attr.key.as_ref() {
                        b"name" => name = Some(attr.unescape_value()?.into_owned()),
                        b"localSheetId" => {
                            local_sheet_id = attr.unescape_value()?.trim().parse::<u32>().ok();
                        }
                        b"comment" => comment = Some(attr.unescape_value()?.into_owned()),
                        b"hidden" => {
                            let v = attr.unescape_value()?;
                            hidden = parse_xml_bool(&v);
                        }
                        _ => {}
                    }
                }
                let Some(name) = name else {
                    continue;
                };
                defined_names.push(ParsedDefinedName {
                    name,
                    local_sheet_id,
                    comment,
                    hidden,
                    value: String::new(),
                });
            }
            Event::Text(e) if current_defined.is_some() => {
                if let Some(ref mut dn) = current_defined {
                    dn.value.push_str(&e.unescape()?.to_string());
                }
            }
            Event::CData(e) if current_defined.is_some() => {
                if let Some(ref mut dn) = current_defined {
                    dn.value.push_str(std::str::from_utf8(e.as_ref())?);
                }
            }
            Event::End(e) if e.local_name().as_ref() == b"definedName" => {
                if let Some(dn) = current_defined.take() {
                    // Defined name `refersTo` values follow the same `_xlfn.` forward-compatibility
                    // convention as cell formulas. Strip `_xlfn.` prefixes so the model uses the
                    // UI-facing formula text (matching how we store `Cell::formula`).
                    let value = crate::formula_text::strip_xlfn_prefixes(&dn.value);
                    defined_names.push(ParsedDefinedName { value, ..dn });
                }
            }
            Event::Eof => break,
            _ => {}
        }
        buf.clear();
    }

    Ok((
        date_system,
        calc_pr,
        sheets,
        defined_names,
        workbook_protection,
        workbook_view,
    ))
}

fn parse_worksheet_into_model(
    worksheet: &mut formula_model::Worksheet,
    worksheet_id: formula_model::WorksheetId,
    worksheet_xml: &[u8],
    shared_strings: &[RichText],
    styles_part: &StylesPart,
    mut cell_meta_map: Option<
        &mut std::collections::HashMap<(formula_model::WorksheetId, CellRef), CellMeta>,
    >,
    mut rich_value_cells: Option<
        &mut std::collections::HashMap<(formula_model::WorksheetId, CellRef), u32>,
    >,
    metadata_part: Option<&MetadataPart>,
) -> Result<(), ReadError> {
    let mut reader = Reader::from_reader(worksheet_xml);
    reader.config_mut().trim_text(false);
    let mut buf = Vec::new();

    let mut in_sheet_data = false;
    let mut in_cols = false;
    let mut in_sheet_views = false;
    let mut in_sheet_view = false;
    let mut in_first_sheet_view = false;
    let mut sheet_view_idx: usize = 0;
    let mut parsed_view = SheetView::default();

    // Data validations can appear anywhere in the worksheet XML (typically after `<sheetData>`).
    // Parse them inline so both the full and streaming readers materialize validations.

    // When we don't retain the full `cell_meta` map (fast reader), we still want to materialize
    // shared-formula followers into the worksheet model so formulas match the full reader.
    let mut shared_formula_groups: Option<HashMap<u32, SharedFormulaGroup>> =
        cell_meta_map.is_none().then(HashMap::new);

    let mut current_ref: Option<CellRef> = None;
    let mut current_t: Option<String> = None;
    let mut current_style: u32 = 0;
    let mut current_cm: Option<String> = None;
    let mut current_vm: Option<String> = None;
    let mut current_formula: Option<FormulaMeta> = None;
    let mut current_value_text: Option<String> = None;
    let mut current_inline_value: Option<CellValue> = None;
    let mut current_inline_phonetic: Option<String> = None;
    let mut in_v = false;
    let mut in_f = false;
    let mut pending_vm_cells: Vec<(CellRef, u32)> = Vec::new();

    loop {
        match reader.read_event_into(&mut buf)? {
            Event::Start(e) | Event::Empty(e) if e.local_name().as_ref() == b"sheetFormatPr" => {
                for attr in e.attributes() {
                    let attr = attr?;
                    match attr.key.as_ref() {
                        b"defaultRowHeight" => {
                            worksheet.default_row_height =
                                attr.unescape_value()?.trim().parse::<f32>().ok();
                        }
                        b"defaultColWidth" => {
                            worksheet.default_col_width =
                                attr.unescape_value()?.trim().parse::<f32>().ok();
                        }
                        b"baseColWidth" => {
                            worksheet.base_col_width =
                                attr.unescape_value()?.trim().parse::<u16>().ok();
                        }
                        _ => {}
                    }
                }
            }
            Event::Start(e) if e.local_name().as_ref() == b"dataValidation" => {
                if let Some((ranges, validation)) = parse_data_validation_start(&mut reader, &e)? {
                    worksheet.add_data_validation(ranges, validation);
                }
            }
            Event::Empty(e) if e.local_name().as_ref() == b"dataValidation" => {
                if let Some((ranges, validation)) = parse_data_validation_empty(&e) {
                    worksheet.add_data_validation(ranges, validation);
                }
            }
            Event::Start(e) if e.local_name().as_ref() == b"cols" => in_cols = true,
            Event::End(e) if e.local_name().as_ref() == b"cols" => in_cols = false,
            Event::Empty(e) if e.local_name().as_ref() == b"cols" => {
                in_cols = false;
                drop(e);
            }
            Event::Start(e) | Event::Empty(e) if in_cols && e.local_name().as_ref() == b"col" => {
                let mut min: Option<u32> = None;
                let mut max: Option<u32> = None;
                let mut width: Option<f32> = None;
                let mut custom_width: Option<bool> = None;
                let mut hidden = false;
                let mut style: Option<u32> = None;
                // `customFormat` is not a standard `<col>` attribute in SpreadsheetML, but some
                // producers emit it alongside `style`. If present and explicitly `0`, treat it as
                // clearing any default column formatting.
                let mut custom_format: Option<bool> = None;

                for attr in e.attributes() {
                    let attr = attr?;
                    match attr.key.as_ref() {
                        b"min" => min = Some(attr.unescape_value()?.trim().parse().unwrap_or(0)),
                        b"max" => max = Some(attr.unescape_value()?.trim().parse().unwrap_or(0)),
                        b"width" => {
                            width = attr.unescape_value()?.trim().parse::<f32>().ok();
                        }
                        b"customWidth" => {
                            let v = attr.unescape_value()?;
                            custom_width = Some(parse_xml_bool(&v));
                        }
                        b"style" => {
                            style = attr.unescape_value()?.trim().parse::<u32>().ok();
                        }
                        b"customFormat" => {
                            let v = attr.unescape_value()?;
                            custom_format = Some(parse_xml_bool(&v));
                        }
                        b"hidden" => {
                            let v = attr.unescape_value()?;
                            hidden = parse_xml_bool(&v);
                        }
                        _ => {}
                    }
                }

                let Some(min) = min else {
                    continue;
                };
                let max = max.unwrap_or(min).min(formula_model::EXCEL_MAX_COLS);
                if min == 0 || max == 0 || min > formula_model::EXCEL_MAX_COLS {
                    continue;
                }

                for col_1_based in min..=max {
                    let col = col_1_based - 1;
                    if custom_width != Some(false) {
                        if let Some(width) = width {
                            worksheet.set_col_width(col, Some(width));
                        }
                    }
                    if hidden {
                        worksheet.set_col_hidden(col, true);
                    }
                    // Column default formatting (`col/@style`) is separate from per-cell formatting.
                    // It is required to interpret Excel serial dates where cells inherit the
                    // number format from the column.
                    if custom_format == Some(false) {
                        worksheet.set_col_style_id(col, None);
                    } else if let Some(xf_index) = style {
                        let style_id = styles_part.style_id_for_xf(xf_index);
                        worksheet.set_col_style_id(col, (style_id != 0).then_some(style_id));
                    }
                }
            }

            Event::Start(e) if e.local_name().as_ref() == b"sheetData" => in_sheet_data = true,
            Event::End(e) if e.local_name().as_ref() == b"sheetData" => in_sheet_data = false,
            Event::Empty(e) if e.local_name().as_ref() == b"sheetData" => {
                in_sheet_data = false;
                drop(e);
            }

            Event::Start(e) if e.local_name().as_ref() == b"sheetViews" => in_sheet_views = true,
            Event::End(e) if e.local_name().as_ref() == b"sheetViews" => {
                in_sheet_views = false;
                in_sheet_view = false;
                in_first_sheet_view = false;
                drop(e);
            }
            Event::Empty(e) if e.local_name().as_ref() == b"sheetViews" => {
                in_sheet_views = false;
                in_sheet_view = false;
                in_first_sheet_view = false;
                drop(e);
            }

            Event::Start(e) if in_sheet_views && e.local_name().as_ref() == b"sheetView" => {
                in_sheet_view = true;
                in_first_sheet_view = sheet_view_idx == 0;
                sheet_view_idx += 1;
                if in_first_sheet_view {
                    for attr in e.attributes() {
                        let attr = attr?;
                        let val = attr.unescape_value()?.into_owned();
                        match attr.key.as_ref() {
                            b"zoomScale" => {
                                if let Ok(scale) = val.parse::<f32>() {
                                    parsed_view.zoom = scale / 100.0;
                                }
                            }
                            b"showGridLines" => parsed_view.show_grid_lines = parse_xml_bool(&val),
                            b"showRowColHeaders" | b"showHeadings" => {
                                parsed_view.show_headings = parse_xml_bool(&val)
                            }
                            b"showZeros" => parsed_view.show_zeros = parse_xml_bool(&val),
                            _ => {}
                        }
                    }
                }
            }
            Event::Empty(e) if in_sheet_views && e.local_name().as_ref() == b"sheetView" => {
                let is_first = sheet_view_idx == 0;
                sheet_view_idx += 1;
                if is_first {
                    for attr in e.attributes() {
                        let attr = attr?;
                        let val = attr.unescape_value()?.into_owned();
                        match attr.key.as_ref() {
                            b"zoomScale" => {
                                if let Ok(scale) = val.parse::<f32>() {
                                    parsed_view.zoom = scale / 100.0;
                                }
                            }
                            b"showGridLines" => parsed_view.show_grid_lines = parse_xml_bool(&val),
                            b"showRowColHeaders" | b"showHeadings" => {
                                parsed_view.show_headings = parse_xml_bool(&val)
                            }
                            b"showZeros" => parsed_view.show_zeros = parse_xml_bool(&val),
                            _ => {}
                        }
                    }
                }
                // Self-closing `<sheetView/>` does not establish an active parse context.
                in_sheet_view = false;
                in_first_sheet_view = false;
                drop(e);
            }
            Event::End(e) if in_sheet_view && e.local_name().as_ref() == b"sheetView" => {
                in_sheet_view = false;
                in_first_sheet_view = false;
                drop(e);
            }

            Event::Start(e) | Event::Empty(e)
                if in_sheet_view && in_first_sheet_view && e.local_name().as_ref() == b"pane" =>
            {
                let mut state: Option<String> = None;
                let mut x_split_raw: Option<String> = None;
                let mut y_split_raw: Option<String> = None;
                let mut top_left_cell: Option<String> = None;
                for attr in e.attributes() {
                    let attr = attr?;
                    let val = attr.unescape_value()?.into_owned();
                    match attr.key.as_ref() {
                        b"state" => state = Some(val),
                        b"xSplit" => x_split_raw = Some(val),
                        b"ySplit" => y_split_raw = Some(val),
                        b"topLeftCell" => top_left_cell = Some(val),
                        _ => {}
                    }
                }

                if let Some(a1) = top_left_cell.as_deref() {
                    if let Ok(cell) = CellRef::from_a1(a1.trim()) {
                        parsed_view.pane.top_left_cell = Some(cell);
                    }
                }

                let state = state.as_deref().map(|s| s.trim());
                if matches!(state, Some("frozen") | Some("frozenSplit")) {
                    let frozen_cols = x_split_raw
                        .as_deref()
                        .and_then(|s| s.trim().parse::<u32>().ok())
                        .or_else(|| {
                            x_split_raw
                                .as_deref()
                                .and_then(|s| s.trim().parse::<f32>().ok())
                                .map(|f| f.round() as u32)
                        })
                        .unwrap_or(0);
                    let frozen_rows = y_split_raw
                        .as_deref()
                        .and_then(|s| s.trim().parse::<u32>().ok())
                        .or_else(|| {
                            y_split_raw
                                .as_deref()
                                .and_then(|s| s.trim().parse::<f32>().ok())
                                .map(|f| f.round() as u32)
                        })
                        .unwrap_or(0);
                    parsed_view.pane.frozen_cols = frozen_cols;
                    parsed_view.pane.frozen_rows = frozen_rows;
                    parsed_view.pane.x_split = None;
                    parsed_view.pane.y_split = None;
                } else {
                    parsed_view.pane.frozen_cols = 0;
                    parsed_view.pane.frozen_rows = 0;
                    parsed_view.pane.x_split =
                        x_split_raw.as_deref().and_then(|s| s.trim().parse::<f32>().ok());
                    parsed_view.pane.y_split =
                        y_split_raw.as_deref().and_then(|s| s.trim().parse::<f32>().ok());
                }
            }

            Event::Start(e) | Event::Empty(e)
                if in_sheet_view
                    && in_first_sheet_view
                    && e.local_name().as_ref() == b"selection" =>
            {
                // Excel can emit multiple `<selection>` blocks (one per pane). We only model the
                // first one.
                if parsed_view.selection.is_some() {
                    drop(e);
                } else {
                    let mut active_cell: Option<CellRef> = None;
                    let mut sqref: Option<String> = None;
                    for attr in e.attributes() {
                        let attr = attr?;
                        let val = attr.unescape_value()?.into_owned();
                        match attr.key.as_ref() {
                            b"activeCell" => {
                                if let Ok(cell) = CellRef::from_a1(val.trim()) {
                                    active_cell = Some(cell);
                                }
                            }
                            b"sqref" => sqref = Some(val),
                            _ => {}
                        }
                    }

                    if let Some(active_cell) = active_cell {
                        parsed_view.selection = match sqref.as_deref() {
                            Some(sqref) => SheetSelection::from_sqref(active_cell, sqref.trim()).ok(),
                            None => Some(SheetSelection::new(active_cell, Vec::new())),
                        };
                    }
                }
            }

            Event::Start(e) | Event::Empty(e) if e.local_name().as_ref() == b"sheetProtection" => {
                // Parse a subset of the legacy `sheetProtection` element into the model's
                // allow-list booleans. This is best-effort; unsupported attributes are ignored.
                //
                // Note: SpreadsheetML uses `objects`/`scenarios` as "protected" flags, while the
                // model stores `edit_objects` / `edit_scenarios` as "allowed" flags.
                let mut protection = SheetProtection::default();
                protection.enabled = true;
                for attr in e.attributes() {
                    let attr = attr?;
                    let val = attr.unescape_value()?.into_owned();
                    match attr.key.as_ref() {
                        b"sheet" => protection.enabled = parse_xml_bool(&val),
                        b"selectLockedCells" => {
                            protection.select_locked_cells = parse_xml_bool(&val)
                        }
                        b"selectUnlockedCells" => {
                            protection.select_unlocked_cells = parse_xml_bool(&val)
                        }
                        b"formatCells" => protection.format_cells = parse_xml_bool(&val),
                        b"formatColumns" => protection.format_columns = parse_xml_bool(&val),
                        b"formatRows" => protection.format_rows = parse_xml_bool(&val),
                        b"insertColumns" => protection.insert_columns = parse_xml_bool(&val),
                        b"insertRows" => protection.insert_rows = parse_xml_bool(&val),
                        b"insertHyperlinks" => protection.insert_hyperlinks = parse_xml_bool(&val),
                        b"deleteColumns" => protection.delete_columns = parse_xml_bool(&val),
                        b"deleteRows" => protection.delete_rows = parse_xml_bool(&val),
                        b"sort" => protection.sort = parse_xml_bool(&val),
                        b"autoFilter" => protection.auto_filter = parse_xml_bool(&val),
                        b"pivotTables" => protection.pivot_tables = parse_xml_bool(&val),
                        b"objects" => protection.edit_objects = !parse_xml_bool(&val),
                        b"scenarios" => protection.edit_scenarios = !parse_xml_bool(&val),
                        b"password" => {
                            protection.password_hash =
                                parse_xml_u16_hex(&val).filter(|hash| *hash != 0);
                        }
                        _ => {}
                    }
                }
                worksheet.sheet_protection = protection;
            }

            Event::Start(e) | Event::Empty(e)
                if in_sheet_data && e.local_name().as_ref() == b"row" =>
            {
                let mut row_1_based: Option<u32> = None;
                let mut height: Option<f32> = None;
                let mut custom_height: Option<bool> = None;
                let mut hidden = false;
                let mut style_xf: Option<u32> = None;
                let mut custom_format: Option<bool> = None;

                for attr in e.attributes() {
                    let attr = attr?;
                    match attr.key.as_ref() {
                        b"r" => {
                            row_1_based = Some(attr.unescape_value()?.trim().parse().unwrap_or(0));
                        }
                        b"ht" => {
                            height = attr.unescape_value()?.trim().parse::<f32>().ok();
                        }
                        b"customHeight" => {
                            let v = attr.unescape_value()?;
                            custom_height = Some(parse_xml_bool(&v));
                        }
                        b"hidden" => {
                            let v = attr.unescape_value()?;
                            hidden = parse_xml_bool(&v);
                        }
                        b"s" => {
                            style_xf = attr.unescape_value()?.trim().parse::<u32>().ok();
                        }
                        b"customFormat" => {
                            let v = attr.unescape_value()?;
                            custom_format = Some(parse_xml_bool(&v));
                        }
                        _ => {}
                    }
                }

                if let Some(row_1_based) = row_1_based {
                    // Accept any 1-based row index that fits in `u32` (the OOXML schema uses
                    // unsigned integers and our model supports sheets beyond Excel's UI limits).
                    if row_1_based > 0 {
                        let row = row_1_based - 1;
                        if custom_height != Some(false) {
                            if let Some(height) = height {
                                worksheet.set_row_height(row, Some(height));
                            }
                        }
                        if hidden {
                            worksheet.set_row_hidden(row, true);
                        }
                        // Row default formatting (`row/@s`) is nominally gated by `customFormat`,
                        // but some producers omit the flag while still intending the style to
                        // apply. We only treat `customFormat="0"` as an explicit "no style" signal.
                        if custom_format == Some(false) {
                            worksheet.set_row_style_id(row, None);
                        } else if let Some(xf_index) = style_xf {
                            let style_id = styles_part.style_id_for_xf(xf_index);
                            worksheet.set_row_style_id(row, (style_id != 0).then_some(style_id));
                        }
                    }
                }
            }

            Event::Start(e) if in_sheet_data && e.local_name().as_ref() == b"c" => {
                current_ref = None;
                current_t = None;
                current_style = 0;
                current_cm = None;
                current_vm = None;
                current_formula = None;
                current_value_text = None;
                current_inline_value = None;
                current_inline_phonetic = None;
                in_v = false;
                in_f = false;

                for attr in e.attributes() {
                    let attr = attr?;
                    match attr.key.as_ref() {
                        b"r" => {
                            let a1 = attr.unescape_value()?.into_owned();
                            current_ref = Some(
                                CellRef::from_a1(a1.trim())
                                    .map_err(|_| ReadError::InvalidCellRef(a1))?,
                            );
                        }
                        b"t" => current_t = Some(attr.unescape_value()?.into_owned()),
                        b"s" => {
                            if let Ok(xf_index) = attr.unescape_value()?.trim().parse::<u32>() {
                                current_style = styles_part.style_id_for_xf(xf_index);
                            }
                        }
                        b"cm" => current_cm = Some(attr.unescape_value()?.into_owned()),
                        b"vm" => current_vm = Some(attr.unescape_value()?.into_owned()),
                        _ => {}
                    }
                }
            }
            Event::Empty(e) if in_sheet_data && e.local_name().as_ref() == b"c" => {
                let mut cell_ref = None;
                let mut style_id = 0u32;
                let mut cm: Option<String> = None;
                let mut vm: Option<String> = None;
                for attr in e.attributes() {
                    let attr = attr?;
                    match attr.key.as_ref() {
                        b"r" => {
                            let a1 = attr.unescape_value()?.into_owned();
                            cell_ref = Some(
                                CellRef::from_a1(a1.trim())
                                    .map_err(|_| ReadError::InvalidCellRef(a1))?,
                            );
                        }
                        b"s" => {
                            if let Ok(xf_index) = attr.unescape_value()?.trim().parse::<u32>() {
                                style_id = styles_part.style_id_for_xf(xf_index);
                            }
                        }
                        b"cm" => cm = Some(attr.unescape_value()?.into_owned()),
                        b"vm" => vm = Some(attr.unescape_value()?.into_owned()),
                        _ => {}
                    }
                }
                if let Some(cell_ref) = cell_ref {
                    // Skip non-anchor cells inside merged regions. Excel stores the value
                    // (and typically formatting) on the top-left cell only.
                    if worksheet.merged_regions.resolve_cell(cell_ref) == cell_ref {
                        if let (Some(vm), Some(_metadata_part), Some(_rich_value_cells)) =
                            (vm.as_deref(), metadata_part, rich_value_cells.as_mut())
                        {
                            if let Ok(vm_idx) = vm.trim().parse::<u32>() {
                                pending_vm_cells.push((cell_ref, vm_idx));
                            }
                        }

                        if style_id != 0 {
                            let mut cell = Cell::default();
                            cell.style_id = style_id;
                            worksheet.set_cell(cell_ref, cell);
                        }
                    }

                    if let Some(cell_meta_map) = cell_meta_map.as_mut() {
                        if cm.is_some() || vm.is_some() {
                            let mut meta = CellMeta::default();
                            meta.cm = cm;
                            meta.vm = vm;
                            cell_meta_map.insert((worksheet_id, cell_ref), meta);
                        }
                    }
                }
            }

            Event::End(e) if in_sheet_data && e.local_name().as_ref() == b"c" => {
                if let Some(cell_ref) = current_ref {
                    if worksheet.merged_regions.resolve_cell(cell_ref) == cell_ref {
                        let (value, value_kind, raw_value) = if cell_meta_map.is_some() {
                            interpret_cell_value(
                                current_t.as_deref(),
                                &current_value_text,
                                &current_inline_value,
                                shared_strings,
                            )
                        } else {
                            (
                                interpret_cell_value_without_meta(
                                    current_t.as_deref(),
                                    &current_value_text,
                                    &current_inline_value,
                                    shared_strings,
                                ),
                                None,
                                None,
                            )
                        };

                        let formula_in_model = current_formula.as_ref().and_then(|f| {
                            let stripped = crate::formula_text::strip_xlfn_prefixes(&f.file_text);
                            normalize_formula_text(&stripped)
                        });

                        if let Some(groups) = shared_formula_groups.as_mut() {
                            let is_shared_master =
                                current_formula.as_ref().is_some_and(|formula| {
                                    formula.t.as_deref() == Some("shared")
                                        && formula.reference.is_some()
                                        && formula.shared_index.is_some()
                                        && !formula.file_text.is_empty()
                                });

                            if is_shared_master {
                                if let Some(formula) = current_formula.as_ref() {
                                    if let (Some(reference), Some(shared_index)) =
                                        (formula.reference.as_deref(), formula.shared_index)
                                    {
                                        if let Ok(range) = Range::from_a1(reference) {
                                            let master_display =
                                                crate::formula_text::strip_xlfn_prefixes(
                                                    &formula.file_text,
                                                );
                                            let mut opts = ParseOptions::default();
                                            opts.normalize_relative_to =
                                                Some(CellAddr::new(cell_ref.row, cell_ref.col));

                                            if let Ok(ast) = parse_formula(&master_display, opts) {
                                                groups.insert(
                                                    shared_index,
                                                    SharedFormulaGroup { range, ast },
                                                );
                                            }
                                        }
                                    }
                                }
                            }
                        }

                        let mut cell = Cell::default();
                        cell.value = value;
                        cell.formula = formula_in_model;
                        cell.phonetic = match current_t.as_deref() {
                            Some("s") => {
                                let idx = current_value_text
                                    .as_deref()
                                    .unwrap_or("0")
                                    .trim()
                                    .parse::<usize>()
                                    .unwrap_or(0);
                                shared_strings.get(idx).and_then(|s| s.phonetic.clone())
                            }
                            Some("inlineStr") => current_inline_phonetic.take(),
                            _ => None,
                        };
                        cell.style_id = current_style;

                        if let (Some(vm), Some(_metadata_part), Some(_rich_value_cells)) = (
                            current_vm.as_deref(),
                            metadata_part,
                            rich_value_cells.as_mut(),
                        ) {
                            if let Ok(vm_idx) = vm.trim().parse::<u32>() {
                                pending_vm_cells.push((cell_ref, vm_idx));
                            }
                        }

                        if !cell.is_truly_empty() {
                            worksheet.set_cell(cell_ref, cell);
                        }

                        if let Some(cell_meta_map) = cell_meta_map.as_mut() {
                            let mut meta = CellMeta::default();
                            meta.value_kind = value_kind;
                            meta.raw_value = raw_value;
                            meta.formula = current_formula.take();
                            meta.cm = current_cm.take();
                            meta.vm = current_vm.take();

                            if meta.value_kind.is_some()
                                || meta.raw_value.is_some()
                                || meta.formula.is_some()
                                || current_style != 0
                                || meta.cm.is_some()
                                || meta.vm.is_some()
                            {
                                cell_meta_map.insert((worksheet_id, cell_ref), meta);
                            }
                        }
                    }
                }

                current_ref = None;
                current_t = None;
                current_style = 0;
                current_cm = None;
                current_vm = None;
                current_formula = None;
                current_value_text = None;
                current_inline_value = None;
                current_inline_phonetic = None;
                in_v = false;
                in_f = false;
            }

            Event::Start(e)
                if in_sheet_data && current_ref.is_some() && e.local_name().as_ref() == b"v" =>
            {
                in_v = true;
            }
            Event::End(e) if in_sheet_data && e.local_name().as_ref() == b"v" => in_v = false,
            Event::Text(e) if in_sheet_data && in_v => {
                current_value_text = Some(e.unescape()?.into_owned());
            }

            Event::Start(e)
                if in_sheet_data && current_ref.is_some() && e.local_name().as_ref() == b"f" =>
            {
                in_f = true;
                let mut formula = FormulaMeta::default();
                for attr in e.attributes() {
                    let attr = attr?;
                    match attr.key.as_ref() {
                        b"t" => formula.t = Some(attr.unescape_value()?.into_owned()),
                        b"ref" => formula.reference = Some(attr.unescape_value()?.into_owned()),
                        b"si" => {
                            formula.shared_index =
                                Some(attr.unescape_value()?.trim().parse().unwrap_or(0))
                        }
                        b"aca" => {
                            let v = attr.unescape_value()?;
                            formula.always_calc = Some(parse_xml_bool(&v))
                        }
                        _ => {}
                    }
                }
                current_formula = Some(formula);
            }
            Event::Empty(e)
                if in_sheet_data && current_ref.is_some() && e.local_name().as_ref() == b"f" =>
            {
                let mut formula = FormulaMeta::default();
                for attr in e.attributes() {
                    let attr = attr?;
                    match attr.key.as_ref() {
                        b"t" => formula.t = Some(attr.unescape_value()?.into_owned()),
                        b"ref" => formula.reference = Some(attr.unescape_value()?.into_owned()),
                        b"si" => {
                            formula.shared_index =
                                Some(attr.unescape_value()?.trim().parse().unwrap_or(0))
                        }
                        b"aca" => {
                            let v = attr.unescape_value()?;
                            formula.always_calc = Some(parse_xml_bool(&v))
                        }
                        _ => {}
                    }
                }
                current_formula = Some(formula);
            }
            Event::End(e) if in_sheet_data && e.local_name().as_ref() == b"f" => in_f = false,
            Event::Text(e) if in_sheet_data && in_f => {
                if let Some(formula) = current_formula.as_mut() {
                    formula.file_text = e.unescape()?.into_owned();
                }
            }

            Event::Start(e)
                if in_sheet_data
                    && current_ref.is_some()
                    && current_t.as_deref() == Some("inlineStr")
                    && e.local_name().as_ref() == b"is" =>
            {
                let (value, phonetic) = parse_inline_is_value(&mut reader)?;
                current_inline_value = Some(value);
                current_inline_phonetic = phonetic;
            }
            Event::Empty(e)
                if in_sheet_data
                    && current_ref.is_some()
                    && current_t.as_deref() == Some("inlineStr")
                    && e.local_name().as_ref() == b"is" =>
            {
                current_inline_value = Some(CellValue::String(String::new()));
                current_inline_phonetic = None;
            }

            Event::Eof => break,
            _ => {}
        }
        buf.clear();
    }

    // If the worksheet XML included `<sheetViews>`, use the parsed view model. Otherwise keep the
    // default view state.
    worksheet.view = parsed_view;

    // Keep legacy fields in sync for backward compatibility.
    worksheet.zoom = worksheet.view.zoom;
    worksheet.frozen_rows = worksheet.view.pane.frozen_rows;
    worksheet.frozen_cols = worksheet.view.pane.frozen_cols;

    // Resolve `c/@vm` values to rich value record indices after parsing the sheet.
    //
    // Note: `c/@vm` is ambiguous across producers (0-based vs 1-based). `MetadataPart` performs
    // best-effort resolution using workbook-level heuristics, so we always pass through the raw
    // `vm` value here (do not apply an additional offset).
    if !pending_vm_cells.is_empty() {
        if let (Some(metadata_part), Some(rich_value_cells)) = (metadata_part, rich_value_cells) {
            for (cell_ref, vm) in pending_vm_cells {
                if let Some(idx) = metadata_part.vm_to_rich_value_index(vm) {
                    rich_value_cells.insert((worksheet_id, cell_ref), idx);
                }
            }
        }
    }

    if let Some(groups) = shared_formula_groups {
        for group in groups.values() {
            for cell_ref in group.range.iter() {
                // Avoid overwriting any explicit formula already present in the worksheet. The goal is
                // to materialize formulas for shared-formula followers that are textless in the file.
                if worksheet.formula(cell_ref).is_some() {
                    continue;
                }

                let mut ser = SerializeOptions::default();
                ser.origin = Some(CellAddr::new(cell_ref.row, cell_ref.col));
                ser.omit_equals = true;

                let display = match group.ast.to_string(ser) {
                    Ok(s) => s,
                    Err(_) => continue,
                };

                worksheet.set_formula(cell_ref, Some(display));
            }
        }
    }

    // Best-effort: parse outline (row/column grouping) metadata.
    //
    // Notes:
    // - `RowProperties.hidden` / `ColProperties.hidden` represent *user-hidden* state (Hide Row/Col).
    // - Outline-hidden state (collapsed groups) lives in `worksheet.outline.*.hidden.outline`.
    //
    // The main sheet reader above parses `row/@hidden` and `col/@hidden` as `set_*_hidden`, which
    // (intentionally) maps to the model's user-hidden bit. However, in OOXML the same `hidden="1"`
    // attribute is also used for outline-collapsed detail rows/cols. We therefore re-parse outline
    // state using the dedicated outline parser and then sync the per-row/col "user hidden" bits to
    // match the outline parser's `hidden.user` inference.
    if let Ok(xml) = std::str::from_utf8(worksheet_xml) {
        if let Ok(outline) = crate::outline::read_outline_from_worksheet_xml(xml) {
            worksheet.outline = outline;

            // Sync existing row/col properties entries so their `hidden` flags reflect only the
            // user-hidden bit (not outline-hidden).
            //
            // Note: outline indices are 1-based, while row/col properties are 0-based.
            let row_keys: Vec<u32> = worksheet.row_properties.keys().copied().collect();
            for row_0based in row_keys {
                let user_hidden = worksheet
                    .outline
                    .rows
                    .entry(row_0based.saturating_add(1))
                    .hidden
                    .user;
                worksheet.set_row_hidden(row_0based, user_hidden);
            }

            let col_keys: Vec<u32> = worksheet.col_properties.keys().copied().collect();
            for col_0based in col_keys {
                let user_hidden = worksheet
                    .outline
                    .cols
                    .entry(col_0based.saturating_add(1))
                    .hidden
                    .user;
                worksheet.set_col_hidden(col_0based, user_hidden);
            }
        }
    }

    Ok(())
}

fn parse_data_validation_kind(val: &str) -> Option<DataValidationKind> {
    match val {
        v if v.eq_ignore_ascii_case("whole") => Some(DataValidationKind::Whole),
        v if v.eq_ignore_ascii_case("decimal") => Some(DataValidationKind::Decimal),
        v if v.eq_ignore_ascii_case("list") => Some(DataValidationKind::List),
        v if v.eq_ignore_ascii_case("date") => Some(DataValidationKind::Date),
        v if v.eq_ignore_ascii_case("time") => Some(DataValidationKind::Time),
        v if v.eq_ignore_ascii_case("textLength") => Some(DataValidationKind::TextLength),
        v if v.eq_ignore_ascii_case("custom") => Some(DataValidationKind::Custom),
        _ => None,
    }
}

fn parse_data_validation_operator(val: &str) -> Option<DataValidationOperator> {
    match val {
        v if v.eq_ignore_ascii_case("between") => Some(DataValidationOperator::Between),
        v if v.eq_ignore_ascii_case("notBetween") => Some(DataValidationOperator::NotBetween),
        v if v.eq_ignore_ascii_case("equal") => Some(DataValidationOperator::Equal),
        v if v.eq_ignore_ascii_case("notEqual") => Some(DataValidationOperator::NotEqual),
        v if v.eq_ignore_ascii_case("greaterThan") => Some(DataValidationOperator::GreaterThan),
        v if v.eq_ignore_ascii_case("greaterThanOrEqual") => {
            Some(DataValidationOperator::GreaterThanOrEqual)
        }
        v if v.eq_ignore_ascii_case("lessThan") => Some(DataValidationOperator::LessThan),
        v if v.eq_ignore_ascii_case("lessThanOrEqual") => {
            Some(DataValidationOperator::LessThanOrEqual)
        }
        _ => None,
    }
}

fn parse_data_validation_error_style(val: &str) -> Option<DataValidationErrorStyle> {
    match val {
        v if v.eq_ignore_ascii_case("stop") => Some(DataValidationErrorStyle::Stop),
        v if v.eq_ignore_ascii_case("warning") => Some(DataValidationErrorStyle::Warning),
        v if v.eq_ignore_ascii_case("information") => Some(DataValidationErrorStyle::Information),
        _ => None,
    }
}

fn parse_sqref_ranges(val: &str) -> Vec<Range> {
    val.split_whitespace()
        .filter_map(|part| Range::from_a1(part).ok())
        .collect()
}

fn normalize_data_validation_formula(raw: &str) -> String {
    // Data validation formulas are stored without a leading `=` in the model.
    //
    // Preserve the original text but:
    // - strip a single leading `=` if present (model invariant)
    // - strip `_xlfn.` prefixes at function-call boundaries (matches cell formula normalization)
    crate::formula_text::strip_xlfn_prefixes(raw.strip_prefix('=').unwrap_or(raw))
}

fn parse_data_validation_empty(
    e: &quick_xml::events::BytesStart<'_>,
) -> Option<(Vec<Range>, DataValidation)> {
    let mut kind: Option<DataValidationKind> = None;
    let mut operator: Option<DataValidationOperator> = None;
    let mut allow_blank = false;
    let mut show_input_message = false;
    let mut show_error_message = false;
    let mut show_drop_down: Option<bool> = None;

    let mut prompt_title: Option<String> = None;
    let mut prompt_body: Option<String> = None;
    let mut error_style: Option<DataValidationErrorStyle> = None;
    let mut error_title: Option<String> = None;
    let mut error_body: Option<String> = None;

    let mut ranges: Vec<Range> = Vec::new();

    for attr in e.attributes().with_checks(false).flatten() {
        let key = crate::openxml::local_name(attr.key.as_ref());
        let Ok(val) = attr.unescape_value() else {
            continue;
        };
        let val = val.into_owned();
        match key {
            b"type" => kind = parse_data_validation_kind(&val),
            b"operator" => operator = parse_data_validation_operator(&val),
            b"allowBlank" => allow_blank = parse_xml_bool(&val),
            b"showInputMessage" => show_input_message = parse_xml_bool(&val),
            b"showErrorMessage" => show_error_message = parse_xml_bool(&val),
            b"showDropDown" => {
                // NOTE: Excel's `showDropDown` is historically inverted relative to the UI:
                // `showDropDown="1"` means "suppress the in-cell dropdown arrow".
                //
                // The in-memory model stores the UI-facing behavior: `show_drop_down=true` means
                // "show the in-cell dropdown arrow". Therefore we invert the OOXML attribute.
                //
                // See: ECMA-376 SpreadsheetML `CT_DataValidation` / `showDropDown` and
                // https://learn.microsoft.com/en-us/openspecs/office_standards/ms-xlsx/
                show_drop_down = Some(!parse_xml_bool(&val));
            }
            b"promptTitle" => prompt_title = Some(val),
            b"prompt" => prompt_body = Some(val),
            b"errorStyle" => error_style = parse_data_validation_error_style(&val),
            b"errorTitle" => error_title = Some(val),
            b"error" => error_body = Some(val),
            b"sqref" => ranges = parse_sqref_ranges(&val),
            _ => {}
        }
    }

    let kind = kind?;
    if ranges.is_empty() {
        return None;
    }

    // Only list validations have an in-cell dropdown affordance. When the OOXML attribute is
    // omitted, Excel shows the arrow by default.
    let show_drop_down = if kind == DataValidationKind::List {
        show_drop_down.unwrap_or(true)
    } else {
        false
    };

    let input_message = if prompt_title.is_some() || prompt_body.is_some() {
        Some(DataValidationInputMessage {
            title: prompt_title,
            body: prompt_body,
        })
    } else {
        None
    };

    let error_alert = if error_style.is_some() || error_title.is_some() || error_body.is_some() {
        Some(DataValidationErrorAlert {
            style: error_style.unwrap_or_default(),
            title: error_title,
            body: error_body,
        })
    } else {
        None
    };

    Some((
        ranges,
        DataValidation {
            kind,
            operator,
            formula1: String::new(),
            formula2: None,
            allow_blank,
            show_input_message,
            show_error_message,
            show_drop_down,
            input_message,
            error_alert,
        },
    ))
}

fn parse_data_validation_start<R: std::io::BufRead>(
    reader: &mut Reader<R>,
    e: &quick_xml::events::BytesStart<'_>,
) -> Result<Option<(Vec<Range>, DataValidation)>, ReadError> {
    let mut kind: Option<DataValidationKind> = None;
    let mut operator: Option<DataValidationOperator> = None;
    let mut allow_blank = false;
    let mut show_input_message = false;
    let mut show_error_message = false;
    let mut show_drop_down: Option<bool> = None;

    let mut prompt_title: Option<String> = None;
    let mut prompt_body: Option<String> = None;
    let mut error_style: Option<DataValidationErrorStyle> = None;
    let mut error_title: Option<String> = None;
    let mut error_body: Option<String> = None;

    let mut ranges: Vec<Range> = Vec::new();

    for attr in e.attributes().with_checks(false).flatten() {
        let key = crate::openxml::local_name(attr.key.as_ref());
        let Ok(val) = attr.unescape_value() else {
            continue;
        };
        let val = val.into_owned();
        match key {
            b"type" => kind = parse_data_validation_kind(&val),
            b"operator" => operator = parse_data_validation_operator(&val),
            b"allowBlank" => allow_blank = parse_xml_bool(&val),
            b"showInputMessage" => show_input_message = parse_xml_bool(&val),
            b"showErrorMessage" => show_error_message = parse_xml_bool(&val),
            b"showDropDown" => {
                // NOTE: Excel's `showDropDown` is historically inverted relative to the UI:
                // `showDropDown="1"` means "suppress the in-cell dropdown arrow".
                //
                // The in-memory model stores the UI-facing behavior: `show_drop_down=true` means
                // "show the in-cell dropdown arrow". Therefore we invert the OOXML attribute.
                //
                // See: ECMA-376 SpreadsheetML `CT_DataValidation` / `showDropDown` and
                // https://learn.microsoft.com/en-us/openspecs/office_standards/ms-xlsx/
                show_drop_down = Some(!parse_xml_bool(&val));
            }
            b"promptTitle" => prompt_title = Some(val),
            b"prompt" => prompt_body = Some(val),
            b"errorStyle" => error_style = parse_data_validation_error_style(&val),
            b"errorTitle" => error_title = Some(val),
            b"error" => error_body = Some(val),
            b"sqref" => ranges = parse_sqref_ranges(&val),
            _ => {}
        }
    }

    let should_store = kind.is_some() && !ranges.is_empty();

    let mut formula1 = String::new();
    let mut formula2: Option<String> = None;

    let mut buf = Vec::new();
    loop {
        match reader.read_event_into(&mut buf)? {
            Event::Start(e) if e.local_name().as_ref() == b"formula1" => {
                formula1 = normalize_data_validation_formula(&read_text(reader, b"formula1")?);
            }
            Event::Empty(e) if e.local_name().as_ref() == b"formula1" => {
                formula1.clear();
            }
            Event::Start(e) if e.local_name().as_ref() == b"formula2" => {
                let normalized =
                    normalize_data_validation_formula(&read_text(reader, b"formula2")?);
                if normalized.is_empty() {
                    formula2 = None;
                } else {
                    formula2 = Some(normalized);
                }
            }
            Event::Empty(e) if e.local_name().as_ref() == b"formula2" => {
                formula2 = None;
            }
            Event::Start(e) => {
                // Best-effort: skip any unmodeled nested content.
                reader.read_to_end_into(e.name(), &mut Vec::new())?;
            }
            Event::End(e) if e.local_name().as_ref() == b"dataValidation" => break,
            Event::Eof => {
                return Err(ReadError::Xlsx(XlsxError::Invalid(
                    "unexpected EOF while parsing <dataValidation>".to_string(),
                )))
            }
            _ => {}
        }
        buf.clear();
    }

    if !should_store {
        return Ok(None);
    }

    let Some(kind) = kind else {
        debug_assert!(
            false,
            "should_store was true but data validation kind was None (ranges={})",
            ranges.len()
        );
        return Ok(None);
    };

    // Only list validations have an in-cell dropdown affordance. When the OOXML attribute is
    // omitted, Excel shows the arrow by default.
    let show_drop_down = if kind == DataValidationKind::List {
        show_drop_down.unwrap_or(true)
    } else {
        false
    };

    let input_message = if prompt_title.is_some() || prompt_body.is_some() {
        Some(DataValidationInputMessage {
            title: prompt_title,
            body: prompt_body,
        })
    } else {
        None
    };

    let error_alert = if error_style.is_some() || error_title.is_some() || error_body.is_some() {
        Some(DataValidationErrorAlert {
            style: error_style.unwrap_or_default(),
            title: error_title,
            body: error_body,
        })
    } else {
        None
    };

    Ok(Some((
        ranges,
        DataValidation {
            kind,
            operator,
            formula1,
            formula2,
            allow_blank,
            show_input_message,
            show_error_message,
            show_drop_down,
            input_message,
            error_alert,
        },
    )))
}

#[derive(Debug, Clone)]
struct SharedFormulaGroup {
    range: Range,
    ast: formula_engine::Ast,
}

fn expand_shared_formulas(
    worksheet: &mut formula_model::Worksheet,
    worksheet_id: formula_model::WorksheetId,
    cell_meta_map: &HashMap<(formula_model::WorksheetId, CellRef), CellMeta>,
) {
    let mut groups: HashMap<u32, SharedFormulaGroup> = HashMap::new();

    for ((ws_id, cell_ref), meta) in cell_meta_map {
        if *ws_id != worksheet_id {
            continue;
        }
        let Some(formula) = meta.formula.as_ref() else {
            continue;
        };

        let is_shared_master = formula.t.as_deref() == Some("shared")
            && formula.reference.is_some()
            && formula.shared_index.is_some()
            && !formula.file_text.is_empty();
        if !is_shared_master {
            continue;
        }

        let Some(reference) = formula.reference.as_deref() else {
            continue;
        };
        let Some(shared_index) = formula.shared_index else {
            continue;
        };

        let range = match Range::from_a1(reference) {
            Ok(range) => range,
            Err(_) => continue,
        };

        let master_display = crate::formula_text::strip_xlfn_prefixes(&formula.file_text);
        let mut opts = ParseOptions::default();
        opts.normalize_relative_to = Some(CellAddr::new(cell_ref.row, cell_ref.col));

        let ast = match parse_formula(&master_display, opts) {
            Ok(ast) => ast,
            Err(_) => continue,
        };

        groups.insert(shared_index, SharedFormulaGroup { range, ast });
    }

    for group in groups.values() {
        for cell_ref in group.range.iter() {
            // Avoid overwriting any explicit formula already present in the worksheet. The goal is
            // to materialize formulas for shared-formula followers that are textless in the file.
            if worksheet.formula(cell_ref).is_some() {
                continue;
            }

            let mut ser = SerializeOptions::default();
            ser.origin = Some(CellAddr::new(cell_ref.row, cell_ref.col));
            ser.omit_equals = true;

            let display = match group.ast.to_string(ser) {
                Ok(s) => s,
                Err(_) => continue,
            };

            worksheet.set_formula(cell_ref, Some(display));
        }
    }
}

fn parse_xml_bool(val: &str) -> bool {
    let trimmed = val.trim();
    trimmed == "1" || trimmed.eq_ignore_ascii_case("true")
}

fn parse_xml_u16_hex(val: &str) -> Option<u16> {
    let trimmed = val.trim();
    if trimmed.is_empty() {
        return None;
    }
    u16::from_str_radix(trimmed, 16).ok()
}

fn page_setup_to_model(setup: crate::print::PageSetup) -> formula_model::PageSetup {
    let orientation = match setup.orientation {
        crate::print::Orientation::Portrait => formula_model::Orientation::Portrait,
        crate::print::Orientation::Landscape => formula_model::Orientation::Landscape,
    };

    let scaling = match setup.scaling {
        crate::print::Scaling::Percent(pct) => formula_model::Scaling::Percent(pct),
        crate::print::Scaling::FitTo { width, height } => {
            formula_model::Scaling::FitTo { width, height }
        }
    };

    formula_model::PageSetup {
        orientation,
        paper_size: formula_model::PaperSize {
            code: setup.paper_size.code,
        },
        margins: formula_model::PageMargins {
            left: setup.margins.left,
            right: setup.margins.right,
            top: setup.margins.top,
            bottom: setup.margins.bottom,
            header: setup.margins.header,
            footer: setup.margins.footer,
        },
        scaling,
    }
}

fn manual_page_breaks_to_model(
    breaks: crate::print::ManualPageBreaks,
) -> formula_model::ManualPageBreaks {
    let mut out = formula_model::ManualPageBreaks::default();
    for id in breaks.row_breaks_after {
        if id > 0 {
            out.row_breaks_after.insert(id - 1);
        }
    }
    for id in breaks.col_breaks_after {
        if id > 0 {
            out.col_breaks_after.insert(id - 1);
        }
    }
    out
}

fn parse_table_part_ids(xml: &[u8]) -> Result<Vec<String>, ReadError> {
    let mut reader = Reader::from_reader(xml);
    reader.config_mut().trim_text(true);
    let mut buf = Vec::new();
    let mut out = Vec::new();

    loop {
        match reader.read_event_into(&mut buf)? {
            Event::Start(e) | Event::Empty(e)
                if crate::openxml::local_name(e.name().as_ref()) == b"tablePart" =>
            {
                for attr in e.attributes() {
                    let attr = attr?;
                    if crate::openxml::local_name(attr.key.as_ref()) == b"id" {
                        out.push(attr.unescape_value()?.into_owned());
                    }
                }
            }
            Event::Eof => break,
            _ => {}
        }
        buf.clear();
    }

    Ok(out)
}

fn parse_sheet_drawing_part_ids(xml: &[u8]) -> Result<Vec<String>, ReadError> {
    let mut reader = Reader::from_reader(xml);
    reader.config_mut().trim_text(true);
    let mut buf = Vec::new();
    let mut out = Vec::new();

    loop {
        match reader.read_event_into(&mut buf)? {
            Event::Start(e) | Event::Empty(e)
                if crate::openxml::local_name(e.name().as_ref()) == b"drawing" =>
            {
                for attr in e.attributes() {
                    let attr = attr?;
                    if crate::openxml::local_name(attr.key.as_ref()) == b"id" {
                        let rel_id = attr.unescape_value()?.into_owned();
                        let rel_id = rel_id.trim();
                        if !rel_id.is_empty() {
                            out.push(rel_id.to_string());
                        }
                    }
                }
            }
            Event::Eof => break,
            _ => {}
        }
        buf.clear();
    }

    Ok(out)
}

fn parse_inline_is_value<R: std::io::BufRead>(
    reader: &mut Reader<R>,
) -> Result<(CellValue, Option<String>), ReadError> {
    let mut buf = Vec::new();
    let mut segments: Vec<(String, RichTextRunStyle)> = Vec::new();
    let mut phonetic: Option<String> = None;

    loop {
        match reader.read_event_into(&mut buf)? {
            Event::Start(e) if e.local_name().as_ref() == b"t" => {
                let t = read_text(reader, b"t")?;
                segments.push((t, RichTextRunStyle::default()));
            }
            Event::Start(e) if e.local_name().as_ref() == b"r" => {
                segments.push(parse_inline_r(reader)?);
            }
            Event::Start(e) if e.local_name().as_ref() == b"rPh" => {
                let run = parse_inline_rph_text(reader)?;
                phonetic.get_or_insert_with(String::new).push_str(&run);
            }
            Event::Empty(e) if e.local_name().as_ref() == b"rPh" => {
                // Presence of `<rPh/>` implies phonetic metadata even if it contains no text.
                phonetic.get_or_insert_with(String::new);
            }
            Event::Start(e) if e.local_name().as_ref() == b"phoneticPr" => {
                // `phoneticPr` carries metadata for the phonetic runs. Preserve its presence so we
                // can distinguish between "no phonetic" and "phonetic with empty text".
                phonetic.get_or_insert_with(String::new);
                reader.read_to_end_into(e.name(), &mut Vec::new())?;
            }
            Event::Empty(e) if e.local_name().as_ref() == b"phoneticPr" => {
                phonetic.get_or_insert_with(String::new);
            }
            Event::Start(e) => {
                // Only treat `<t>` as visible text when it is a direct child of `<is>` or inside
                // `<is><r>...</r></is>`. Other subtrees (phonetic/ruby runs, extensions) may contain
                // `<t>` elements that should not be concatenated into the display string, so skip
                // them wholesale.
                reader.read_to_end_into(e.name(), &mut Vec::new())?;
            }
            Event::End(e) if e.local_name().as_ref() == b"is" => break,
            Event::Eof => {
                return Err(ReadError::Xlsx(XlsxError::Invalid(
                    "unexpected EOF while parsing inline string <is>".to_string(),
                )))
            }
            _ => {}
        }
        buf.clear();
    }

    let value = if segments.is_empty() {
        CellValue::String(String::new())
    } else if segments.iter().all(|(_, style)| style.is_empty()) {
        CellValue::String(segments.into_iter().map(|(text, _)| text).collect())
    } else {
        CellValue::RichText(RichText::from_segments(segments))
    };

    Ok((value, phonetic))
}

fn parse_inline_r<R: std::io::BufRead>(
    reader: &mut Reader<R>,
) -> Result<(String, RichTextRunStyle), ReadError> {
    let mut buf = Vec::new();
    let mut style = RichTextRunStyle::default();
    let mut text = String::new();

    loop {
        match reader.read_event_into(&mut buf)? {
            Event::Start(e) if e.local_name().as_ref() == b"rPr" => {
                style = parse_inline_rpr(reader)?;
            }
            Event::Start(e) if e.local_name().as_ref() == b"t" => {
                text.push_str(&read_text(reader, b"t")?);
            }
            Event::Start(e) => {
                // Skip any unexpected subtrees to avoid accidentally capturing `<t>` elements
                // from phonetic/ruby annotations (best-effort).
                reader.read_to_end_into(e.name(), &mut Vec::new())?;
            }
            Event::End(e) if e.local_name().as_ref() == b"r" => break,
            Event::Eof => {
                return Err(ReadError::Xlsx(XlsxError::Invalid(
                    "unexpected EOF while parsing inline string <r>".to_string(),
                )))
            }
            _ => {}
        }
        buf.clear();
    }

    Ok((text, style))
}

fn parse_inline_rpr<R: std::io::BufRead>(
    reader: &mut Reader<R>,
) -> Result<RichTextRunStyle, ReadError> {
    let mut buf = Vec::new();
    let mut style = RichTextRunStyle::default();

    loop {
        match reader.read_event_into(&mut buf)? {
            Event::Empty(e) => parse_inline_rpr_tag(&e, &mut style)?,
            Event::Start(e) => {
                parse_inline_rpr_tag(&e, &mut style)?;
                reader.read_to_end_into(e.name(), &mut Vec::new())?;
            }
            Event::End(e) if e.local_name().as_ref() == b"rPr" => break,
            Event::Eof => {
                return Err(ReadError::Xlsx(XlsxError::Invalid(
                    "unexpected EOF while parsing inline string <rPr>".to_string(),
                )))
            }
            _ => {}
        }
        buf.clear();
    }

    Ok(style)
}

fn parse_inline_rpr_tag(e: &BytesStart<'_>, style: &mut RichTextRunStyle) -> Result<(), ReadError> {
    match e.local_name().as_ref() {
        b"b" => style.bold = Some(parse_inline_rpr_bool_val(e)?),
        b"i" => style.italic = Some(parse_inline_rpr_bool_val(e)?),
        b"u" => {
            let val = attr_value(e, b"val")?;
            if let Some(ul) = Underline::from_ooxml(val.as_deref()) {
                style.underline = Some(ul);
            }
        }
        b"color" => {
            if let Some(rgb) = attr_value(e, b"rgb")? {
                if rgb.len() == 8 {
                    if let Ok(argb) = u32::from_str_radix(&rgb, 16) {
                        style.color = Some(Color::new_argb(argb));
                    }
                }
            }
        }
        b"rFont" | b"name" => {
            if let Some(val) = attr_value(e, b"val")? {
                style.font = Some(val);
            }
        }
        b"sz" => {
            if let Some(val) = attr_value(e, b"val")? {
                if let Some(sz) = parse_size_100pt(&val) {
                    style.size_100pt = Some(sz);
                }
            }
        }
        _ => {}
    }
    Ok(())
}

fn parse_inline_rpr_bool_val(e: &BytesStart<'_>) -> Result<bool, ReadError> {
    let Some(val) = attr_value(e, b"val")? else {
        return Ok(true);
    };
    Ok(!(val == "0" || val.eq_ignore_ascii_case("false")))
}

fn attr_value(e: &BytesStart<'_>, key: &[u8]) -> Result<Option<String>, ReadError> {
    for attr in e.attributes().with_checks(false) {
        let attr = attr?;
        if attr.key.as_ref() == key {
            return Ok(Some(attr.unescape_value()?.into_owned()));
        }
    }
    Ok(None)
}

fn parse_size_100pt(val: &str) -> Option<u16> {
    let val = val.trim();
    if val.is_empty() {
        return None;
    }

    if let Some((int_part, frac_part)) = val.split_once('.') {
        let int: u16 = int_part.parse().ok()?;
        let mut frac = frac_part.chars().take(2).collect::<String>();
        while frac.len() < 2 {
            frac.push('0');
        }
        let frac: u16 = frac.parse().ok()?;
        int.checked_mul(100)?.checked_add(frac)
    } else {
        let int: u16 = val.parse().ok()?;
        int.checked_mul(100)
    }
}

fn parse_inline_rph_text<R: std::io::BufRead>(reader: &mut Reader<R>) -> Result<String, ReadError> {
    let mut buf = Vec::new();
    let mut out = String::new();
    loop {
        match reader.read_event_into(&mut buf)? {
            Event::Start(e) if e.local_name().as_ref() == b"t" => {
                out.push_str(&read_text(reader, b"t")?);
            }
            Event::Start(e) => {
                reader.read_to_end_into(e.name(), &mut Vec::new())?;
            }
            Event::End(e) if e.local_name().as_ref() == b"rPh" => break,
            Event::Eof => {
                return Err(ReadError::Xlsx(XlsxError::Invalid(
                    "unexpected EOF while parsing inline string <rPh>".to_string(),
                )))
            }
            _ => {}
        }
        buf.clear();
    }
    Ok(out)
}

fn read_text<R: std::io::BufRead>(
    reader: &mut Reader<R>,
    end_local: &[u8],
) -> Result<String, ReadError> {
    let mut buf = Vec::new();
    let mut text = String::new();
    loop {
        match reader.read_event_into(&mut buf)? {
            Event::Text(e) => text.push_str(&e.unescape()?.into_owned()),
            Event::CData(e) => text.push_str(std::str::from_utf8(e.as_ref())?),
            Event::End(e) if e.local_name().as_ref() == end_local => break,
            Event::Eof => {
                return Err(ReadError::Xlsx(XlsxError::Invalid(
                    "unexpected EOF while parsing inline string <t>".to_string(),
                )))
            }
            _ => {}
        }
        buf.clear();
    }
    Ok(text)
}

fn interpret_cell_value(
    t: Option<&str>,
    v_text: &Option<String>,
    inline_value: &Option<CellValue>,
    shared_strings: &[RichText],
) -> (CellValue, Option<CellValueKind>, Option<String>) {
    match t {
        Some("s") => {
            let raw = v_text.clone().unwrap_or_default();
            let idx: u32 = raw.parse().unwrap_or(0);
            let value = shared_strings
                .get(idx as usize)
                .map(|rich| {
                    if rich.runs.is_empty() {
                        CellValue::String(rich.text.clone())
                    } else {
                        CellValue::RichText(rich.clone())
                    }
                })
                .unwrap_or_else(|| CellValue::String(String::new()));
            (
                value,
                Some(CellValueKind::SharedString { index: idx }),
                Some(raw),
            )
        }
        Some("b") => {
            let raw = v_text.clone().unwrap_or_default();
            (
                CellValue::Boolean(raw == "1"),
                Some(CellValueKind::Bool),
                Some(raw),
            )
        }
        Some("e") => {
            let raw = v_text.clone().unwrap_or_default();
            let err = raw.parse::<ErrorValue>().unwrap_or(ErrorValue::Unknown);
            (CellValue::Error(err), Some(CellValueKind::Error), Some(raw))
        }
        Some("str") => {
            let raw = v_text.clone().unwrap_or_default();
            (
                CellValue::String(raw.clone()),
                Some(CellValueKind::Str),
                Some(raw),
            )
        }
        Some("inlineStr") => {
            let value = inline_value
                .clone()
                .unwrap_or_else(|| CellValue::String(String::new()));
            let raw = match &value {
                CellValue::String(s) => s.clone(),
                CellValue::RichText(rich) => rich.text.clone(),
                _ => String::new(),
            };
            (value, Some(CellValueKind::InlineString), Some(raw))
        }
        Some("n") | None => {
            if let Some(raw) = v_text.clone() {
                if let Ok(n) = raw.parse::<f64>() {
                    (CellValue::Number(n), Some(CellValueKind::Number), Some(raw))
                } else {
                    // A missing/number cell type with a non-numeric payload is invalid SpreadsheetML.
                    // Preserve as a plain string so we don't accidentally emit a numeric cell on write.
                    (
                        CellValue::String(raw.clone()),
                        Some(CellValueKind::Str),
                        Some(raw),
                    )
                }
            } else {
                (CellValue::Empty, None, None)
            }
        }
        Some(other) => {
            // Preserve unknown/less-common `t=` values (e.g. `t="d"` for ISO-8601 date text).
            if let Some(raw) = v_text.clone() {
                (
                    CellValue::String(raw.clone()),
                    Some(CellValueKind::Other {
                        t: other.to_string(),
                    }),
                    Some(raw),
                )
            } else {
                (
                    CellValue::Empty,
                    Some(CellValueKind::Other {
                        t: other.to_string(),
                    }),
                    None,
                )
            }
        }
    }
}

fn interpret_cell_value_without_meta(
    t: Option<&str>,
    v_text: &Option<String>,
    inline_value: &Option<CellValue>,
    shared_strings: &[RichText],
) -> CellValue {
    match t {
        Some("s") => {
            let raw = v_text.as_deref().unwrap_or_default();
            let idx: u32 = raw.parse().unwrap_or(0);
            shared_strings
                .get(idx as usize)
                .map(|rich| {
                    if rich.runs.is_empty() {
                        CellValue::String(rich.text.clone())
                    } else {
                        CellValue::RichText(rich.clone())
                    }
                })
                .unwrap_or_else(|| CellValue::String(String::new()))
        }
        Some("b") => CellValue::Boolean(v_text.as_deref() == Some("1")),
        Some("e") => {
            let raw = v_text.as_deref().unwrap_or_default();
            let err = raw.parse::<ErrorValue>().unwrap_or(ErrorValue::Unknown);
            CellValue::Error(err)
        }
        Some("str") => CellValue::String(v_text.clone().unwrap_or_default()),
        Some("inlineStr") => inline_value
            .clone()
            .unwrap_or_else(|| CellValue::String(String::new())),
        Some("n") | None => {
            if let Some(raw) = v_text.as_deref() {
                raw.parse::<f64>()
                    .map(CellValue::Number)
                    .unwrap_or_else(|_| CellValue::String(raw.to_string()))
            } else {
                CellValue::Empty
            }
        }
        Some(_) => {
            if let Some(raw) = v_text.as_deref() {
                CellValue::String(raw.to_string())
            } else {
                CellValue::Empty
            }
        }
    }
}

#[cfg(test)]
mod tests {
    use std::io::{Cursor, Write};

    use formula_model::CellRef;
    use formula_model::CellValue;
    use formula_model::ErrorValue;

    use super::{load_from_bytes, read_workbook_model_from_bytes};

    fn build_minimal_xlsx(sheet_xml: &str) -> Vec<u8> {
        let workbook_xml = r#"<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<workbook xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main"
 xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">
  <sheets>
    <sheet name="Sheet1" sheetId="1" r:id="rId1"/>
  </sheets>
</workbook>"#;

        let workbook_rels = r#"<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
  <Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/worksheet" Target="worksheets/sheet1.xml"/>
</Relationships>"#;

        let cursor = Cursor::new(Vec::new());
        let mut zip = zip::ZipWriter::new(cursor);
        let options = zip::write::FileOptions::<()>::default()
            .compression_method(zip::CompressionMethod::Deflated);

        zip.start_file("xl/workbook.xml", options).unwrap();
        zip.write_all(workbook_xml.as_bytes()).unwrap();

        zip.start_file("xl/_rels/workbook.xml.rels", options)
            .unwrap();
        zip.write_all(workbook_rels.as_bytes()).unwrap();

        zip.start_file("xl/worksheets/sheet1.xml", options).unwrap();
        zip.write_all(sheet_xml.as_bytes()).unwrap();

        zip.finish().unwrap().into_inner()
    }

    #[test]
    fn resolves_percent_encoded_relationship_targets_for_worksheets() {
        let workbook_xml = r#"<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<workbook xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main"
 xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">
  <sheets>
    <sheet name="Sheet1" sheetId="1" r:id="rId1"/>
  </sheets>
</workbook>"#;

        // The relationship target is percent-encoded, but the ZIP entry is stored unescaped.
        let workbook_rels = r#"<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
  <Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/worksheet" Target="worksheets/sheet%201.xml"/>
</Relationships>"#;

        let sheet_xml = r#"<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main">
  <dimension ref="A1"/>
  <sheetData>
    <row r="1">
      <c r="A1"><v>42</v></c>
    </row>
  </sheetData>
</worksheet>"#;

        let cursor = Cursor::new(Vec::new());
        let mut zip = zip::ZipWriter::new(cursor);
        let options = zip::write::FileOptions::<()>::default()
            .compression_method(zip::CompressionMethod::Deflated);

        zip.start_file("xl/workbook.xml", options).unwrap();
        zip.write_all(workbook_xml.as_bytes()).unwrap();

        zip.start_file("xl/_rels/workbook.xml.rels", options)
            .unwrap();
        zip.write_all(workbook_rels.as_bytes()).unwrap();

        // The part name uses a literal space.
        zip.start_file("xl/worksheets/sheet 1.xml", options)
            .unwrap();
        zip.write_all(sheet_xml.as_bytes()).unwrap();

        let bytes = zip.finish().unwrap().into_inner();

        let parts = crate::worksheet_parts_from_reader(Cursor::new(bytes.clone()))
            .expect("worksheet_parts_from_reader");
        assert_eq!(parts.len(), 1);
        assert_eq!(parts[0].worksheet_part, "xl/worksheets/sheet 1.xml");

        let workbook =
            read_workbook_model_from_bytes(&bytes).expect("read_workbook_model_from_bytes");
        assert_eq!(workbook.sheets.len(), 1);
        assert_eq!(workbook.sheets[0].name, "Sheet1");
        assert_eq!(
            workbook.sheets[0].value_a1("A1").unwrap(),
            CellValue::Number(42.0)
        );
    }

    #[test]
    fn reads_cell_cm_and_vm_attributes_into_cell_meta() {
        // The cell is otherwise empty, so `cm`/`vm` are the only reason it should appear in
        // `doc.meta.cell_meta`.
        let worksheet_xml = r#"<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main">
  <dimension ref="A1"/>
  <sheetData>
    <row r="1">
      <c r="A1" cm="7" vm="9"></c>
    </row>
  </sheetData>
</worksheet>"#;

        let bytes = build_minimal_xlsx(worksheet_xml);
        let mut doc = load_from_bytes(&bytes).expect("load_from_bytes");
        let sheet_id = doc.workbook.sheets[0].id;
        let cell_ref = CellRef::from_a1("A1").unwrap();

        let meta = doc
            .meta
            .cell_meta
            .get(&(sheet_id, cell_ref))
            .expect("expected cell meta entry for A1");
        assert_eq!(meta.cm.as_deref(), Some("7"));
        assert_eq!(meta.vm.as_deref(), Some("9"));

        // Ensure the higher-level editing API doesn't accidentally discard the metadata-only entry.
        doc.set_cell_value(sheet_id, cell_ref, CellValue::Empty);
        let meta = doc
            .meta
            .cell_meta
            .get(&(sheet_id, cell_ref))
            .expect("expected cell meta entry for A1 after set_cell_value(empty)");
        assert_eq!(meta.cm.as_deref(), Some("7"));
        assert_eq!(meta.vm.as_deref(), Some("9"));

        doc.set_cell_formula(sheet_id, cell_ref, None);
        let meta = doc
            .meta
            .cell_meta
            .get(&(sheet_id, cell_ref))
            .expect("expected cell meta entry for A1 after set_cell_formula(None)");
        assert_eq!(meta.cm.as_deref(), Some("7"));
        assert_eq!(meta.vm.as_deref(), Some("9"));
    }

    #[test]
    fn set_cell_value_clears_vm_when_cell_value_is_not_rich_value_placeholder() {
        let mut workbook = formula_model::Workbook::new();
        let sheet_id = workbook.add_sheet("Sheet1".to_string()).unwrap();
        let sheet = workbook.sheet_mut(sheet_id).expect("sheet exists");
        let a1 = CellRef::from_a1("A1").unwrap();
        sheet.set_value(a1, CellValue::Number(1.0));

        let mut doc = crate::XlsxDocument::new(workbook);
        doc.meta.cell_meta.insert(
            (sheet_id, a1),
            crate::CellMeta {
                vm: Some("1".to_string()),
                cm: Some("2".to_string()),
                ..Default::default()
            },
        );

        // Keep the cell value unchanged so `value_changed` stays false and we exercise the
        // placeholder retention logic in `set_cell_value`.
        doc.set_cell_value(sheet_id, a1, CellValue::Number(1.0));
        let meta = doc.cell_meta(sheet_id, a1).expect("cell meta exists");
        assert_eq!(
            meta.vm, None,
            "expected vm to be dropped for non-placeholder values"
        );
        assert_eq!(
            meta.cm.as_deref(),
            Some("2"),
            "expected cm metadata to be preserved"
        );
    }

    #[test]
    fn set_cell_value_preserves_vm_when_cell_value_is_rich_value_placeholder() {
        let mut workbook = formula_model::Workbook::new();
        let sheet_id = workbook.add_sheet("Sheet1".to_string()).unwrap();
        let sheet = workbook.sheet_mut(sheet_id).expect("sheet exists");
        let a1 = CellRef::from_a1("A1").unwrap();
        sheet.set_value(a1, CellValue::Error(ErrorValue::Value));

        let mut doc = crate::XlsxDocument::new(workbook);
        doc.meta.cell_meta.insert(
            (sheet_id, a1),
            crate::CellMeta {
                vm: Some("1".to_string()),
                cm: Some("2".to_string()),
                ..Default::default()
            },
        );

        // Ensure the vm metadata survives when the cell retains the `#VALUE!` placeholder used for
        // rich values (e.g. images-in-cell).
        doc.set_cell_value(sheet_id, a1, CellValue::Error(ErrorValue::Value));
        let meta = doc.cell_meta(sheet_id, a1).expect("cell meta exists");
        assert_eq!(
            meta.vm.as_deref(),
            Some("1"),
            "expected vm to be preserved for rich-value placeholders"
        );
        assert_eq!(
            meta.cm.as_deref(),
            Some("2"),
            "expected cm metadata to be preserved"
        );
    }

    #[test]
    fn ignores_external_workbook_relationships_for_metadata_part() {
        let workbook_xml = r#"<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<workbook xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main"
 xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">
  <sheets>
    <sheet name="Sheet1" sheetId="1" r:id="rId1"/>
  </sheets>
</workbook>"#;

        // External metadata relationship is listed first and should be ignored. Otherwise we would
        // attempt to resolve `https://...` as a package part name and fail to load the real
        // `xl/metadata.xml`.
        let workbook_rels = r#"<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
  <Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/worksheet" Target="worksheets/sheet1.xml"/>
  <Relationship Id="rId2" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/metadata" Target="https://example.com/metadata.xml" TargetMode="External"/>
  <Relationship Id="rId3" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/metadata" Target="metadata.xml"/>
</Relationships>"#;

        // Minimal metadata.xml using the "direct" mapping variant (`rc/@v` stores the rich value
        // index directly).
        let metadata_xml = r#"<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<metadata xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main">
  <metadataTypes count="1">
    <metadataType name="XLRICHVALUE"/>
  </metadataTypes>
  <valueMetadata count="1">
    <bk><rc t="1" v="42"/></bk>
  </valueMetadata>
</metadata>"#;

        // vm="1" should map to rich value index 42.
        let sheet_xml = r#"<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main">
  <dimension ref="A1"/>
  <sheetData>
    <row r="1">
      <c r="A1" vm="1"/>
    </row>
  </sheetData>
</worksheet>"#;

        let cursor = Cursor::new(Vec::new());
        let mut zip = zip::ZipWriter::new(cursor);
        let options = zip::write::FileOptions::<()>::default()
            .compression_method(zip::CompressionMethod::Deflated);

        zip.start_file("xl/workbook.xml", options).unwrap();
        zip.write_all(workbook_xml.as_bytes()).unwrap();

        zip.start_file("xl/_rels/workbook.xml.rels", options)
            .unwrap();
        zip.write_all(workbook_rels.as_bytes()).unwrap();

        zip.start_file("xl/metadata.xml", options).unwrap();
        zip.write_all(metadata_xml.as_bytes()).unwrap();

        zip.start_file("xl/worksheets/sheet1.xml", options).unwrap();
        zip.write_all(sheet_xml.as_bytes()).unwrap();

        let bytes = zip.finish().unwrap().into_inner();

        let doc = load_from_bytes(&bytes).expect("load_from_bytes");
        let sheet_id = doc.workbook.sheets[0].id;
        let cell_ref = CellRef::from_a1("A1").unwrap();
        assert_eq!(
            doc.xlsx_meta().rich_value_cells.get(&(sheet_id, cell_ref)),
            Some(&42)
        );
    }
}
