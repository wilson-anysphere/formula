use std::io::Write;
use std::path::Path;

use formula_model::{CellRef, CellValue, DateSystem, DefinedNameScope, Workbook};
use zip::write::FileOptions;
use zip::{CompressionMethod, ZipArchive, ZipWriter};

fn fixture_path(rel: &str) -> std::path::PathBuf {
    Path::new(env!("CARGO_MANIFEST_DIR")).join("../../").join(rel)
}

fn assert_cell_matches(full: &Workbook, fast: &Workbook, sheet_idx: usize, a1: &str) {
    let cell_ref = CellRef::from_a1(a1).expect("valid cell ref");
    let full_sheet = &full.sheets[sheet_idx];
    let fast_sheet = &fast.sheets[sheet_idx];

    let full_cell = full_sheet.cell(cell_ref);
    let fast_cell = fast_sheet.cell(cell_ref);

    assert_eq!(
        full_cell.map(|c| &c.value),
        fast_cell.map(|c| &c.value),
        "cell value mismatch for {}!{}",
        full_sheet.name,
        a1
    );
    assert_eq!(
        full_cell.and_then(|c| c.formula.as_deref()),
        fast_cell.and_then(|c| c.formula.as_deref()),
        "cell formula mismatch for {}!{}",
        full_sheet.name,
        a1
    );
    assert_eq!(
        full_cell.map(|c| c.style_id),
        fast_cell.map(|c| c.style_id),
        "cell style_id mismatch for {}!{}",
        full_sheet.name,
        a1
    );
}

#[test]
fn fast_reader_matches_full_reader_for_values_and_formulas() {
    struct Case<'a> {
        fixture: &'a str,
        sheet_idx: usize,
        cells: &'a [&'a str],
    }

    let cases = [
        Case {
            fixture: "fixtures/xlsx/basic/basic.xlsx",
            sheet_idx: 0,
            cells: &["A1", "B1"],
        },
        Case {
            fixture: "fixtures/xlsx/formulas/formulas.xlsx",
            sheet_idx: 0,
            cells: &["A1", "B1", "C1"],
        },
        Case {
            fixture: "fixtures/xlsx/styles/rich-text-shared-strings.xlsx",
            sheet_idx: 0,
            cells: &["A1", "A2"],
        },
        // Explicitly exercises `styles.xml` + `c/@s` -> `style_id` mapping.
        Case {
            fixture: "fixtures/xlsx/styles/styles.xlsx",
            sheet_idx: 0,
            cells: &["A1"],
        },
    ];

    for case in cases {
        let bytes = std::fs::read(fixture_path(case.fixture)).expect("read fixture");
        let full = formula_xlsx::load_from_bytes(&bytes)
            .expect("load_from_bytes")
            .workbook;
        let fast = formula_xlsx::read_workbook_model_from_bytes(&bytes)
            .expect("fast reader should succeed");

        let full_names: Vec<_> = full.sheets.iter().map(|s| s.name.as_str()).collect();
        let fast_names: Vec<_> = fast.sheets.iter().map(|s| s.name.as_str()).collect();
        assert_eq!(
            full_names, fast_names,
            "sheet list mismatch for fixture {}",
            case.fixture
        );

        assert_eq!(
            full.styles.len(),
            fast.styles.len(),
            "style table size mismatch for fixture {}",
            case.fixture
        );

        for cell in case.cells {
            assert_cell_matches(&full, &fast, case.sheet_idx, cell);
        }

        // Spot-check expected semantics for the key fixtures so this test catches
        // regressions even if both readers drift together.
        if case.fixture.ends_with("basic.xlsx") {
            assert_eq!(
                full.sheets[0]
                    .cell(CellRef::from_a1("A1").unwrap())
                    .unwrap()
                    .value,
                CellValue::Number(1.0)
            );
            assert_eq!(
                full.sheets[0]
                    .cell(CellRef::from_a1("B1").unwrap())
                    .unwrap()
                    .value,
                CellValue::String("Hello".to_string())
            );
        }

        if case.fixture.ends_with("formulas.xlsx") {
            assert_eq!(
                fast.sheets[0]
                    .cell(CellRef::from_a1("C1").unwrap())
                    .unwrap()
                    .formula
                    .as_deref(),
                Some("A1+B1")
            );
        }

        if case.fixture.ends_with("rich-text-shared-strings.xlsx") {
            let value = &fast.sheets[0]
                .cell(CellRef::from_a1("A1").unwrap())
                .unwrap()
                .value;
            match value {
                CellValue::RichText(rich) => {
                    assert_eq!(rich.text, "Hello Bold Italic");
                    assert!(
                        !rich.runs.is_empty(),
                        "expected rich text runs to be preserved for shared string"
                    );
                }
                other => panic!("expected A1 to be rich text, got {other:?}"),
            }
        }
    }
}

#[test]
fn fast_reader_does_not_require_unmodeled_parts() {
    let bytes =
        std::fs::read(fixture_path("fixtures/charts/xlsx/basic-chart.xlsx")).expect("read fixture");
    let workbook = formula_xlsx::read_workbook_model_from_bytes(&bytes)
        .expect("fast reader should ignore chart/drawing parts");
    assert_eq!(workbook.sheets.len(), 1);
    assert_eq!(workbook.sheets[0].name, "Sheet1");
}

fn normalized_defined_names(
    workbook: &Workbook,
) -> Vec<(
    Option<String>,
    String,
    String,
    Option<String>,
    bool,
    Option<u32>,
)> {
    let mut out = Vec::new();
    for dn in &workbook.defined_names {
        let scope = match dn.scope {
            DefinedNameScope::Workbook => None,
            DefinedNameScope::Sheet(sheet_id) => workbook.sheet(sheet_id).map(|s| s.name.clone()),
        };
        out.push((
            scope,
            dn.name.clone(),
            dn.refers_to.clone(),
            dn.comment.clone(),
            dn.hidden,
            dn.xlsx_local_sheet_id,
        ));
    }
    out.sort_by(|a, b| (a.0.as_deref(), a.1.as_str()).cmp(&(b.0.as_deref(), b.1.as_str())));
    out
}

#[test]
fn fast_reader_matches_full_reader_for_defined_names() -> Result<(), Box<dyn std::error::Error>> {
    let mut workbook = Workbook::new();
    let sheet1 = workbook.add_sheet("Sheet1")?;
    let sheet2 = workbook.add_sheet("Sheet2")?;

    workbook.create_defined_name(
        DefinedNameScope::Workbook,
        "MyX",
        "Sheet1!A1",
        None,
        false,
        None,
    )?;
    workbook.create_defined_name(
        DefinedNameScope::Sheet(sheet2),
        "LocalFoo",
        // Leading '=' should be stripped on read/write.
        "=Sheet2!B2",
        Some("example comment".to_string()),
        true,
        None,
    )?;

    // Ensure ids from the initial workbook don't leak into output comparisons.
    assert_ne!(sheet1, sheet2);

    let mut buf = std::io::Cursor::new(Vec::new());
    formula_xlsx::write_workbook_to_writer(&workbook, &mut buf)?;
    let bytes = buf.into_inner();

    let full = formula_xlsx::load_from_bytes(&bytes)?.workbook;
    let fast = formula_xlsx::read_workbook_model_from_bytes(&bytes)?;

    assert_eq!(normalized_defined_names(&full), normalized_defined_names(&fast));
    Ok(())
}

fn with_workbook_pr_date1904(bytes: &[u8]) -> Vec<u8> {
    let cursor = std::io::Cursor::new(bytes);
    let mut archive = ZipArchive::new(cursor).expect("parse zip");

    let mut parts: Vec<(String, Vec<u8>)> = Vec::new();
    for i in 0..archive.len() {
        let mut file = archive.by_index(i).expect("zip entry");
        if file.is_dir() {
            continue;
        }
        let name = file.name().to_string();
        // Do not trust `ZipFile::size()` for allocation; ZIP metadata is untrusted and can
        // advertise enormous uncompressed sizes (zip-bomb style OOM).
        let mut buf = Vec::new();
        std::io::Read::read_to_end(&mut file, &mut buf).expect("read zip entry");
        if name == "xl/workbook.xml" {
            let mut xml = String::from_utf8(buf).expect("workbook.xml utf8");
            if !xml.contains("date1904=") {
                let start = xml.find("<workbook").expect("workbook root element");
                let end = xml[start..]
                    .find('>')
                    .map(|idx| start + idx)
                    .expect("workbook start tag terminator");
                xml.insert_str(end + 1, "\n  <workbookPr date1904=\"1\"/>\n");
            }
            buf = xml.into_bytes();
        }
        parts.push((name, buf));
    }

    let mut out = std::io::Cursor::new(Vec::new());
    let mut writer = ZipWriter::new(&mut out);
    let options = FileOptions::<()>::default().compression_method(CompressionMethod::Stored);
    for (name, buf) in parts {
        writer.start_file(name, options).expect("write zip header");
        writer.write_all(&buf).expect("write zip body");
    }
    writer.finish().expect("finalize zip");
    out.into_inner()
}

#[test]
fn fast_reader_sets_workbook_date_system_from_workbook_pr() {
    let bytes = std::fs::read(fixture_path("fixtures/xlsx/basic/basic.xlsx")).expect("read fixture");
    let bytes = with_workbook_pr_date1904(&bytes);

    let full = formula_xlsx::load_from_bytes(&bytes)
        .expect("load_from_bytes")
        .workbook;
    let fast =
        formula_xlsx::read_workbook_model_from_bytes(&bytes).expect("fast reader should succeed");

    assert_eq!(full.date_system, DateSystem::Excel1904);
    assert_eq!(fast.date_system, DateSystem::Excel1904);
}

#[test]
fn fast_reader_discovers_table_parts() {
    let path = Path::new(env!("CARGO_MANIFEST_DIR")).join("tests/fixtures/table.xlsx");
    let bytes = std::fs::read(&path).expect("read table fixture");
    let workbook =
        formula_xlsx::read_workbook_model_from_bytes(&bytes).expect("fast reader should succeed");

    assert_eq!(workbook.sheets.len(), 1);
    let sheet = &workbook.sheets[0];
    assert_eq!(sheet.tables.len(), 1);

    let table = &sheet.tables[0];
    assert_eq!(table.name, "Table1");
    assert_eq!(table.range.to_string(), "A1:D4");
    assert_eq!(table.relationship_id.as_deref(), Some("rId1"));
    assert_eq!(table.part_path.as_deref(), Some("xl/tables/table1.xml"));
}
