use formula_model::pivots::{
    DefinedNameIdentifier, PivotDestination, PivotSource, PivotTableModel,
};
use formula_model::table::{Table, TableColumn, TableIdentifier};
use formula_model::{CellRef, DefinedNameScope, Range, Workbook};

use uuid::Uuid;

fn sample_table(name: &str) -> Table {
    Table {
        id: 1,
        name: name.to_string(),
        display_name: name.to_string(),
        range: Range::from_a1("A1:B3").unwrap(),
        header_row_count: 1,
        totals_row_count: 0,
        columns: vec![
            TableColumn {
                id: 1,
                name: "A".to_string(),
                formula: None,
                totals_formula: None,
            },
            TableColumn {
                id: 2,
                name: "B".to_string(),
                formula: None,
                totals_formula: None,
            },
        ],
        style: None,
        auto_filter: None,
        relationship_id: None,
        part_path: None,
    }
}

#[test]
fn rename_table_rewrites_pivot_table_sources() {
    let mut wb = Workbook::new();
    let sheet_id = wb.add_sheet("Sheet1").unwrap();
    wb.add_table(sheet_id, sample_table("Table1")).unwrap();

    wb.pivot_tables.push(PivotTableModel {
        id: Uuid::from_u128(1),
        name: "Pivot1".to_string(),
        source: PivotSource::Table {
            table: TableIdentifier::Name("TABLE1".to_string()),
        },
        destination: PivotDestination::Cell {
            sheet_id,
            cell: CellRef::new(0, 0),
        },
        config: Default::default(),
        cache_id: None,
    });
    wb.pivot_tables.push(PivotTableModel {
        id: Uuid::from_u128(2),
        name: "Pivot2".to_string(),
        source: PivotSource::Table {
            table: TableIdentifier::Id(1),
        },
        destination: PivotDestination::Cell {
            sheet_id,
            cell: CellRef::new(10, 0),
        },
        config: Default::default(),
        cache_id: None,
    });

    wb.rename_table("table1", "RenamedTable").unwrap();

    assert_eq!(
        wb.pivot_tables[0].source,
        PivotSource::Table {
            table: TableIdentifier::Name("RenamedTable".to_string())
        }
    );
    // Stable id references should not be rewritten.
    assert_eq!(
        wb.pivot_tables[1].source,
        PivotSource::Table {
            table: TableIdentifier::Id(1)
        }
    );
}

#[test]
fn rename_defined_name_rewrites_pivot_named_range_sources() {
    let mut wb = Workbook::new();
    let sheet_id = wb.add_sheet("Sheet1").unwrap();
    let id = wb
        .create_defined_name(
            DefinedNameScope::Workbook,
            "MyRange",
            "Sheet1!A1:B2",
            None,
            false,
            None,
        )
        .unwrap();

    wb.pivot_tables.push(PivotTableModel {
        id: Uuid::from_u128(1),
        name: "Pivot1".to_string(),
        source: PivotSource::NamedRange {
            name: DefinedNameIdentifier::Name("MYRANGE".to_string()),
        },
        destination: PivotDestination::Cell {
            sheet_id,
            cell: CellRef::new(0, 0),
        },
        config: Default::default(),
        cache_id: None,
    });
    wb.pivot_tables.push(PivotTableModel {
        id: Uuid::from_u128(2),
        name: "Pivot2".to_string(),
        source: PivotSource::NamedRange {
            name: DefinedNameIdentifier::Id(id),
        },
        destination: PivotDestination::Cell {
            sheet_id,
            cell: CellRef::new(10, 0),
        },
        config: Default::default(),
        cache_id: None,
    });

    wb.rename_defined_name(id, "RenamedRange").unwrap();

    assert_eq!(
        wb.pivot_tables[0].source,
        PivotSource::NamedRange {
            name: DefinedNameIdentifier::Name("RenamedRange".to_string())
        }
    );
    // Stable id references should not be rewritten.
    assert_eq!(
        wb.pivot_tables[1].source,
        PivotSource::NamedRange {
            name: DefinedNameIdentifier::Id(id)
        }
    );
}

#[test]
fn rename_sheet_rewrites_string_based_sheet_refs_in_pivots() {
    let mut wb = Workbook::new();
    let sheet_id = wb.add_sheet("Data").unwrap();

    wb.pivot_tables.push(PivotTableModel {
        id: Uuid::from_u128(1),
        name: "Pivot1".to_string(),
        source: PivotSource::RangeName {
            sheet_name: "DATA".to_string(),
            range: Range::from_a1("A1:C10").unwrap(),
        },
        destination: PivotDestination::CellName {
            sheet_name: "Data".to_string(),
            cell: CellRef::new(0, 0),
        },
        config: Default::default(),
        cache_id: None,
    });

    wb.rename_sheet(sheet_id, "Renamed").unwrap();

    assert_eq!(
        wb.pivot_tables[0].source,
        PivotSource::RangeName {
            sheet_name: "Renamed".to_string(),
            range: Range::from_a1("A1:C10").unwrap(),
        }
    );
    assert_eq!(
        wb.pivot_tables[0].destination,
        PivotDestination::CellName {
            sheet_name: "Renamed".to_string(),
            cell: CellRef::new(0, 0),
        }
    );
}

#[test]
fn rename_operations_are_noops_when_no_pivots_exist() {
    let mut wb = Workbook::new();
    let sheet_id = wb.add_sheet("Sheet1").unwrap();
    wb.add_table(sheet_id, sample_table("Table1")).unwrap();
    let name_id = wb
        .create_defined_name(
            DefinedNameScope::Workbook,
            "MyRange",
            "Sheet1!A1",
            None,
            false,
            None,
        )
        .unwrap();

    assert!(wb.pivot_tables.is_empty());

    wb.rename_sheet(sheet_id, "RenamedSheet").unwrap();
    wb.rename_table("Table1", "RenamedTable").unwrap();
    wb.rename_defined_name(name_id, "RenamedRange").unwrap();

    assert!(wb.pivot_tables.is_empty());
}

#[test]
fn duplicate_sheet_duplicates_pivots_and_updates_destinations_and_sources() {
    let mut wb = Workbook::new();
    let sheet_id = wb.add_sheet("Sheet1").unwrap();
    wb.add_table(sheet_id, sample_table("Table1")).unwrap();
    let range = Range::from_a1("A1:B3").unwrap();

    wb.pivot_tables.push(PivotTableModel {
        id: Uuid::from_u128(1),
        name: "Pivot1".to_string(),
        source: PivotSource::Table {
            table: TableIdentifier::Id(1),
        },
        destination: PivotDestination::Cell {
            sheet_id,
            cell: CellRef::new(0, 0),
        },
        config: Default::default(),
        cache_id: Some(Uuid::from_u128(1234)),
    });
    wb.pivot_tables.push(PivotTableModel {
        id: Uuid::from_u128(2),
        name: "Pivot2".to_string(),
        source: PivotSource::RangeName {
            sheet_name: "Sheet1".to_string(),
            range: range.clone(),
        },
        destination: PivotDestination::CellName {
            sheet_name: "SHEET1".to_string(),
            cell: CellRef::new(10, 0),
        },
        config: Default::default(),
        cache_id: None,
    });

    let duplicated_sheet_id = wb.duplicate_sheet(sheet_id, Some("Copy")).unwrap();

    // We should now have 2 original pivots + 2 duplicates.
    assert_eq!(wb.pivot_tables.len(), 4);

    // Original pivot still points at original sheet/table.
    assert!(wb.pivot_tables.iter().any(|p| {
        p.id == Uuid::from_u128(1)
            && p.destination
                == PivotDestination::Cell {
                    sheet_id,
                    cell: CellRef::new(0, 0),
                }
            && p.source
                == PivotSource::Table {
                    table: TableIdentifier::Id(1),
                }
    }));

    // Original string-based pivot remains unchanged.
    assert!(wb.pivot_tables.iter().any(|p| {
        p.id == Uuid::from_u128(2)
            && p.destination
                == PivotDestination::CellName {
                    sheet_name: "SHEET1".to_string(),
                    cell: CellRef::new(10, 0),
                }
            && p.source
                == PivotSource::RangeName {
                    sheet_name: "Sheet1".to_string(),
                    range: range.clone(),
                }
    }));

    // The duplicated pivots should target the copied sheet and split their caches so they can be
    // refreshed independently.
    let duplicated_table_pivot = wb
        .pivot_tables
        .iter()
        .find(|p| {
            p.destination
                == PivotDestination::Cell {
                    sheet_id: duplicated_sheet_id,
                    cell: CellRef::new(0, 0),
                }
                && p.source
                    == PivotSource::Table {
                        table: TableIdentifier::Id(2),
                    }
        })
        .expect("expected duplicated table-based pivot");
    let duplicated_table_cache_id = duplicated_table_pivot
        .cache_id
        .expect("duplicated pivot should have its own cache id");
    assert!(wb.pivot_caches.iter().any(|cache| {
        cache.id == duplicated_table_cache_id
            && cache.needs_refresh
            && cache.source == duplicated_table_pivot.source
    }));

    let duplicated_range_pivot = wb
        .pivot_tables
        .iter()
        .find(|p| {
            p.destination
                == PivotDestination::CellName {
                    sheet_name: "Copy".to_string(),
                    cell: CellRef::new(10, 0),
                }
                && p.source
                    == PivotSource::RangeName {
                        sheet_name: "Copy".to_string(),
                        range: range.clone(),
                    }
        })
        .expect("expected duplicated range-based pivot");
    let duplicated_range_cache_id = duplicated_range_pivot
        .cache_id
        .expect("duplicated pivot should have its own cache id");
    assert!(wb.pivot_caches.iter().any(|cache| {
        cache.id == duplicated_range_cache_id
            && cache.needs_refresh
            && cache.source == duplicated_range_pivot.source
    }));

    assert_eq!(wb.pivot_caches.len(), 2);
}
