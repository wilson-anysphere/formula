use formula_model::pivots::{
    DefinedNameIdentifier, PivotCacheModel, PivotChartModel, PivotConfig, PivotDestination,
    PivotSource, PivotTableModel, SlicerModel,
};
use formula_model::{
    CellRef, DefinedNameScope, Range, Table, TableColumn, TableIdentifier, Workbook,
};

#[test]
fn duplicate_sheet_duplicates_pivot_tables_and_rewrites_sources() {
    let mut wb = Workbook::new();
    let sheet1 = wb.add_sheet("Sheet1").unwrap();

    // Add a table so we can validate table-backed pivot sources are rewritten.
    wb.sheet_mut(sheet1).unwrap().tables.push(Table {
        id: 1,
        name: "Table1".to_string(),
        display_name: "Table1".to_string(),
        range: Range::from_a1("A1:B3").unwrap(),
        header_row_count: 1,
        totals_row_count: 0,
        columns: vec![
            TableColumn {
                id: 1,
                name: "Col1".to_string(),
                formula: None,
                totals_formula: None,
            },
            TableColumn {
                id: 2,
                name: "Col2".to_string(),
                formula: None,
                totals_formula: None,
            },
        ],
        style: None,
        auto_filter: None,
        relationship_id: None,
        part_path: None,
    });

    // Workbook-scoped named-range source used to document/lock in our duplication behavior.
    let name_id = wb
        .create_defined_name(
            DefinedNameScope::Workbook,
            "MyData",
            "=Sheet1!$A$1:$A$10",
            None,
            false,
            None,
        )
        .unwrap();

    // Sheet-scoped named-range source should be rewritten to point at the duplicated defined name.
    let local_name_id = wb
        .create_defined_name(
            DefinedNameScope::Sheet(sheet1),
            "LocalData",
            "=Sheet1!$B$1:$B$10",
            None,
            false,
            None,
        )
        .unwrap();

    let range = Range::from_a1("A1:B10").unwrap();

    let cache_range_id = uuid::Uuid::from_u128(1);
    wb.pivot_caches.push(PivotCacheModel {
        id: cache_range_id,
        source: PivotSource::Range {
            sheet_id: sheet1,
            range,
        },
        needs_refresh: false,
    });
    let pivot1_id = uuid::Uuid::from_u128(11);
    wb.pivot_tables.push(PivotTableModel {
        id: pivot1_id,
        name: "PivotTable1".to_string(),
        source: PivotSource::Range {
            sheet_id: sheet1,
            range,
        },
        destination: PivotDestination::Cell {
            sheet_id: sheet1,
            cell: CellRef::from_a1("D1").unwrap(),
        },
        config: PivotConfig::default(),
        cache_id: Some(cache_range_id),
    });

    let cache_table_id = uuid::Uuid::from_u128(2);
    wb.pivot_caches.push(PivotCacheModel {
        id: cache_table_id,
        source: PivotSource::Table {
            table: TableIdentifier::Id(1),
        },
        needs_refresh: false,
    });
    let pivot2_id = uuid::Uuid::from_u128(12);
    wb.pivot_tables.push(PivotTableModel {
        id: pivot2_id,
        name: "PivotTable2".to_string(),
        source: PivotSource::Table {
            table: TableIdentifier::Id(1),
        },
        destination: PivotDestination::Cell {
            sheet_id: sheet1,
            cell: CellRef::from_a1("D10").unwrap(),
        },
        config: PivotConfig::default(),
        cache_id: Some(cache_table_id),
    });

    let cache_name_id = uuid::Uuid::from_u128(3);
    wb.pivot_caches.push(PivotCacheModel {
        id: cache_name_id,
        source: PivotSource::NamedRange {
            name: DefinedNameIdentifier::Id(name_id),
        },
        needs_refresh: false,
    });
    wb.pivot_tables.push(PivotTableModel {
        id: uuid::Uuid::from_u128(13),
        name: "PivotTable3".to_string(),
        source: PivotSource::NamedRange {
            name: DefinedNameIdentifier::Id(name_id),
        },
        destination: PivotDestination::Cell {
            sheet_id: sheet1,
            cell: CellRef::from_a1("D20").unwrap(),
        },
        config: PivotConfig::default(),
        cache_id: Some(cache_name_id),
    });

    let cache_local_name_id = uuid::Uuid::from_u128(4);
    wb.pivot_caches.push(PivotCacheModel {
        id: cache_local_name_id,
        source: PivotSource::NamedRange {
            name: DefinedNameIdentifier::Id(local_name_id),
        },
        needs_refresh: false,
    });
    wb.pivot_tables.push(PivotTableModel {
        id: uuid::Uuid::from_u128(14),
        name: "PivotTable4".to_string(),
        source: PivotSource::NamedRange {
            name: DefinedNameIdentifier::Id(local_name_id),
        },
        destination: PivotDestination::Cell {
            sheet_id: sheet1,
            cell: CellRef::from_a1("D30").unwrap(),
        },
        config: PivotConfig::default(),
        cache_id: Some(cache_local_name_id),
    });

    wb.pivot_charts.push(PivotChartModel {
        id: uuid::Uuid::from_u128(21),
        name: "PivotChart1".to_string(),
        pivot_table_id: pivot1_id,
        chart_part: None,
        sheet_id: Some(sheet1),
    });

    wb.slicers.push(SlicerModel {
        id: uuid::Uuid::from_u128(31),
        name: "Slicer1".to_string(),
        connected_pivots: vec![pivot1_id, pivot2_id],
        sheet_id: sheet1,
    });

    let copied_sheet = wb.duplicate_sheet(sheet1, None).unwrap();
    let copied_table_id = wb.sheet(copied_sheet).unwrap().tables[0].id;
    assert_ne!(copied_table_id, 1);

    // We duplicated 4 pivots on the source sheet.
    assert_eq!(wb.pivot_tables.len(), 8);
    // Three sources were rewritten (range + table + sheet-scoped name), so we allocate 3 new caches.
    assert_eq!(wb.pivot_caches.len(), 7);

    let p1_copy = wb
        .pivot_tables
        .iter()
        .find(|p| p.name == "PivotTable1 (2)")
        .expect("expected duplicated PivotTable1");
    assert_eq!(
        p1_copy.destination,
        PivotDestination::Cell {
            sheet_id: copied_sheet,
            cell: CellRef::from_a1("D1").unwrap()
        }
    );
    let p1_copy_cache_id = p1_copy.cache_id.expect("expected cache id");
    assert_ne!(p1_copy_cache_id, cache_range_id);
    assert_eq!(
        p1_copy.source,
        PivotSource::Range {
            sheet_id: copied_sheet,
            range
        }
    );
    let p1_cache = wb
        .pivot_caches
        .iter()
        .find(|c| c.id == p1_copy_cache_id)
        .expect("expected duplicated cache for PivotTable1");
    assert!(p1_cache.needs_refresh);
    assert_eq!(p1_cache.source, p1_copy.source);

    let p2_copy = wb
        .pivot_tables
        .iter()
        .find(|p| p.name == "PivotTable2 (2)")
        .expect("expected duplicated PivotTable2");
    assert_eq!(
        p2_copy.destination,
        PivotDestination::Cell {
            sheet_id: copied_sheet,
            cell: CellRef::from_a1("D10").unwrap()
        }
    );
    let p2_copy_cache_id = p2_copy.cache_id.expect("expected cache id");
    assert_ne!(p2_copy_cache_id, cache_table_id);
    assert_eq!(
        p2_copy.source,
        PivotSource::Table {
            table: TableIdentifier::Id(copied_table_id)
        }
    );
    let p2_cache = wb
        .pivot_caches
        .iter()
        .find(|c| c.id == p2_copy_cache_id)
        .expect("expected duplicated cache for PivotTable2");
    assert!(p2_cache.needs_refresh);
    assert_eq!(p2_cache.source, p2_copy.source);

    let p3_copy = wb
        .pivot_tables
        .iter()
        .find(|p| p.name == "PivotTable3 (2)")
        .expect("expected duplicated PivotTable3");
    assert_eq!(
        p3_copy.destination,
        PivotDestination::Cell {
            sheet_id: copied_sheet,
            cell: CellRef::from_a1("D20").unwrap()
        }
    );
    // Workbook-scoped named-range sources are not rewritten during sheet duplication.
    assert_eq!(p3_copy.cache_id, Some(cache_name_id));
    assert_eq!(
        p3_copy.source,
        PivotSource::NamedRange {
            name: DefinedNameIdentifier::Id(name_id)
        }
    );

    let local_name_copy_id = wb
        .get_defined_name(DefinedNameScope::Sheet(copied_sheet), "LocalData")
        .expect("expected duplicated sheet-scoped name")
        .id;
    assert_ne!(local_name_copy_id, local_name_id);
    let p4_copy = wb
        .pivot_tables
        .iter()
        .find(|p| p.name == "PivotTable4 (2)")
        .expect("expected duplicated PivotTable4");
    assert_eq!(
        p4_copy.destination,
        PivotDestination::Cell {
            sheet_id: copied_sheet,
            cell: CellRef::from_a1("D30").unwrap()
        }
    );
    let p4_copy_cache_id = p4_copy.cache_id.expect("expected cache id");
    assert_ne!(p4_copy_cache_id, cache_local_name_id);
    assert_eq!(
        p4_copy.source,
        PivotSource::NamedRange {
            name: DefinedNameIdentifier::Id(local_name_copy_id)
        }
    );
    let p4_cache = wb
        .pivot_caches
        .iter()
        .find(|c| c.id == p4_copy_cache_id)
        .expect("expected duplicated cache for PivotTable4");
    assert!(p4_cache.needs_refresh);
    assert_eq!(p4_cache.source, p4_copy.source);

    assert_eq!(wb.pivot_charts.len(), 2);
    let chart_copy = wb
        .pivot_charts
        .iter()
        .find(|c| c.sheet_id == Some(copied_sheet))
        .expect("expected duplicated pivot chart");
    assert_eq!(chart_copy.pivot_table_id, p1_copy.id);

    assert_eq!(wb.slicers.len(), 2);
    let slicer_copy = wb
        .slicers
        .iter()
        .find(|s| s.sheet_id == copied_sheet)
        .expect("expected duplicated slicer");
    assert_eq!(slicer_copy.connected_pivots, vec![p1_copy.id, p2_copy.id]);
}

#[test]
fn duplicate_sheet_generates_unique_pivot_names_case_insensitively_for_unicode_text() {
    let mut wb = Workbook::new();
    let data = wb.add_sheet("Data").unwrap();
    let report = wb.add_sheet("Report").unwrap();
    let other = wb.add_sheet("Other").unwrap();

    wb.pivot_tables.push(PivotTableModel {
        id: uuid::Uuid::from_u128(1),
        name: "Straße".to_string(),
        source: PivotSource::Range {
            sheet_id: data,
            range: Range::from_a1("A1:A2").unwrap(),
        },
        destination: PivotDestination::Cell {
            sheet_id: report,
            cell: CellRef::from_a1("A1").unwrap(),
        },
        config: PivotConfig::default(),
        cache_id: None,
    });

    // Pre-seed a pivot name that would collide with the default duplicated name `"Straße (2)"`
    // under Unicode-aware case folding (ß -> SS).
    wb.pivot_tables.push(PivotTableModel {
        id: uuid::Uuid::from_u128(2),
        name: "STRASSE (2)".to_string(),
        source: PivotSource::Range {
            sheet_id: data,
            range: Range::from_a1("A1:A2").unwrap(),
        },
        destination: PivotDestination::Cell {
            sheet_id: other,
            cell: CellRef::from_a1("A1").unwrap(),
        },
        config: PivotConfig::default(),
        cache_id: None,
    });

    let copied_sheet = wb.duplicate_sheet(report, None).unwrap();

    // Only the pivot on the duplicated sheet is copied.
    assert_eq!(wb.pivot_tables.len(), 3);

    let duplicated = wb
        .pivot_tables
        .iter()
        .find(|p| {
            matches!(
                &p.destination,
                PivotDestination::Cell { sheet_id, .. } if *sheet_id == copied_sheet
            )
        })
        .expect("expected duplicated pivot on copied sheet");
    assert_eq!(duplicated.name, "Straße (3)");
}

#[test]
fn delete_sheet_removes_pivots_and_dependent_objects() {
    let mut wb = Workbook::new();
    let data = wb.add_sheet("Data").unwrap();
    let report = wb.add_sheet("Report").unwrap();

    let report_cache_id = uuid::Uuid::from_u128(10);
    wb.pivot_caches.push(PivotCacheModel {
        id: report_cache_id,
        source: PivotSource::Range {
            sheet_id: data,
            range: Range::from_a1("A1:A10").unwrap(),
        },
        needs_refresh: false,
    });
    let report_pivot_id = uuid::Uuid::from_u128(11);
    wb.pivot_tables.push(PivotTableModel {
        id: report_pivot_id,
        name: "PivotTable1".to_string(),
        source: PivotSource::Range {
            sheet_id: data,
            range: Range::from_a1("A1:A10").unwrap(),
        },
        destination: PivotDestination::Cell {
            sheet_id: report,
            cell: CellRef::from_a1("A1").unwrap(),
        },
        config: PivotConfig::default(),
        cache_id: Some(report_cache_id),
    });

    wb.pivot_charts.push(PivotChartModel {
        id: uuid::Uuid::from_u128(20),
        name: "Chart1".to_string(),
        pivot_table_id: report_pivot_id,
        chart_part: None,
        sheet_id: Some(report),
    });

    wb.slicers.push(SlicerModel {
        id: uuid::Uuid::from_u128(30),
        name: "SlicerOnReport".to_string(),
        connected_pivots: vec![report_pivot_id],
        sheet_id: report,
    });
    wb.slicers.push(SlicerModel {
        id: uuid::Uuid::from_u128(31),
        name: "SlicerOnData".to_string(),
        connected_pivots: vec![report_pivot_id],
        sheet_id: data,
    });

    // A second pivot/cache pair that should survive the deletion.
    let data_cache_id = uuid::Uuid::from_u128(12);
    wb.pivot_caches.push(PivotCacheModel {
        id: data_cache_id,
        source: PivotSource::Range {
            sheet_id: data,
            range: Range::from_a1("B1:B10").unwrap(),
        },
        needs_refresh: false,
    });
    wb.pivot_tables.push(PivotTableModel {
        id: uuid::Uuid::from_u128(13),
        name: "PivotTable2".to_string(),
        source: PivotSource::Range {
            sheet_id: data,
            range: Range::from_a1("B1:B10").unwrap(),
        },
        destination: PivotDestination::Cell {
            sheet_id: data,
            cell: CellRef::from_a1("D1").unwrap(),
        },
        config: PivotConfig::default(),
        cache_id: Some(data_cache_id),
    });

    wb.delete_sheet(report).unwrap();

    assert!(wb.pivot_tables.iter().all(|p| p.id != report_pivot_id));
    assert!(
        wb.pivot_charts.is_empty(),
        "expected pivot chart bound to removed pivot to be deleted"
    );
    assert!(
        wb.slicers.is_empty(),
        "expected slicers connected only to removed pivots to be deleted"
    );
    assert_eq!(wb.pivot_caches.len(), 1);
    assert_eq!(wb.pivot_caches[0].id, data_cache_id);
}
