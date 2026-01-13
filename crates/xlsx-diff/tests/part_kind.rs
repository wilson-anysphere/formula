use xlsx_diff::{classify_part, PartKind};

#[test]
fn classify_part_common_parts() {
    assert_eq!(classify_part("[Content_Types].xml"), PartKind::ContentTypes);
    assert_eq!(
        classify_part("/[Content_Types].xml"),
        PartKind::ContentTypes
    );

    assert_eq!(classify_part("_rels/.rels"), PartKind::Rels);
    assert_eq!(classify_part("xl/_rels/workbook.xml.rels"), PartKind::Rels);
    assert_eq!(
        classify_part(r"xl\worksheets\_rels\sheet1.xml.rels"),
        PartKind::Rels
    );

    assert_eq!(classify_part("docProps/app.xml"), PartKind::DocProps);
    assert_eq!(classify_part("docProps/core.xml"), PartKind::DocProps);

    assert_eq!(classify_part("xl/workbook.xml"), PartKind::Workbook);
    assert_eq!(classify_part("xl/workbook.bin"), PartKind::Workbook);

    assert_eq!(
        classify_part("xl/worksheets/sheet1.xml"),
        PartKind::Worksheet
    );
    assert_eq!(
        classify_part("xl/worksheets/sheet1.bin"),
        PartKind::Worksheet
    );
    assert_eq!(
        classify_part(r"\xl\worksheets\..\worksheets\sheet1.xml"),
        PartKind::Worksheet
    );

    assert_eq!(classify_part("xl/styles.xml"), PartKind::Styles);
    assert_eq!(classify_part("xl/tableStyles.xml"), PartKind::Styles);

    assert_eq!(
        classify_part("xl/sharedStrings.xml"),
        PartKind::SharedStrings
    );

    assert_eq!(classify_part("xl/theme/theme1.xml"), PartKind::Theme);

    assert_eq!(classify_part("xl/calcChain.xml"), PartKind::CalcChain);
    assert_eq!(classify_part("xl/calcChain.bin"), PartKind::CalcChain);

    assert_eq!(classify_part("xl/media/image1.png"), PartKind::Media);
    assert_eq!(
        classify_part("xl/drawings/drawing1.xml"),
        PartKind::Drawings
    );

    assert_eq!(classify_part("xl/charts/chart1.xml"), PartKind::Charts);
    assert_eq!(classify_part("xl/chartsheets/sheet1.xml"), PartKind::Charts);

    assert_eq!(classify_part("xl/tables/table1.xml"), PartKind::Tables);

    assert_eq!(
        classify_part("xl/pivotTables/pivotTable1.xml"),
        PartKind::Pivot
    );
    assert_eq!(
        classify_part("xl/pivotCache/pivotCacheDefinition1.xml"),
        PartKind::Pivot
    );

    assert_eq!(classify_part("xl/vbaProject.bin"), PartKind::Vba);
    assert_eq!(classify_part("xl/vbaProjectSignature.bin"), PartKind::Vba);

    assert_eq!(classify_part("customXml/item1.xml"), PartKind::Other);
}
