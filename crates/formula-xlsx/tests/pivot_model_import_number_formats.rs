use formula_model::Workbook;
use formula_xlsx::pivots::model_bridge::pivot_table_to_model_value_fields;
use formula_xlsx::{StylesPart, XlsxPackage};

use pretty_assertions::assert_eq;

const FIXTURE: &[u8] = include_bytes!("fixtures/pivot-datafield-numfmtid.xlsx");

#[test]
fn imports_pivot_value_field_num_fmt_id_into_model_number_format() {
    let pkg = XlsxPackage::from_bytes(FIXTURE).expect("read package");

    let table = pkg
        .pivot_table_definition("xl/pivotTables/pivotTable1.xml")
        .expect("parse pivot table definition");

    let cache_def = pkg
        .pivot_cache_definition("xl/pivotCache/pivotCacheDefinition1.xml")
        .expect("parse cache def")
        .expect("cache definition exists");

    // Match the workbook reader pipeline: parse `styles.xml` once into a
    // `StylesPart` and plumb its `numFmtId -> formatCode` mapping into pivot
    // import logic (no re-parsing).
    let mut workbook = Workbook::new();
    let styles_part = StylesPart::parse_or_default(pkg.part("xl/styles.xml"), &mut workbook.styles)
        .expect("parse styles");

    let value_fields = pivot_table_to_model_value_fields(&table, &cache_def, &styles_part);
    assert_eq!(value_fields.len(), 3);

    assert_eq!(value_fields[0].name, "Sum of Sales");
    assert_eq!(value_fields[0].number_format.as_deref(), Some("0.000"));

    assert_eq!(value_fields[1].name, "Average of Sales");
    assert_eq!(value_fields[1].number_format.as_deref(), Some("#,##0.00"));

    assert_eq!(value_fields[2].name, "Count of Sales");
    assert_eq!(
        value_fields[2].number_format.as_deref(),
        Some("__builtin_numFmtId:63")
    );
}

