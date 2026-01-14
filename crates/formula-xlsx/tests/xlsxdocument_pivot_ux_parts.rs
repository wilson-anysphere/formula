use formula_xlsx::{load_from_bytes, XlsxPackage};

#[test]
fn xlsxdocument_pivot_slicer_parts_match_package() -> Result<(), Box<dyn std::error::Error>> {
    let fixture = include_bytes!("fixtures/pivot_slicers_and_chart.xlsx");
    let package = XlsxPackage::from_bytes(fixture)?;
    let doc = load_from_bytes(fixture)?;

    assert_eq!(doc.pivot_slicer_parts()?, package.pivot_slicer_parts()?);
    Ok(())
}

#[test]
fn xlsxdocument_pivot_chart_parts_match_package() -> Result<(), Box<dyn std::error::Error>> {
    let fixture = include_bytes!("fixtures/pivot_slicers_and_chart.xlsx");
    let package = XlsxPackage::from_bytes(fixture)?;
    let doc = load_from_bytes(fixture)?;

    assert_eq!(doc.pivot_chart_parts()?, package.pivot_chart_parts()?);
    Ok(())
}

#[test]
fn xlsxdocument_pivot_chart_parts_with_placement_match_package(
) -> Result<(), Box<dyn std::error::Error>> {
    let fixture = include_bytes!("fixtures/pivot_slicers_and_chart.xlsx");
    let package = XlsxPackage::from_bytes(fixture)?;
    let doc = load_from_bytes(fixture)?;

    assert_eq!(
        doc.pivot_chart_parts_with_placement()?,
        package.pivot_chart_parts_with_placement()?
    );
    Ok(())
}

#[test]
fn xlsxdocument_pivot_graph_matches_package() -> Result<(), Box<dyn std::error::Error>> {
    let fixture = include_bytes!("fixtures/pivot-full.xlsx");
    let package = XlsxPackage::from_bytes(fixture)?;
    let doc = load_from_bytes(fixture)?;

    assert_eq!(doc.pivot_graph()?, package.pivot_graph()?);
    Ok(())
}

#[test]
fn xlsxdocument_pivot_ux_graph_matches_package() -> Result<(), Box<dyn std::error::Error>> {
    let fixture = include_bytes!("fixtures/pivot_slicers_and_chart.xlsx");
    let package = XlsxPackage::from_bytes(fixture)?;
    let doc = load_from_bytes(fixture)?;

    assert_eq!(doc.pivot_ux_graph()?, package.pivot_ux_graph()?);
    Ok(())
}
