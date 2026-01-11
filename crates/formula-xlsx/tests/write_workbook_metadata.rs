use formula_model::{Hyperlink, HyperlinkTarget, Range, SheetVisibility, TabColor, Workbook};
use tempfile::tempdir;

#[test]
fn write_workbook_emits_basic_worksheet_metadata() -> Result<(), Box<dyn std::error::Error>> {
    let mut workbook = Workbook::new();
    let sheet1_id = workbook.add_sheet("Sheet1")?;
    let hidden_id = workbook.add_sheet("Hidden")?;

    {
        let sheet1 = workbook.sheet_mut(sheet1_id).unwrap();
        sheet1.tab_color = Some(TabColor::rgb("FF00FF00"));
        sheet1
            .merge_range(Range::from_a1("A1:B2")?)
            .expect("merge ok");
        sheet1.hyperlinks.push(Hyperlink {
            range: Range::from_a1("A1")?,
            target: HyperlinkTarget::ExternalUrl {
                uri: "https://example.com".to_string(),
            },
            display: None,
            tooltip: None,
            // The simple `write_workbook` exporter should allocate a relationship id.
            rel_id: None,
        });
    }

    {
        let hidden = workbook.sheet_mut(hidden_id).unwrap();
        hidden.visibility = SheetVisibility::Hidden;
    }

    let dir = tempdir()?;
    let out_path = dir.path().join("metadata.xlsx");
    formula_xlsx::write_workbook(&workbook, &out_path)?;

    let loaded = formula_xlsx::read_workbook(&out_path)?;
    let sheet1 = loaded.sheet_by_name("Sheet1").unwrap();
    assert_eq!(sheet1.tab_color, Some(TabColor::rgb("FF00FF00")));
    assert_eq!(sheet1.merged_regions.region_count(), 1);
    assert_eq!(sheet1.hyperlinks.len(), 1);
    assert_eq!(
        sheet1.hyperlinks[0].target,
        HyperlinkTarget::ExternalUrl {
            uri: "https://example.com".to_string()
        }
    );
    assert!(sheet1.hyperlinks[0].rel_id.is_some());

    let hidden = loaded.sheet_by_name("Hidden").unwrap();
    assert_eq!(hidden.visibility, SheetVisibility::Hidden);

    Ok(())
}

