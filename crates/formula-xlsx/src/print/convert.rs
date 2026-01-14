use formula_model::{CellRef, Range as ModelRange};

use super::{
    CellRange, ColRange, ManualPageBreaks, Orientation, PageMargins, PageSetup, PaperSize,
    PrintTitles, RowRange, Scaling, SheetPrintSettings, WorkbookPrintSettings,
};

fn one_based_to_zero_based(v: u32) -> u32 {
    v.saturating_sub(1)
}

fn zero_based_to_one_based(v: u32) -> u32 {
    v.saturating_add(1)
}

impl CellRange {
    /// Convert this 1-based XLSX range into the 0-based model range.
    pub fn to_model(&self) -> ModelRange {
        let r = self.normalized();
        ModelRange::new(
            CellRef::new(
                one_based_to_zero_based(r.start_row),
                one_based_to_zero_based(r.start_col),
            ),
            CellRef::new(
                one_based_to_zero_based(r.end_row),
                one_based_to_zero_based(r.end_col),
            ),
        )
    }
}

impl From<ModelRange> for CellRange {
    fn from(range: ModelRange) -> Self {
        Self::from(&range)
    }
}

impl From<&ModelRange> for CellRange {
    fn from(range: &ModelRange) -> Self {
        Self {
            start_row: zero_based_to_one_based(range.start.row),
            end_row: zero_based_to_one_based(range.end.row),
            start_col: zero_based_to_one_based(range.start.col),
            end_col: zero_based_to_one_based(range.end.col),
        }
    }
}

impl RowRange {
    pub fn to_model(&self) -> formula_model::RowRange {
        let r = self.normalized();
        formula_model::RowRange {
            start: one_based_to_zero_based(r.start),
            end: one_based_to_zero_based(r.end),
        }
    }
}

impl From<formula_model::RowRange> for RowRange {
    fn from(range: formula_model::RowRange) -> Self {
        Self::from(&range)
    }
}

impl From<&formula_model::RowRange> for RowRange {
    fn from(range: &formula_model::RowRange) -> Self {
        Self {
            start: zero_based_to_one_based(range.start),
            end: zero_based_to_one_based(range.end),
        }
    }
}

impl ColRange {
    pub fn to_model(&self) -> formula_model::ColRange {
        let r = self.normalized();
        formula_model::ColRange {
            start: one_based_to_zero_based(r.start),
            end: one_based_to_zero_based(r.end),
        }
    }
}

impl From<formula_model::ColRange> for ColRange {
    fn from(range: formula_model::ColRange) -> Self {
        Self::from(&range)
    }
}

impl From<&formula_model::ColRange> for ColRange {
    fn from(range: &formula_model::ColRange) -> Self {
        Self {
            start: zero_based_to_one_based(range.start),
            end: zero_based_to_one_based(range.end),
        }
    }
}

impl PrintTitles {
    pub fn to_model(&self) -> formula_model::PrintTitles {
        formula_model::PrintTitles {
            repeat_rows: self.repeat_rows.as_ref().map(|r| r.to_model()),
            repeat_cols: self.repeat_cols.as_ref().map(|r| r.to_model()),
        }
    }
}

impl From<formula_model::PrintTitles> for PrintTitles {
    fn from(titles: formula_model::PrintTitles) -> Self {
        Self::from(&titles)
    }
}

impl From<&formula_model::PrintTitles> for PrintTitles {
    fn from(titles: &formula_model::PrintTitles) -> Self {
        Self {
            repeat_rows: titles.repeat_rows.map(RowRange::from),
            repeat_cols: titles.repeat_cols.map(ColRange::from),
        }
    }
}

impl Orientation {
    pub fn to_model(&self) -> formula_model::Orientation {
        match self {
            Orientation::Portrait => formula_model::Orientation::Portrait,
            Orientation::Landscape => formula_model::Orientation::Landscape,
        }
    }
}

impl From<formula_model::Orientation> for Orientation {
    fn from(orientation: formula_model::Orientation) -> Self {
        match orientation {
            formula_model::Orientation::Portrait => Orientation::Portrait,
            formula_model::Orientation::Landscape => Orientation::Landscape,
        }
    }
}

impl PaperSize {
    pub fn to_model(&self) -> formula_model::PaperSize {
        formula_model::PaperSize { code: self.code }
    }
}

impl From<formula_model::PaperSize> for PaperSize {
    fn from(size: formula_model::PaperSize) -> Self {
        PaperSize { code: size.code }
    }
}

impl PageMargins {
    pub fn to_model(&self) -> formula_model::PageMargins {
        formula_model::PageMargins {
            left: self.left,
            right: self.right,
            top: self.top,
            bottom: self.bottom,
            header: self.header,
            footer: self.footer,
        }
    }
}

impl From<formula_model::PageMargins> for PageMargins {
    fn from(m: formula_model::PageMargins) -> Self {
        Self {
            left: m.left,
            right: m.right,
            top: m.top,
            bottom: m.bottom,
            header: m.header,
            footer: m.footer,
        }
    }
}

impl Scaling {
    pub fn to_model(&self) -> formula_model::Scaling {
        match *self {
            Scaling::Percent(pct) => formula_model::Scaling::Percent(pct),
            Scaling::FitTo { width, height } => formula_model::Scaling::FitTo { width, height },
        }
    }
}

impl From<formula_model::Scaling> for Scaling {
    fn from(s: formula_model::Scaling) -> Self {
        match s {
            formula_model::Scaling::Percent(pct) => Scaling::Percent(pct),
            formula_model::Scaling::FitTo { width, height } => Scaling::FitTo { width, height },
        }
    }
}

impl PageSetup {
    pub fn to_model(&self) -> formula_model::PageSetup {
        formula_model::PageSetup {
            orientation: self.orientation.to_model(),
            paper_size: self.paper_size.to_model(),
            margins: self.margins.to_model(),
            scaling: self.scaling.to_model(),
        }
    }
}

impl From<formula_model::PageSetup> for PageSetup {
    fn from(setup: formula_model::PageSetup) -> Self {
        Self::from(&setup)
    }
}

impl From<&formula_model::PageSetup> for PageSetup {
    fn from(setup: &formula_model::PageSetup) -> Self {
        Self {
            orientation: Orientation::from(setup.orientation),
            paper_size: PaperSize::from(setup.paper_size),
            margins: PageMargins::from(setup.margins),
            scaling: Scaling::from(setup.scaling),
        }
    }
}

impl ManualPageBreaks {
    pub fn to_model(&self) -> formula_model::ManualPageBreaks {
        formula_model::ManualPageBreaks {
            row_breaks_after: self
                .row_breaks_after
                .iter()
                .copied()
                .map(one_based_to_zero_based)
                .collect(),
            col_breaks_after: self
                .col_breaks_after
                .iter()
                .copied()
                .map(one_based_to_zero_based)
                .collect(),
        }
    }
}

impl From<formula_model::ManualPageBreaks> for ManualPageBreaks {
    fn from(breaks: formula_model::ManualPageBreaks) -> Self {
        Self::from(&breaks)
    }
}

impl From<&formula_model::ManualPageBreaks> for ManualPageBreaks {
    fn from(breaks: &formula_model::ManualPageBreaks) -> Self {
        Self {
            row_breaks_after: breaks
                .row_breaks_after
                .iter()
                .copied()
                .map(zero_based_to_one_based)
                .collect(),
            col_breaks_after: breaks
                .col_breaks_after
                .iter()
                .copied()
                .map(zero_based_to_one_based)
                .collect(),
        }
    }
}

impl SheetPrintSettings {
    pub fn to_model(&self) -> formula_model::SheetPrintSettings {
        formula_model::SheetPrintSettings {
            sheet_name: self.sheet_name.clone(),
            print_area: self.print_area.as_ref().map(|areas| {
                areas.iter().map(|r| r.to_model()).collect()
            }),
            print_titles: self.print_titles.as_ref().map(|t| t.to_model()),
            page_setup: self.page_setup.to_model(),
            manual_page_breaks: self.manual_page_breaks.to_model(),
        }
    }
}

impl From<formula_model::SheetPrintSettings> for SheetPrintSettings {
    fn from(settings: formula_model::SheetPrintSettings) -> Self {
        Self::from(&settings)
    }
}

impl From<&formula_model::SheetPrintSettings> for SheetPrintSettings {
    fn from(settings: &formula_model::SheetPrintSettings) -> Self {
        Self {
            sheet_name: settings.sheet_name.clone(),
            print_area: settings
                .print_area
                .as_ref()
                .map(|areas| areas.iter().copied().map(CellRange::from).collect()),
            print_titles: settings.print_titles.map(PrintTitles::from),
            page_setup: PageSetup::from(&settings.page_setup),
            manual_page_breaks: ManualPageBreaks::from(&settings.manual_page_breaks),
        }
    }
}

impl WorkbookPrintSettings {
    pub fn to_model(&self) -> formula_model::WorkbookPrintSettings {
        formula_model::WorkbookPrintSettings {
            sheets: self.sheets.iter().map(|s| s.to_model()).collect(),
        }
    }
}

impl From<formula_model::WorkbookPrintSettings> for WorkbookPrintSettings {
    fn from(settings: formula_model::WorkbookPrintSettings) -> Self {
        Self::from(&settings)
    }
}

impl From<&formula_model::WorkbookPrintSettings> for WorkbookPrintSettings {
    fn from(settings: &formula_model::WorkbookPrintSettings) -> Self {
        Self {
            sheets: settings.sheets.iter().map(SheetPrintSettings::from).collect(),
        }
    }
}

#[cfg(test)]
mod tests {
    use super::*;
    use std::collections::BTreeSet;

    #[test]
    fn print_area_conversion_is_off_by_one_correct() {
        let xlsx = CellRange {
            start_row: 1,
            end_row: 2,
            start_col: 1,
            end_col: 3,
        };

        let model = xlsx.to_model();
        assert_eq!(
            model,
            ModelRange::new(CellRef::new(0, 0), CellRef::new(1, 2))
        );

        let roundtrip = CellRange::from(model);
        assert_eq!(roundtrip, xlsx);
    }

    #[test]
    fn print_titles_conversion_is_off_by_one_correct() {
        let xlsx = PrintTitles {
            repeat_rows: Some(RowRange { start: 1, end: 3 }),
            repeat_cols: Some(ColRange { start: 2, end: 4 }),
        };

        let model = xlsx.to_model();
        assert_eq!(
            model,
            formula_model::PrintTitles {
                repeat_rows: Some(formula_model::RowRange { start: 0, end: 2 }),
                repeat_cols: Some(formula_model::ColRange { start: 1, end: 3 }),
            }
        );

        let roundtrip = PrintTitles::from(model);
        assert_eq!(roundtrip, xlsx);
    }

    #[test]
    fn manual_page_breaks_conversion_is_off_by_one_correct() {
        let xlsx = ManualPageBreaks {
            row_breaks_after: BTreeSet::from([1, 5]),
            col_breaks_after: BTreeSet::from([2]),
        };

        let model = xlsx.to_model();
        assert_eq!(
            model,
            formula_model::ManualPageBreaks {
                row_breaks_after: BTreeSet::from([0, 4]),
                col_breaks_after: BTreeSet::from([1]),
            }
        );

        let roundtrip = ManualPageBreaks::from(model);
        assert_eq!(roundtrip, xlsx);
    }
}
