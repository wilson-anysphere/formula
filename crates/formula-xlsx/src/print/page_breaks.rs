use super::{
    CellRange, ManualPageBreaks, Orientation, PageSetup, Scaling, DEFAULT_COL_WIDTH_POINTS,
    DEFAULT_ROW_HEIGHT_POINTS,
};

#[derive(Debug, Clone, Copy, PartialEq, Eq)]
pub struct Page {
    /// 1-based row number.
    pub start_row: u32,
    /// 1-based row number (inclusive).
    pub end_row: u32,
    /// 1-based column number.
    pub start_col: u32,
    /// 1-based column number (inclusive).
    pub end_col: u32,
}

pub fn calculate_pages(
    print_area: CellRange,
    col_widths_points: &[f64],
    row_heights_points: &[f64],
    page_setup: &PageSetup,
    manual_breaks: &ManualPageBreaks,
) -> Vec<Page> {
    let print_area = print_area.normalized();

    let (paper_w_in, paper_h_in) = page_setup.paper_size.dimensions_in_inches();
    let mut page_w = paper_w_in * 72.0;
    let mut page_h = paper_h_in * 72.0;
    if page_setup.orientation == Orientation::Landscape {
        std::mem::swap(&mut page_w, &mut page_h);
    }

    let margins = page_setup.margins;
    let printable_w = page_w - (margins.left + margins.right) * 72.0;
    let printable_h = page_h - (margins.top + margins.bottom) * 72.0;

    let scale_factor = match page_setup.scaling {
        Scaling::Percent(pct) => (pct as f64) / 100.0,
        Scaling::FitTo { width, height } => {
            let content_w =
                sum_slice_range(
                    col_widths_points,
                    print_area.start_col,
                    print_area.end_col,
                    DEFAULT_COL_WIDTH_POINTS,
                );
            let content_h =
                sum_slice_range(
                    row_heights_points,
                    print_area.start_row,
                    print_area.end_row,
                    DEFAULT_ROW_HEIGHT_POINTS,
                );

            let scale_w = if width > 0 && content_w > 0.0 {
                Some((width as f64 * printable_w) / content_w)
            } else {
                None
            };
            let scale_h = if height > 0 && content_h > 0.0 {
                Some((height as f64 * printable_h) / content_h)
            } else {
                None
            };

            match (scale_w, scale_h) {
                (Some(w), Some(h)) => w.min(h),
                (Some(w), None) => w,
                (None, Some(h)) => h,
                (None, None) => 1.0,
            }
        }
    };

    let effective_w = if scale_factor > 0.0 {
        printable_w / scale_factor
    } else {
        printable_w
    };
    let effective_h = if scale_factor > 0.0 {
        printable_h / scale_factor
    } else {
        printable_h
    };

    let col_starts = compute_break_starts(
        print_area.start_col,
        print_area.end_col,
        col_widths_points,
        effective_w,
        manual_breaks
            .col_breaks_after
            .iter()
            .map(|break_after| break_after.saturating_add(1)),
        DEFAULT_COL_WIDTH_POINTS,
    );

    let row_starts = compute_break_starts(
        print_area.start_row,
        print_area.end_row,
        row_heights_points,
        effective_h,
        manual_breaks
            .row_breaks_after
            .iter()
            .map(|break_after| break_after.saturating_add(1)),
        DEFAULT_ROW_HEIGHT_POINTS,
    );

    let col_segments = starts_to_segments(&col_starts, print_area.end_col);
    let row_segments = starts_to_segments(&row_starts, print_area.end_row);

    let mut pages = Vec::new();
    let est = col_segments.len().saturating_mul(row_segments.len());
    let _ = pages.try_reserve(est);
    for row in row_segments {
        for col in &col_segments {
            pages.push(Page {
                start_row: row.0,
                end_row: row.1,
                start_col: col.0,
                end_col: col.1,
            });
        }
    }

    pages
}

fn starts_to_segments(starts: &[u32], end_inclusive: u32) -> Vec<(u32, u32)> {
    let mut segments = Vec::new();
    for (idx, start) in starts.iter().enumerate() {
        let end = starts
            .get(idx + 1)
            .map(|next| next.saturating_sub(1))
            .unwrap_or(end_inclusive);
        if *start <= end {
            segments.push((*start, end));
        }
    }
    segments
}

fn compute_break_starts<I: IntoIterator<Item = u32>>(
    start: u32,
    end: u32,
    sizes: &[f64],
    page_capacity: f64,
    manual_starts: I,
    default_size: f64,
) -> Vec<u32> {
    let mut starts = Vec::new();
    starts.push(start);

    let mut current = start;
    while current <= end {
        let mut acc = 0.0;
        let mut next = current;

        while next <= end {
            let size = sizes
                .get((next - 1) as usize)
                .copied()
                .unwrap_or(default_size);

            if next > current && acc + size > page_capacity {
                break;
            }

            acc += size;
            next += 1;
        }

        if next == current {
            next += 1;
        }

        current = next;
        if current <= end {
            starts.push(current);
        }
    }

    for manual_start in manual_starts {
        if manual_start > start && manual_start <= end {
            starts.push(manual_start);
        }
    }

    starts.sort_unstable();
    starts.dedup();
    starts
}

fn sum_slice_range(sizes: &[f64], start: u32, end: u32, default_size: f64) -> f64 {
    let mut sum = 0.0;
    for idx in start..=end {
        sum += sizes
            .get((idx - 1) as usize)
            .copied()
            .unwrap_or(default_size);
    }
    sum
}
