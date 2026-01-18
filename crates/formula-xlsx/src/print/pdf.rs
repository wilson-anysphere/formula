use super::{
    calculate_pages, CellRange, ManualPageBreaks, Orientation, PageSetup, PrintError, Scaling,
    DEFAULT_COL_WIDTH_POINTS, DEFAULT_ROW_HEIGHT_POINTS,
};

/// Generate a basic PDF for a sheet range using the same pagination rules as `calculate_pages`.
///
/// MVP scope:
/// - Text-only rendering (no cell borders, no rich formatting).
/// - Built-in Helvetica font (Type1).
/// - Pagination respects page size/orientation/margins/scaling/manual breaks.
pub fn export_range_to_pdf_bytes<F>(
    sheet_name: &str,
    print_area: CellRange,
    col_widths_points: &[f64],
    row_heights_points: &[f64],
    page_setup: &PageSetup,
    manual_breaks: &ManualPageBreaks,
    mut cell_text: F,
) -> Result<Vec<u8>, PrintError>
where
    F: FnMut(u32, u32) -> Option<String>,
{
    let pages = calculate_pages(
        print_area,
        col_widths_points,
        row_heights_points,
        page_setup,
        manual_breaks,
    );

    let scale_factor = calculate_scale_factor(
        print_area,
        col_widths_points,
        row_heights_points,
        page_setup,
    );

    let (paper_w_in, paper_h_in) = page_setup.paper_size.dimensions_in_inches();
    let mut page_w = paper_w_in * 72.0;
    let mut page_h = paper_h_in * 72.0;
    if page_setup.orientation == Orientation::Landscape {
        std::mem::swap(&mut page_w, &mut page_h);
    }

    let margins = page_setup.margins;
    let left = margins.left * 72.0;
    let top = margins.top * 72.0;

    let font_size = 10.0;
    let padding = 2.0;

    let mut page_streams = Vec::new();
    if page_streams.try_reserve_exact(pages.len()).is_err() {
        return Err(PrintError::AllocationFailure("export_range_to_pdf_bytes page_streams"));
    }
    for (page_idx, page) in pages.iter().enumerate() {
        let mut stream = String::new();
        stream.push_str("BT\n");
        stream.push_str(&format!("/F1 {font_size} Tf\n"));

        // Header: sheet name + page number.
        let header_y = page_h - top - font_size;
        stream.push_str(&format!(
            "1 0 0 1 {x:.2} {y:.2} Tm ({text}) Tj\n",
            x = left,
            y = header_y,
            text = escape_pdf_string(&sanitize_pdf_text(&format!(
                "{sheet_name}  (Page {page_no})",
                page_no = page_idx + 1
            )))
        ));

        // Cells.
        let mut y_cursor = page_h - top - (font_size * 2.0); // leave room under header
        for row in page.start_row..=page.end_row {
            let row_height = row_heights_points
                .get((row - 1) as usize)
                .copied()
                .unwrap_or(DEFAULT_ROW_HEIGHT_POINTS)
                * scale_factor;

            let mut x_cursor = left;
            for col in page.start_col..=page.end_col {
                let col_width = col_widths_points
                    .get((col - 1) as usize)
                    .copied()
                    .unwrap_or(DEFAULT_COL_WIDTH_POINTS)
                    * scale_factor;

                if let Some(text) = cell_text(row, col) {
                    let text = sanitize_pdf_text(&text);
                    if !text.is_empty() {
                        let text = escape_pdf_string(&text);
                        let x = x_cursor + padding;
                        let y = y_cursor - padding;
                        stream.push_str(&format!("1 0 0 1 {x:.2} {y:.2} Tm ({text}) Tj\n"));
                    }
                }

                x_cursor += col_width;
            }

            y_cursor -= row_height.max(font_size + padding);
        }

        stream.push_str("ET\n");
        page_streams.push(stream.into_bytes());
    }

    Ok(build_pdf(page_w, page_h, &page_streams))
}

fn calculate_scale_factor(
    print_area: CellRange,
    col_widths_points: &[f64],
    row_heights_points: &[f64],
    page_setup: &PageSetup,
) -> f64 {
    let (paper_w_in, paper_h_in) = page_setup.paper_size.dimensions_in_inches();
    let mut page_w = paper_w_in * 72.0;
    let mut page_h = paper_h_in * 72.0;
    if page_setup.orientation == Orientation::Landscape {
        std::mem::swap(&mut page_w, &mut page_h);
    }

    let margins = page_setup.margins;
    let printable_w = page_w - (margins.left + margins.right) * 72.0;
    let printable_h = page_h - (margins.top + margins.bottom) * 72.0;

    match page_setup.scaling {
        Scaling::Percent(pct) => (pct as f64) / 100.0,
        Scaling::FitTo { width, height } => {
            let print_area = print_area.normalized();
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
    }
}

fn sanitize_pdf_text(text: &str) -> String {
    text.chars()
        .map(|c| if c.is_ascii() { c } else { '?' })
        .collect()
}

fn escape_pdf_string(text: &str) -> String {
    let mut out = String::new();
    let _ = out.try_reserve(text.len());
    for ch in text.chars() {
        match ch {
            '\\' => out.push_str("\\\\"),
            '(' => out.push_str("\\("),
            ')' => out.push_str("\\)"),
            '\n' | '\r' => out.push(' '),
            _ => out.push(ch),
        }
    }
    out
}

fn build_pdf(page_w: f64, page_h: f64, page_streams: &[Vec<u8>]) -> Vec<u8> {
    // Object numbers:
    // 1 = catalog
    // 2 = pages
    // 3..(3+N-1) = page objects
    // (3+N)..(3+2N-1) = content stream objects
    // last = font object
    let page_count = page_streams.len();
    let pages_obj = 2u32;
    let first_page_obj = 3u32;
    let first_content_obj = first_page_obj + page_count as u32;
    let font_obj = first_content_obj + page_count as u32;
    let total_objs = font_obj;

    let mut bytes = Vec::new();
    bytes.extend_from_slice(b"%PDF-1.4\n");

    let mut offsets = Vec::new();
    let _ = offsets.try_reserve_exact((total_objs.saturating_add(1)) as usize);
    offsets.push(0u64); // xref entry 0

    let write_obj = |obj_no: u32, content: &[u8], bytes: &mut Vec<u8>, offsets: &mut Vec<u64>| {
        offsets.push(bytes.len() as u64);
        bytes.extend_from_slice(format!("{obj_no} 0 obj\n").as_bytes());
        bytes.extend_from_slice(content);
        bytes.extend_from_slice(b"\nendobj\n");
    };

    // 1: catalog
    write_obj(
        1,
        format!("<< /Type /Catalog /Pages {pages_obj} 0 R >>").as_bytes(),
        &mut bytes,
        &mut offsets,
    );

    // 2: pages
    let kids = (0..page_count)
        .map(|i| format!("{} 0 R", first_page_obj + i as u32))
        .collect::<Vec<_>>()
        .join(" ");
    write_obj(
        pages_obj,
        format!("<< /Type /Pages /Kids [ {kids} ] /Count {page_count} >>").as_bytes(),
        &mut bytes,
        &mut offsets,
    );

    // Page objects + content streams
    for i in 0..page_count {
        let page_obj_no = first_page_obj + i as u32;
        let content_obj_no = first_content_obj + i as u32;

        write_obj(
            page_obj_no,
            format!(
                "<< /Type /Page /Parent {pages_obj} 0 R /MediaBox [0 0 {page_w:.2} {page_h:.2}] /Resources << /Font << /F1 {font_obj} 0 R >> >> /Contents {content_obj_no} 0 R >>"
            )
            .as_bytes(),
            &mut bytes,
            &mut offsets,
        );

        let stream = &page_streams[i];
        let stream_obj = build_stream_obj(stream);
        write_obj(content_obj_no, &stream_obj, &mut bytes, &mut offsets);
    }

    // Font object
    write_obj(
        font_obj,
        b"<< /Type /Font /Subtype /Type1 /BaseFont /Helvetica >>",
        &mut bytes,
        &mut offsets,
    );

    // xref
    let xref_offset = bytes.len() as u64;
    bytes.extend_from_slice(format!("xref\n0 {}\n", total_objs + 1).as_bytes());
    bytes.extend_from_slice(b"0000000000 65535 f \n");

    for offset in offsets.iter().skip(1) {
        bytes.extend_from_slice(format!("{offset:010} 00000 n \n").as_bytes());
    }

    // trailer
    bytes.extend_from_slice(
        format!(
            "trailer\n<< /Size {size} /Root 1 0 R >>\nstartxref\n{xref_offset}\n%%EOF\n",
            size = total_objs + 1
        )
        .as_bytes(),
    );

    bytes
}

fn build_stream_obj(stream: &[u8]) -> Vec<u8> {
    let mut out = Vec::new();
    out.extend_from_slice(format!("<< /Length {} >>\nstream\n", stream.len()).as_bytes());
    out.extend_from_slice(stream);
    if !stream.ends_with(b"\n") {
        out.push(b'\n');
    }
    out.extend_from_slice(b"endstream");
    out
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
