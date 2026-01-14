use std::io::{Cursor, Write};

use formula_xlsx::{read_part_from_reader_limited, XlsxError};

#[test]
fn read_part_from_reader_limited_rejects_parts_larger_than_max() {
    let cursor = Cursor::new(Vec::new());
    let mut zip = zip::ZipWriter::new(cursor);
    zip.start_file(
        "xl/oversize.bin",
        zip::write::FileOptions::<()>::default().compression_method(zip::CompressionMethod::Stored),
    )
    .expect("start file");
    zip.write_all(b"0123456789abcdef").expect("write payload");
    let bytes = zip.finish().expect("finish zip").into_inner();

    let err = read_part_from_reader_limited(Cursor::new(bytes), "xl/oversize.bin", 8)
        .expect_err("expected oversize part to fail");

    match err {
        XlsxError::PartTooLarge { part, size, max } => {
            assert_eq!(part, "xl/oversize.bin");
            assert_eq!(size, 16);
            assert_eq!(max, 8);
        }
        other => panic!("expected PartTooLarge, got {other:?}"),
    }
}
