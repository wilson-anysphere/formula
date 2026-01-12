use std::io::{Cursor, Read, Seek, SeekFrom, Write};
use std::sync::atomic::{AtomicUsize, Ordering};
use std::sync::Arc;

use zip::write::FileOptions;
use zip::{CompressionMethod, ZipArchive, ZipWriter};

/// A `Read + Seek` wrapper that counts the number of bytes the consumer reads.
struct CountingReader<R> {
    inner: R,
    bytes_read: Arc<AtomicUsize>,
}

impl<R> CountingReader<R> {
    fn new(inner: R) -> (Self, Arc<AtomicUsize>) {
        let bytes_read = Arc::new(AtomicUsize::new(0));
        (
            Self {
                inner,
                bytes_read: bytes_read.clone(),
            },
            bytes_read,
        )
    }
}

impl<R: Read> Read for CountingReader<R> {
    fn read(&mut self, buf: &mut [u8]) -> std::io::Result<usize> {
        let n = self.inner.read(buf)?;
        self.bytes_read.fetch_add(n, Ordering::Relaxed);
        Ok(n)
    }
}

impl<R: Seek> Seek for CountingReader<R> {
    fn seek(&mut self, pos: SeekFrom) -> std::io::Result<u64> {
        self.inner.seek(pos)
    }
}

#[test]
fn read_workbook_from_reader_does_not_read_unreferenced_large_parts() {
    // Generate a small XLSX package using the writer, then inject a large unreferenced binary part.
    // If `read_workbook_from_reader` eagerly reads every ZIP entry, this test will observe a large
    // byte count from `CountingReader`.
    let mut workbook = formula_model::Workbook::new();
    let _sheet_id = workbook.add_sheet("Sheet1").expect("add sheet");

    let mut buf = Cursor::new(Vec::new());
    formula_xlsx::write_workbook_to_writer(&workbook, &mut buf).expect("write workbook");
    let original_bytes = buf.into_inner();

    let cursor = Cursor::new(&original_bytes);
    let mut archive = ZipArchive::new(cursor).expect("parse original zip");

    let mut out = Cursor::new(Vec::new());
    let mut writer = ZipWriter::new(&mut out);
    let options = FileOptions::<()>::default().compression_method(CompressionMethod::Stored);

    for i in 0..archive.len() {
        let mut file = archive.by_index(i).expect("zip entry");
        if file.is_dir() {
            continue;
        }
        let name = file.name().to_string();
        let mut data = Vec::with_capacity(file.size() as usize);
        file.read_to_end(&mut data).expect("read zip entry");
        writer.start_file(name, options).expect("start zip entry");
        writer.write_all(&data).expect("write zip entry");
    }

    // Add an unreferenced, *stored* (uncompressed) part large enough that a full-package read would
    // be obvious in `bytes_read`.
    const HUGE_PART: &str = "xl/media/huge.bin";
    const HUGE_SIZE: usize = 5 * 1024 * 1024;
    writer
        .start_file(HUGE_PART, options)
        .expect("start huge part");

    let chunk = vec![0u8; 8192];
    let mut remaining = HUGE_SIZE;
    while remaining > 0 {
        let n = remaining.min(chunk.len());
        writer.write_all(&chunk[..n]).expect("write huge part chunk");
        remaining -= n;
    }

    writer.finish().expect("finalize zip");
    let bytes_with_huge_part = out.into_inner();

    let (reader, bytes_read) = CountingReader::new(Cursor::new(bytes_with_huge_part));
    let _loaded =
        formula_xlsx::read_workbook_from_reader(reader).expect("read workbook from reader");

    let read = bytes_read.load(Ordering::Relaxed);
    assert!(
        read < 1024 * 1024,
        "expected streaming reader to read <1MiB, but read {read} bytes"
    );
}

