use std::fs::File;
use std::io::{Read, Write};
use std::path::Path;

use thiserror::Error;
use zip::write::FileOptions;
use zip::{CompressionMethod, ZipArchive, ZipWriter};

use crate::zip_util::open_zip_part;

use super::{parse_shared_strings_xml, write_shared_strings_xml, SharedStrings};

#[derive(Debug, Error)]
pub enum SharedStringsXlsxError {
    #[error("zip error: {0}")]
    Zip(#[from] zip::result::ZipError),
    #[error("io error: {0}")]
    Io(#[from] std::io::Error),
    #[error("sharedStrings.xml parse error: {0}")]
    Parse(#[from] super::parse::SharedStringsError),
    #[error("sharedStrings.xml write error: {0}")]
    Write(#[from] super::write::WriteSharedStringsError),
    #[error("workbook is missing xl/sharedStrings.xml")]
    MissingSharedStrings,
}

/// Read and parse `xl/sharedStrings.xml` from an `.xlsx` zip archive.
pub fn read_shared_strings_from_xlsx(
    path: impl AsRef<Path>,
) -> Result<SharedStrings, SharedStringsXlsxError> {
    let file = File::open(path)?;
    let mut zip = ZipArchive::new(file)?;
    let mut ss_file = open_zip_part(&mut zip, "xl/sharedStrings.xml").map_err(|e| match e {
        zip::result::ZipError::FileNotFound => SharedStringsXlsxError::MissingSharedStrings,
        other => SharedStringsXlsxError::Zip(other),
    })?;

    let mut xml = String::new();
    ss_file.read_to_string(&mut xml)?;
    Ok(parse_shared_strings_xml(&xml)?)
}

/// Write a new `.xlsx` file, copying all entries from `input_path` and replacing
/// (or adding) `xl/sharedStrings.xml`.
///
/// This is a small utility for testing and round-trip preservation. It does not
/// attempt to preserve all zip metadata.
pub fn write_shared_strings_to_xlsx(
    input_path: impl AsRef<Path>,
    output_path: impl AsRef<Path>,
    shared_strings: &SharedStrings,
) -> Result<(), SharedStringsXlsxError> {
    let input_file = File::open(input_path)?;
    let mut input_zip = ZipArchive::new(input_file)?;

    let output_file = File::create(output_path)?;
    let mut output_zip = ZipWriter::new(output_file);

    let shared_strings_xml = write_shared_strings_xml(shared_strings)?;

    let mut seen_shared_strings = false;
    for i in 0..input_zip.len() {
        let mut file = input_zip.by_index(i)?;
        let name = file.name().strip_prefix('/').unwrap_or(file.name()).to_string();

        if name == "xl/sharedStrings.xml" {
            seen_shared_strings = true;
            let options =
                FileOptions::<()>::default().compression_method(CompressionMethod::Deflated);
            output_zip.start_file(name, options)?;
            output_zip.write_all(shared_strings_xml.as_bytes())?;
            continue;
        }

        let mut buf = Vec::new();
        file.read_to_end(&mut buf)?;

        let mut options = FileOptions::<()>::default().compression_method(match file.compression() {
            CompressionMethod::Stored => CompressionMethod::Stored,
            _ => CompressionMethod::Deflated,
        });
        if let Some(modified) = file.last_modified() {
            options = options.last_modified_time(modified);
        }

        output_zip.start_file(name, options)?;
        output_zip.write_all(&buf)?;
    }

    if !seen_shared_strings {
        output_zip.start_file(
            "xl/sharedStrings.xml",
            FileOptions::<()>::default().compression_method(CompressionMethod::Deflated),
        )?;
        output_zip.write_all(shared_strings_xml.as_bytes())?;
    }

    output_zip.finish()?;
    Ok(())
}
