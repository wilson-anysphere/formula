use formula_model::Range;
use quick_xml::events::Event;
use quick_xml::Reader;
use std::io::{Read, Seek};
use thiserror::Error;
use zip::ZipArchive;

use crate::zip_util::{open_zip_part, read_zip_file_bytes_with_limit};
use crate::{XlsxError, MAX_XLSX_PACKAGE_PART_BYTES};

#[derive(Debug, Error)]
pub enum MergeCellsError {
    #[error("xml parse error: {0}")]
    Xml(#[from] quick_xml::Error),
    #[error("xml attribute error: {0}")]
    Attr(#[from] quick_xml::events::attributes::AttrError),
    #[error("utf-8 error: {0}")]
    Utf8(#[from] std::str::Utf8Error),
    #[error("invalid merge cell reference: {0}")]
    InvalidRef(String),
    #[error("zip error: {0}")]
    Zip(#[from] zip::result::ZipError),
    #[error("io error: {0}")]
    Io(#[from] std::io::Error),
}

pub fn read_merge_cells_from_worksheet_xml(xml: &str) -> Result<Vec<Range>, MergeCellsError> {
    let mut reader = Reader::from_str(xml);
    reader.config_mut().trim_text(true);

    let mut buf = Vec::new();
    let mut merges = Vec::new();

    loop {
        match reader.read_event_into(&mut buf)? {
            Event::Start(e) | Event::Empty(e) if e.local_name().as_ref() == b"mergeCell" => {
                for attr in e.attributes() {
                    let attr = attr?;
                    if attr.key.as_ref() == b"ref" {
                        let value = std::str::from_utf8(&attr.value)?;
                        let range = Range::from_a1(value)
                            .map_err(|_| MergeCellsError::InvalidRef(value.to_owned()))?;
                        merges.push(range);
                    }
                }
            }
            Event::Eof => break,
            _ => {}
        }
        buf.clear();
    }

    Ok(merges)
}

pub fn read_merge_cells_from_xlsx<R: Read + Seek>(
    archive: &mut ZipArchive<R>,
    worksheet_path: &str,
) -> Result<Vec<Range>, MergeCellsError> {
    let mut file = open_zip_part(archive, worksheet_path)?;
    let bytes =
        read_zip_file_bytes_with_limit(&mut file, worksheet_path, MAX_XLSX_PACKAGE_PART_BYTES)
            .map_err(|err| match err {
                XlsxError::Io(err) => MergeCellsError::Io(err),
                other => MergeCellsError::Io(std::io::Error::new(
                    std::io::ErrorKind::InvalidData,
                    other.to_string(),
                )),
            })?;
    let xml = std::str::from_utf8(&bytes)?;
    read_merge_cells_from_worksheet_xml(xml)
}
