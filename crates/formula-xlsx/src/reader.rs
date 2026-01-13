use crate::XlsxError;
use formula_model::Workbook;
use std::fs::File;
use std::io::{Read, Seek};
use std::path::Path;

pub fn read_workbook(path: impl AsRef<Path>) -> Result<Workbook, XlsxError> {
    let file = File::open(path)?;
    read_workbook_from_reader(file)
}

pub fn read_workbook_from_reader<R: Read + Seek>(reader: R) -> Result<Workbook, XlsxError> {
    crate::read_workbook_model_from_reader(reader).map_err(read_error_to_xlsx_error)
}

fn read_error_to_xlsx_error(err: crate::read::ReadError) -> XlsxError {
    match err {
        crate::read::ReadError::Io(err) => XlsxError::Io(err),
        crate::read::ReadError::Zip(err) => XlsxError::Zip(err),
        crate::read::ReadError::Xml(err) => XlsxError::Xml(err),
        crate::read::ReadError::XmlAttr(err) => XlsxError::Attr(err),
        crate::read::ReadError::Utf8(err) => XlsxError::Invalid(format!("utf-8 error: {err}")),
        crate::read::ReadError::SharedStrings(err) => {
            XlsxError::Invalid(format!("shared strings error: {err}"))
        }
        crate::read::ReadError::Styles(err) => XlsxError::Invalid(format!("styles error: {err}")),
        crate::read::ReadError::InvalidSheetName(err) => {
            XlsxError::Invalid(format!("invalid worksheet name: {err}"))
        }
        crate::read::ReadError::Xlsx(err) => err,
        crate::read::ReadError::MissingPart(part) => XlsxError::MissingPart(part.to_string()),
        crate::read::ReadError::InvalidCellRef(a1) => {
            XlsxError::Invalid(format!("invalid cell reference: {a1}"))
        }
        crate::read::ReadError::InvalidRangeRef(range) => {
            XlsxError::Invalid(format!("invalid range reference: {range}"))
        }
        crate::read::ReadError::InvalidPassword => XlsxError::InvalidPassword,
        crate::read::ReadError::UnsupportedEncryption(msg) => XlsxError::UnsupportedEncryption(msg),
        crate::read::ReadError::InvalidEncryptedWorkbook(msg) => XlsxError::InvalidEncryptedWorkbook(msg),
    }
}
