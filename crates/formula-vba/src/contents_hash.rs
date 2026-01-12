use crate::{decompress_container, DirStream, OleFile, ParseError};

/// Build the MS-OVBA "ContentNormalizedData" byte sequence for a VBA project.
///
/// This is a building block used by MS-OVBA when computing the VBA project digest that a
/// `\x05DigitalSignature*` stream signs.
///
/// Today this implementation is intentionally minimal and focused on correct module ordering:
/// module source containers are appended in the stored order specified by the `VBA/dir` stream
/// (i.e. `PROJECTMODULES.Modules` order), **not** sorted alphabetically and not based on OLE
/// directory enumeration order.
///
/// Spec reference: MS-OVBA §2.4.2.1 "Content Normalized Data".
pub fn content_normalized_data(vba_project_bin: &[u8]) -> Result<Vec<u8>, ParseError> {
    let mut ole = OleFile::open(vba_project_bin)?;

    let dir_bytes = ole
        .read_stream_opt("VBA/dir")?
        .ok_or(ParseError::MissingStream("VBA/dir"))?;
    let dir_decompressed = decompress_container(&dir_bytes)?;
    let dir_stream = DirStream::parse(&dir_decompressed)?;

    let mut out = Vec::new();
    for module in &dir_stream.modules {
        let stream_path = format!("VBA/{}", module.stream_name);
        let module_stream = ole
            .read_stream_opt(&stream_path)?
            .ok_or(ParseError::MissingStream("module stream"))?;
        let text_offset = module.text_offset.unwrap_or(0).min(module_stream.len());
        let source_container = &module_stream[text_offset..];
        let source_bytes = decompress_container(source_container)?;
        out.extend_from_slice(&source_bytes);
    }

    Ok(out)
}

