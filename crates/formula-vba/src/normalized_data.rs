use std::io::{Cursor, Read};

use encoding_rs::{Encoding, WINDOWS_1252};

use crate::dir::ModuleRecord;
use crate::{decompress_container, DirStream, ParseError};

/// Compute the MS-OVBA §2.4.2.2 `FormsNormalizedData` byte sequence for a `vbaProject.bin`.
///
/// This is used as input to the MS-OVBA "Agile Content Hash" algorithm (MS-OVBA §2.4.2.4).
///
/// Spec notes (MS-OVBA §2.4.2.2):
/// - Only **Designer Storages** (MS-OVBA §2.2.10) contribute. The list and ordering of designer
///   modules comes from `ProjectDesignerModule` properties (`BaseClass=...`) in the `PROJECT` stream
///   (MS-OVBA §2.3.1.7).
/// - For each designer module, we normalize the corresponding designer storage named by the
///   `MODULESTREAMNAME` record in `VBA/dir` (MS-OVBA §2.3.4.2.3.2.3).
/// - Within a designer storage we recursively traverse nested storages and include the bytes of
///   every stream encountered. Storage element ordering comes from the compound file's red-black
///   tree (MS-CFB §2.6.4), which sorts siblings by name length first and then by case-insensitive
///   UTF-16 code point value.
/// - Stream bytes are emitted in 1023-byte blocks: the final block of a non-empty stream is padded
///   with `0x00` bytes up to 1023 bytes. Empty streams contribute zero bytes.
pub fn forms_normalized_data(vba_project_bin: &[u8]) -> Result<Vec<u8>, ParseError> {
    // `cfb` works over any `Read + Seek`; use a borrowed cursor to avoid copying the input bytes.
    let cursor = Cursor::new(vba_project_bin);
    let mut file = cfb::CompoundFile::open(cursor).map_err(crate::OleError::from)?;

    let project_bytes = read_required_stream(&mut file, "PROJECT", "PROJECT")?;
    let dir_bytes = read_required_stream(&mut file, "VBA/dir", "VBA/dir")?;
    let dir_decompressed = decompress_container(&dir_bytes)?;

    let encoding = crate::detect_project_codepage(&project_bytes)
        .or_else(|| {
            DirStream::detect_codepage(&dir_decompressed)
                .map(|cp| crate::encoding_for_codepage(cp as u32))
        })
        .unwrap_or(WINDOWS_1252);

    let designer_module_identifiers = parse_project_designer_modules(&project_bytes, encoding);
    if designer_module_identifiers.is_empty() {
        return Ok(Vec::new());
    }

    let dir_stream = DirStream::parse_with_encoding(&dir_decompressed, encoding)?;

    let mut out = Vec::new();
    for module_identifier in designer_module_identifiers {
        let storage_name = match_designer_module_stream_name(&dir_stream.modules, &module_identifier)
            .ok_or(ParseError::MissingStream("designer module"))?;

        let entries = match file.walk_storage(storage_name) {
            Ok(entries) => entries,
            Err(err) if err.kind() == std::io::ErrorKind::NotFound => {
                return Err(ParseError::MissingStorage(storage_name.to_owned()))
            }
            Err(err) => return Err(crate::OleError::from(err).into()),
        };

        // `walk_storage` includes the storage itself as the first entry.
        let mut stream_paths = Vec::new();
        let mut first = true;
        for entry in entries {
            if first {
                first = false;
                continue;
            }
            if entry.is_stream() {
                stream_paths.push(entry.path().to_string_lossy().to_string());
            }
        }

        for path in stream_paths {
            let bytes = read_required_stream(&mut file, &path, "designer stream")?;
            append_stream_padded_1023(&mut out, &bytes);
        }
    }

    Ok(out)
}

fn read_required_stream<F: Read + std::io::Seek>(
    file: &mut cfb::CompoundFile<F>,
    path: &str,
    missing_name: &'static str,
) -> Result<Vec<u8>, ParseError> {
    let mut s = match file.open_stream(path) {
        Ok(s) => s,
        Err(err) if err.kind() == std::io::ErrorKind::NotFound => {
            return Err(ParseError::MissingStream(missing_name))
        }
        Err(err) => return Err(crate::OleError::from(err).into()),
    };
    let mut buf = Vec::new();
    s.read_to_end(&mut buf).map_err(crate::OleError::from)?;
    Ok(buf)
}

fn append_stream_padded_1023(out: &mut Vec<u8>, bytes: &[u8]) {
    out.extend_from_slice(bytes);
    let rem = bytes.len() % 1023;
    if rem != 0 {
        out.extend(std::iter::repeat(0u8).take(1023 - rem));
    }
}

fn parse_project_designer_modules(
    project_stream_bytes: &[u8],
    encoding: &'static Encoding,
) -> Vec<String> {
    let (cow, _, _) = encoding.decode(project_stream_bytes);
    let mut out = Vec::new();
    for line in crate::split_crlf_lines(cow.as_ref()) {
        // `ProjectDesignerModule` properties live in the `ProjectProperties` section of the
        // `PROJECT` stream (MS-OVBA §2.3.1). Stop scanning when we reach a section header such as
        // `[Host Extender Info]` or `[Workspace]` so we don't accidentally treat later sections as
        // designer module declarations.
        if line.starts_with('[') && line.ends_with(']') {
            break;
        }
        let Some((key, rest)) = line.split_once('=') else {
            continue;
        };
        if !key.trim().eq_ignore_ascii_case("BaseClass") {
            continue;
        }
        let ident = rest.trim().trim_matches('"');
        if !ident.is_empty() {
            out.push(ident.to_owned());
        }
    }
    out
}

fn match_designer_module_stream_name<'a>(
    modules: &'a [ModuleRecord],
    module_identifier: &str,
) -> Option<&'a str> {
    if let Some(m) = modules.iter().find(|m| m.name == module_identifier) {
        return Some(m.stream_name.as_str());
    }
    let needle = module_identifier.to_ascii_lowercase();
    modules
        .iter()
        .find(|m| m.name.to_ascii_lowercase() == needle)
        .map(|m| m.stream_name.as_str())
}
