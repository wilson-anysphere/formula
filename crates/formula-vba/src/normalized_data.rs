use std::collections::{HashMap, HashSet};

use encoding_rs::{Encoding, WINDOWS_1252};

use crate::dir::ModuleRecord;
use crate::{decompress_container, DirStream, OleFile, ParseError};

/// Compute the MS-OVBA §2.4.2.2 `FormsNormalizedData` byte sequence for a `vbaProject.bin`.
///
/// This is used as input to the MS-OVBA §2.4.2.4 "Agile Content Hash" algorithm, which extends the
/// legacy Content Hash by incorporating designer storages (UserForms / designers).
///
/// The spec describes iterating storage elements "in stored order". The `cfb` crate does not expose
/// the raw directory sibling ordering, so for determinism (and compatibility with the way many OLE
/// producers sort entries) we traverse each storage's immediate children in **case-insensitive OLE
/// entry name** order.
pub fn forms_normalized_data(vba_project_bin: &[u8]) -> Result<Vec<u8>, ParseError> {
    let mut ole = OleFile::open(vba_project_bin)?;

    // Read + decompress `VBA/dir` so we can map designer module identifiers to MODULESTREAMNAME.
    let dir_bytes = ole
        .read_stream_opt("VBA/dir")?
        .ok_or(ParseError::MissingStream("VBA/dir"))?;
    let dir_decompressed = decompress_container(&dir_bytes)?;
    let project_bytes = ole
        .read_stream_opt("PROJECT")?
        .ok_or(ParseError::MissingStream("PROJECT"))?;

    // Decode the `PROJECT` stream using its CodePage= (preferred) or the `PROJECTCODEPAGE` record
    // from `VBA/dir` as a fallback.
    let encoding = crate::detect_project_codepage(&project_bytes)
        .or_else(|| {
            DirStream::detect_codepage(&dir_decompressed)
                .map(|cp| crate::encoding_for_codepage(cp as u32))
        })
        .unwrap_or(WINDOWS_1252);

    let dir_stream = DirStream::parse_with_encoding(&dir_decompressed, encoding)?;

    // MS-OVBA §2.3.1.7: designer modules are identified in the `PROJECT` stream by `BaseClass=` lines.
    let designer_module_identifiers = parse_project_designer_modules(&project_bytes, encoding);
    if designer_module_identifiers.is_empty() {
        return Ok(Vec::new());
    }

    let streams = ole.list_streams()?;

    let mut out = Vec::new();

    // Avoid hashing the same designer storage twice if the PROJECT stream contains duplicates.
    let mut seen = HashSet::<String>::new();
    for module_identifier in designer_module_identifiers {
        if !seen.insert(module_identifier.clone()) {
            continue;
        }

        let storage_name =
            match_designer_module_stream_name(&dir_stream.modules, &module_identifier)
                .ok_or(ParseError::MissingStream("designer module"))?;

        // Per MS-OVBA §2.2.10, a designer storage MUST exist at the OLE root with a name equal to
        // MODULESTREAMNAME. We approximate this by requiring at least one stream with that prefix.
        let prefix = format!("{}/", storage_name);
        if !streams.iter().any(|p| p.starts_with(&prefix)) {
            return Err(ParseError::MissingStorage(storage_name.to_owned()));
        }

        normalize_storage_into_vec(&mut ole, &streams, storage_name, &mut out)?;
    }

    Ok(out)
}

fn parse_project_designer_modules(
    project_stream_bytes: &[u8],
    encoding: &'static Encoding,
) -> Vec<String> {
    let (cow, _, _) = encoding.decode(project_stream_bytes);
    let mut out = Vec::new();
    for line in cow.lines() {
        let line = line.trim_end_matches('\r').trim();
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

fn normalize_storage_into_vec(
    ole: &mut OleFile,
    streams: &[String],
    storage_path: &str,
    out: &mut Vec<u8>,
) -> Result<(), ParseError> {
    // Build a list of immediate children of this storage (streams or nested storages) from the
    // flattened stream path list.
    #[derive(Debug, Clone)]
    enum Child {
        Stream { name: String, path: String },
        Storage { name: String, path: String },
    }

    impl Child {
        fn name(&self) -> &str {
            match self {
                Child::Stream { name, .. } => name,
                Child::Storage { name, .. } => name,
            }
        }
    }

    let prefix = format!("{}/", storage_path);
    let mut children_by_name: HashMap<String, Child> = HashMap::new();
    for p in streams {
        let Some(rel) = p.strip_prefix(&prefix) else {
            continue;
        };
        let mut parts = rel.split('/');
        let Some(first) = parts.next() else {
            continue;
        };
        if first.is_empty() {
            continue;
        }

        let child_path = format!("{}/{}", storage_path, first);
        let child = if parts.next().is_some() {
            Child::Storage {
                name: first.to_owned(),
                path: child_path,
            }
        } else {
            Child::Stream {
                name: first.to_owned(),
                path: child_path,
            }
        };

        match children_by_name.entry(first.to_owned()) {
            std::collections::hash_map::Entry::Vacant(v) => {
                v.insert(child);
            }
            std::collections::hash_map::Entry::Occupied(mut o) => {
                // If both a stream and a storage exist with the same name (shouldn't happen in a
                // valid CFB), prefer treating it as a storage so nested children are processed.
                if matches!(o.get(), Child::Stream { .. }) && matches!(child, Child::Storage { .. })
                {
                    o.insert(child);
                }
            }
        }
    }

    let mut children = children_by_name.into_values().collect::<Vec<_>>();
    children.sort_by(|a, b| {
        a.name()
            .to_ascii_lowercase()
            .cmp(&b.name().to_ascii_lowercase())
            .then_with(|| a.name().cmp(b.name()))
    });

    for child in children {
        match child {
            Child::Stream { path, .. } => {
                let stream_bytes = ole
                    .read_stream_opt(&path)?
                    .ok_or(ParseError::MissingStream("designer stream"))?;
                for chunk in stream_bytes.chunks(1023) {
                    let mut block = [0u8; 1023];
                    block[..chunk.len()].copy_from_slice(chunk);
                    out.extend_from_slice(&block);
                }
            }
            Child::Storage { path, .. } => {
                normalize_storage_into_vec(ole, streams, &path, out)?;
            }
        }
    }

    Ok(())
}
