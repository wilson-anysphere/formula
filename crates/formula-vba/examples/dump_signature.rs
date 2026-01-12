use std::ffi::OsString;
use std::io::Read;
use std::path::{Path, PathBuf};
use std::process::ExitCode;

use formula_vba::{extract_vba_signature_signed_digest, list_vba_digital_signatures};

fn main() -> ExitCode {
    match run() {
        Ok(()) => ExitCode::SUCCESS,
        Err(err) => {
            eprintln!("{err}");
            ExitCode::FAILURE
        }
    }
}

fn run() -> Result<(), String> {
    let mut args = std::env::args_os();
    let program = args
        .next()
        .unwrap_or_else(|| OsString::from("dump_signature"));

    let Some(input) = args.next() else {
        return Err(usage(&program));
    };
    if args.next().is_some() {
        return Err(usage(&program));
    }

    let input_path = PathBuf::from(input);

    let (vba_project_bin, source) = load_vba_project_bin(&input_path)?;

    println!("vbaProject.bin source: {source}");
    println!("vbaProject.bin size: {} bytes", vba_project_bin.len());

    let signatures = list_vba_digital_signatures(&vba_project_bin)
        .map_err(|e| format!("failed to list VBA signature streams: {e}"))?;

    if signatures.is_empty() {
        println!("signature streams: none");
        return Ok(());
    }

    println!("signature streams: {}", signatures.len());
    let mut had_errors = false;
    for (idx, sig) in signatures.iter().enumerate() {
        println!();
        println!("[{}] stream_path: {}", idx + 1, escape_ole_path(&sig.stream_path));
        println!("    stream_len: {} bytes", sig.signature.len());
        println!("    pkcs7_verification: {:?}", sig.verification);
        if matches!(
            sig.verification,
            formula_vba::VbaSignatureVerification::SignedInvalid
                | formula_vba::VbaSignatureVerification::SignedParseError
        ) {
            had_errors = true;
        }
        println!(
            "    signer_subject: {}",
            sig.signer_subject
                .as_deref()
                .unwrap_or("<not found>")
        );

        match (sig.pkcs7_offset, sig.pkcs7_len) {
            (Some(offset), Some(len)) => {
                println!("    pkcs7_location: offset={offset} len={len} (DigSigInfoSerialized)");
            }
            _ => {
                println!("    pkcs7_location: <unknown> (DigSigInfoSerialized not detected)");
            }
        }

        match extract_vba_signature_signed_digest(&sig.signature) {
            Ok(Some(digest)) => {
                println!(
                    "    signed_digest: alg_oid={} digest={}",
                    digest.digest_algorithm_oid,
                    bytes_to_lower_hex(&digest.digest)
                );
            }
            Ok(None) => {
                println!("    signed_digest: <not found>");
            }
            Err(err) => {
                // Keep going: this tool is intended for debugging partially malformed inputs.
                println!("    signed_digest: <error: {err}>");
                had_errors = true;
            }
        }
    }

    if had_errors {
        return Err("one or more signature streams had errors (see stdout)".to_owned());
    }

    Ok(())
}

fn usage(program: &OsString) -> String {
    format!(
        "usage: {} <vbaProject.bin|workbook.xlsm|workbook.xlsx|workbook.xlsb>",
        program.to_string_lossy()
    )
}

fn load_vba_project_bin(path: &Path) -> Result<(Vec<u8>, String), String> {
    match try_extract_vba_project_bin_from_zip(path) {
        Ok(Some(bytes)) => Ok((
            bytes,
            format!("{} (zip entry xl/vbaProject.bin)", path.display()),
        )),
        Ok(None) => {
            // Not a zip workbook; treat as a raw vbaProject.bin OLE file.
            let bytes = std::fs::read(path)
                .map_err(|e| format!("failed to read {}: {e}", path.display()))?;
            Ok((bytes, path.display().to_string()))
        }
        Err(err) => Err(err),
    }
}

fn try_extract_vba_project_bin_from_zip(path: &Path) -> Result<Option<Vec<u8>>, String> {
    let file = match std::fs::File::open(path) {
        Ok(f) => f,
        Err(e) => return Err(format!("failed to open {}: {e}", path.display())),
    };

    let mut archive = match zip::ZipArchive::new(file) {
        Ok(a) => a,
        Err(_) => return Ok(None),
    };

    let mut entry = match archive.by_name("xl/vbaProject.bin") {
        Ok(f) => f,
        Err(zip::result::ZipError::FileNotFound) => {
            return Err(format!(
                "{} is a zip, but does not contain xl/vbaProject.bin",
                path.display()
            ));
        }
        Err(e) => return Err(format!("failed to read zip {}: {e}", path.display())),
    };

    let mut buf = Vec::new();
    entry
        .read_to_end(&mut buf)
        .map_err(|e| format!("failed to read xl/vbaProject.bin from {}: {e}", path.display()))?;
    Ok(Some(buf))
}

fn escape_ole_path(path: &str) -> String {
    path.chars()
        .flat_map(|c| c.escape_default())
        .collect::<String>()
}

fn bytes_to_lower_hex(bytes: &[u8]) -> String {
    let mut out = String::with_capacity(bytes.len() * 2);
    for b in bytes {
        use std::fmt::Write;
        write!(&mut out, "{:02x}", b).expect("writing to String cannot fail");
    }
    out
}
