use std::ffi::OsString;
use std::io::Read;
use std::path::{Path, PathBuf};
use std::process::ExitCode;

use formula_vba::{project_normalized_data_v3, v3_content_normalized_data};

const DEFAULT_HEAD_BYTES: usize = 64;

#[derive(Debug, Clone, Copy, PartialEq, Eq)]
enum Alg {
    Md5,
    Sha256,
}

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
        .unwrap_or_else(|| OsString::from("dump_contents_hash_v3"));

    let (input_path, alg) = parse_args(&program, args)?;

    let (vba_project_bin, source) = load_vba_project_bin(&input_path)?;

    println!("vbaProject.bin source: {source}");
    println!("vbaProject.bin size: {} bytes", vba_project_bin.len());
    println!();

    let v3 = v3_content_normalized_data(&vba_project_bin)
        .map_err(|e| format!("failed to compute V3ContentNormalizedData: {e}"))?;
    println!(
        "v3_content_normalized_data: len={} head[0..{}]={}",
        v3.len(),
        DEFAULT_HEAD_BYTES.min(v3.len()),
        bytes_to_lower_hex(&v3[..DEFAULT_HEAD_BYTES.min(v3.len())])
    );

    let project = project_normalized_data_v3(&vba_project_bin)
        .map_err(|e| format!("failed to compute ProjectNormalizedData: {e}"))?;
    println!(
        "project_normalized_data_v3: len={} head[0..{}]={}",
        project.len(),
        DEFAULT_HEAD_BYTES.min(project.len()),
        bytes_to_lower_hex(&project[..DEFAULT_HEAD_BYTES.min(project.len())])
    );

    // MS-OVBA v3 signature binding hashes the *same transcript* (`ProjectNormalizedData`), but the
    // on-disk `DigestInfo.digestAlgorithm` can vary (SHA-256 is common, but some producers use MD5).
    //
    // This tool prints both digests by default so you can compare against the on-disk digest bytes.
    let digest_md5 = {
        use md5::Digest as _;
        md5::Md5::digest(&project)
    };
    let digest_sha256 = {
        use sha2::Digest as _;
        sha2::Sha256::digest(&project)
    };

    match alg {
        Some(Alg::Md5) => {
            println!(
                "digest_md5(ProjectNormalizedData):    {}",
                bytes_to_lower_hex(digest_md5.as_slice())
            );
        }
        Some(Alg::Sha256) => {
            println!(
                "digest_sha256(ProjectNormalizedData): {}",
                bytes_to_lower_hex(digest_sha256.as_slice())
            );
        }
        None => {
            println!(
                "digest_md5(ProjectNormalizedData):    {}",
                bytes_to_lower_hex(digest_md5.as_slice())
            );
            println!(
                "digest_sha256(ProjectNormalizedData): {}",
                bytes_to_lower_hex(digest_sha256.as_slice())
            );
        }
    }

    Ok(())
}

fn usage(program: &OsString) -> String {
    format!(
        "usage: {} [--alg [md5|sha256]] <vbaProject.bin|workbook.xlsm|workbook.xlsx|workbook.xlsb>",
        program.to_string_lossy()
    )
}

fn parse_args(
    program: &OsString,
    args: impl Iterator<Item = OsString>,
) -> Result<(PathBuf, Option<Alg>), String> {
    let mut args = args.peekable();
    let mut input: Option<PathBuf> = None;
    let mut alg: Option<Alg> = None;

    while let Some(arg) = args.next() {
        let arg_str = arg.to_string_lossy();

        if arg_str == "--help" || arg_str == "-h" {
            return Err(usage(program));
        }

        if arg_str == "--alg" {
            // Optional value. If omitted, default to SHA-256 (the common case).
            //
            // This allows `--alg <input-path>` as shorthand for printing only the SHA-256 digest.
            if let Some(value) = args.peek() {
                if let Some(parsed) = parse_alg(value) {
                    // Consume the value.
                    args.next();
                    alg = Some(parsed);
                    continue;
                }

                // If an input path was already provided, a non-algorithm value after `--alg`
                // indicates a mistake; treat as an error rather than silently defaulting.
                if input.is_some() && !value.to_string_lossy().starts_with('-') {
                    return Err(format!(
                        "invalid --alg value (expected md5|sha256): {}\n{}",
                        value.to_string_lossy(),
                        usage(program)
                    ));
                }
            }

            alg = Some(Alg::Sha256);
            continue;
        }

        if let Some(rest) = arg_str.strip_prefix("--alg=") {
            alg = Some(parse_alg_str(rest).ok_or_else(|| {
                format!(
                    "invalid --alg value (expected md5|sha256): {rest}\n{}",
                    usage(program)
                )
            })?);
            continue;
        }

        if input.is_none() {
            input = Some(PathBuf::from(arg));
            continue;
        }

        return Err(usage(program));
    }

    let Some(input) = input else {
        return Err(usage(program));
    };
    Ok((input, alg))
}

fn parse_alg(arg: &OsString) -> Option<Alg> {
    parse_alg_str(&arg.to_string_lossy())
}

fn parse_alg_str(s: &str) -> Option<Alg> {
    if s.eq_ignore_ascii_case("md5") {
        return Some(Alg::Md5);
    }
    if s.eq_ignore_ascii_case("sha256") {
        return Some(Alg::Sha256);
    }
    None
}

fn load_vba_project_bin(path: &Path) -> Result<(Vec<u8>, String), String> {
    match try_extract_vba_project_bin_from_zip(path) {
        Ok(Some(bytes)) => Ok((
            bytes,
            format!("{} (zip entry xl/vbaProject.bin)", path.display()),
        )),
        Ok(None) => {
            // Not a zip workbook; treat as a raw vbaProject.bin OLE file.
            let bytes =
                std::fs::read(path).map_err(|e| format!("failed to read {}: {e}", path.display()))?;
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

fn bytes_to_lower_hex(bytes: &[u8]) -> String {
    let mut out = String::with_capacity(bytes.len() * 2);
    for b in bytes {
        use std::fmt::Write;
        write!(&mut out, "{:02x}", b).expect("writing to String cannot fail");
    }
    out
}
