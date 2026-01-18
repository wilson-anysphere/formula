use std::ffi::OsString;
use std::path::PathBuf;
use std::process::ExitCode;

use formula_vba::{project_normalized_data_v3_transcript, v3_content_normalized_data};

#[path = "shared/vba_project_bin.rs"]
mod vba_project_bin;
#[path = "shared/broken_pipe.rs"]
mod broken_pipe;

const DEFAULT_HEAD_BYTES: usize = 64;

#[derive(Debug, Clone, Copy, PartialEq, Eq)]
enum Alg {
    Md5,
    Sha256,
}

fn main() -> ExitCode {
    broken_pipe::install();
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

    let (input_path, password, alg) = parse_args(&program, args)?;

    let (vba_project_bin, source) =
        vba_project_bin::load_vba_project_bin(&input_path, password.as_deref())?;

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

    let project = project_normalized_data_v3_transcript(&vba_project_bin)
        .map_err(|e| format!("failed to compute project_normalized_data_v3_transcript: {e}"))?;
    println!(
        "project_normalized_data_v3_transcript: len={} head[0..{}]={}",
        project.len(),
        DEFAULT_HEAD_BYTES.min(project.len()),
        bytes_to_lower_hex(&project[..DEFAULT_HEAD_BYTES.min(project.len())])
    );

    // Spec note (important):
    //
    // - For legacy signature streams (`\x05DigitalSignature` / `\x05DigitalSignatureEx`), Office
    //   uses a 16-byte MD5 binding digest per MS-OSHARED §4.3 even when
    //   `DigestInfo.digestAlgorithm.algorithm` indicates SHA-256.
    // - For the v3 `\x05DigitalSignatureExt` stream, MS-OVBA §2.4.2.7 defines the v3 content-hash
    //   input as:
    //   `ContentBuffer = V3ContentNormalizedData || ProjectNormalizedData`
    //   and hashes it with a generic `Hash(ContentBuffer)` function (SHA-256 is common in the wild).
    //
    // This tool is a debugging helper that prints MD5/SHA-256 digests over the repo's
    // v3 `project_normalized_data_v3_transcript` transcript. SHA-256 output corresponds to
    // `formula_vba::contents_hash_v3`; MD5 output is provided for experimentation/debugging.
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
                "digest_md5(project_normalized_data_v3_transcript):    {}",
                bytes_to_lower_hex(digest_md5.as_slice())
            );
        }
        Some(Alg::Sha256) => {
            println!(
                "digest_sha256(project_normalized_data_v3_transcript): {}",
                bytes_to_lower_hex(digest_sha256.as_slice())
            );
        }
        None => {
            println!(
                "digest_md5(project_normalized_data_v3_transcript):    {}",
                bytes_to_lower_hex(digest_md5.as_slice())
            );
            println!(
                "digest_sha256(project_normalized_data_v3_transcript): {}",
                bytes_to_lower_hex(digest_sha256.as_slice())
            );
        }
    }

    Ok(())
}

fn usage(program: &OsString) -> String {
    format!(
        "usage: {} --input <vbaProject.bin|workbook.xlsm|workbook.xlsx|workbook.xlsb> [--alg [md5|sha256]] [--password <password>|--password-file <path>]",
        program.to_string_lossy()
    )
}

fn parse_args(
    program: &OsString,
    args: impl Iterator<Item = OsString>,
) -> Result<(PathBuf, Option<String>, Option<Alg>), String> {
    let mut args = args.peekable();
    let mut input: Option<PathBuf> = None;
    let mut password: Option<String> = None;
    let mut password_file: Option<PathBuf> = None;
    let mut alg: Option<Alg> = None;

    while let Some(arg) = args.next() {
        let arg_str = arg.to_string_lossy();

        if arg_str == "--help" || arg_str == "-h" {
            return Err(usage(program));
        }

        if arg_str == "--input" {
            let value = args.next().ok_or_else(|| usage(program))?;
            input = Some(PathBuf::from(value));
            continue;
        }
        if let Some(rest) = arg_str.strip_prefix("--input=") {
            input = Some(PathBuf::from(rest));
            continue;
        }

        if arg_str == "--password" {
            let value = args.next().ok_or_else(|| usage(program))?;
            password = Some(value.to_string_lossy().into_owned());
            continue;
        }
        if let Some(rest) = arg_str.strip_prefix("--password=") {
            password = Some(rest.to_string());
            continue;
        }

        if arg_str == "--password-file" {
            let value = args.next().ok_or_else(|| usage(program))?;
            password_file = Some(PathBuf::from(value));
            continue;
        }
        if let Some(rest) = arg_str.strip_prefix("--password-file=") {
            password_file = Some(PathBuf::from(rest));
            continue;
        }

        if arg_str == "--alg" {
            // Optional value. If omitted, default to SHA-256 (commonly observed for v3 binding
            // digests, and the algorithm used by `formula_vba::contents_hash_v3`).
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

    if password.is_some() && password_file.is_some() {
        return Err(format!(
            "use only one of --password or --password-file\n{}",
            usage(program)
        ));
    }

    let password = match password_file {
        Some(path) => {
            let text = std::fs::read_to_string(&path)
                .map_err(|e| format!("failed to read password file {}: {e}", path.display()))?;
            Some(text.trim_end_matches(['\r', '\n']).to_string())
        }
        None => password,
    };

    let Some(input) = input else {
        return Err(usage(program));
    };

    Ok((input, password, alg))
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

fn bytes_to_lower_hex(bytes: &[u8]) -> String {
    let mut out = String::new();
    let _ = out.try_reserve(bytes.len().saturating_mul(2));
    for b in bytes {
        use std::fmt::Write;
        write!(&mut out, "{:02x}", b).expect("writing to String cannot fail");
    }
    out
}
