use std::ffi::OsString;
use std::io::Read;
use std::path::{Path, PathBuf};
use std::process::ExitCode;

use formula_vba::{project_normalized_data_v3, v3_content_normalized_data};

const DEFAULT_HEAD_BYTES: usize = 64;

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
        "project_normalized_data:    len={} head[0..{}]={}",
        project.len(),
        DEFAULT_HEAD_BYTES.min(project.len()),
        bytes_to_lower_hex(&project[..DEFAULT_HEAD_BYTES.min(project.len())])
    );

    // The debug tool is explicitly interested in the MD5 form of the v3 transcript.
    let contents_hash_v3_md5 = {
        use sha2::Digest as _;
        md5::Md5::digest(&project)
    };
    println!(
        "contents_hash_v3_md5:      {}",
        bytes_to_lower_hex(contents_hash_v3_md5.as_slice())
    );

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
