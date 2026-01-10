use std::io::{Cursor, Read};

use thiserror::Error;

#[derive(Debug, Error)]
pub enum OleError {
    #[error("io error: {0}")]
    Io(#[from] std::io::Error),
}

/// Thin wrapper around [`cfb::CompoundFile`] that provides path-based stream reads.
pub struct OleFile {
    file: cfb::CompoundFile<Cursor<Vec<u8>>>,
}

impl OleFile {
    pub fn open(data: &[u8]) -> Result<Self, OleError> {
        let cursor = Cursor::new(data.to_vec());
        let file = cfb::CompoundFile::open(cursor)?;
        Ok(Self { file })
    }

    pub fn read_stream_opt(&mut self, path: &str) -> Result<Option<Vec<u8>>, OleError> {
        let mut s = match self.file.open_stream(path) {
            Ok(s) => s,
            Err(err) if err.kind() == std::io::ErrorKind::NotFound => return Ok(None),
            Err(err) => return Err(OleError::Io(err)),
        };

        let mut buf = Vec::new();
        s.read_to_end(&mut buf)?;
        Ok(Some(buf))
    }

    pub fn list_streams(&mut self) -> Result<Vec<String>, OleError> {
        let mut out = Vec::new();
        for entry in self.file.walk() {
            if entry.is_stream() {
                out.push(entry.path().to_string_lossy().to_string());
            }
        }
        Ok(out)
    }
}
