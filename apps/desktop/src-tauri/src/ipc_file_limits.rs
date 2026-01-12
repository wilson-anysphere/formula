pub const MAX_READ_RANGE_BYTES: u64 = 8 * 1024 * 1024; // 8 MiB
pub const MAX_READ_FULL_BYTES: u64 = 64 * 1024 * 1024; // 64 MiB

#[derive(Debug, thiserror::Error, PartialEq, Eq)]
pub enum IpcFileReadLimitError {
    #[error("Requested length {requested} exceeds maximum allowed range read size ({max} bytes)")]
    RangeTooLarge { requested: u64, max: u64 },

    #[error("File size {size} exceeds maximum allowed full read size ({max} bytes)")]
    FileTooLarge { size: u64, max: u64 },

    #[error("Requested length exceeds platform limits")]
    PlatformLimit,
}

pub fn validate_read_range_length(length: u64) -> Result<usize, IpcFileReadLimitError> {
    if length > MAX_READ_RANGE_BYTES {
        return Err(IpcFileReadLimitError::RangeTooLarge {
            requested: length,
            max: MAX_READ_RANGE_BYTES,
        });
    }

    usize::try_from(length).map_err(|_| IpcFileReadLimitError::PlatformLimit)
}

pub fn validate_full_read_size(size: u64) -> Result<(), IpcFileReadLimitError> {
    if size > MAX_READ_FULL_BYTES {
        return Err(IpcFileReadLimitError::FileTooLarge {
            size,
            max: MAX_READ_FULL_BYTES,
        });
    }
    Ok(())
}

#[cfg(test)]
mod tests {
    use super::*;

    #[test]
    fn rejects_oversized_read_range_length() {
        let err = validate_read_range_length(MAX_READ_RANGE_BYTES + 1).unwrap_err();
        assert_eq!(
            err,
            IpcFileReadLimitError::RangeTooLarge {
                requested: MAX_READ_RANGE_BYTES + 1,
                max: MAX_READ_RANGE_BYTES
            }
        );
    }

    #[test]
    fn rejects_oversized_full_read_size() {
        let err = validate_full_read_size(MAX_READ_FULL_BYTES + 1).unwrap_err();
        assert_eq!(
            err,
            IpcFileReadLimitError::FileTooLarge {
                size: MAX_READ_FULL_BYTES + 1,
                max: MAX_READ_FULL_BYTES
            }
        );
    }
}
