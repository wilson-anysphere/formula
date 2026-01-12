use std::borrow::Cow;

/// BIFF `CONTINUE` record id.
pub(crate) const RECORD_CONTINUE: u16 = 0x003C;
/// BIFF `EOF` record id.
pub(crate) const RECORD_EOF: u16 = 0x000A;
/// BIFF `FILEPASS` record id (workbook encryption / password protection).
///
/// Presence of this record in the workbook globals substream indicates the file
/// is encrypted. We do not currently support decrypting legacy `.xls` files, so
/// the importer uses this as a preflight check to surface a clear error instead
/// of confusing BIFF parse failures from downstream libraries.
pub(crate) const RECORD_FILEPASS: u16 = 0x002F;
/// BIFF8 `BOF` record id.
pub(crate) const RECORD_BOF_BIFF8: u16 = 0x0809;
/// BIFF5 `BOF` record id.
pub(crate) const RECORD_BOF_BIFF5: u16 = 0x0009;

pub(crate) fn is_bof_record(record_id: u16) -> bool {
    record_id == RECORD_BOF_BIFF8 || record_id == RECORD_BOF_BIFF5
}

/// Returns true if the workbook globals substream contains a `FILEPASS` record.
///
/// This is a best-effort scan: malformed/truncated streams simply return `false`.
pub(crate) fn workbook_globals_has_filepass_record(workbook_stream: &[u8]) -> bool {
    // BIFF workbook streams always start with a `BOF` record at offset 0. Guard on that
    // before scanning so we don't misclassify arbitrary/non-Excel streams named `Workbook`
    // as encrypted just because the byte pattern happens to match `FILEPASS`.
    let Some((record_id, _)) = read_biff_record(workbook_stream, 0) else {
        return false;
    };
    if !is_bof_record(record_id) {
        return false;
    }

    let Ok(iter) = BestEffortSubstreamIter::from_offset(workbook_stream, 0) else {
        return false;
    };

    for record in iter {
        if record.record_id == RECORD_FILEPASS {
            return true;
        }

        if record.record_id == RECORD_EOF {
            break;
        }
    }

    false
}

/// Read a single physical BIFF record at `offset`.
pub(crate) fn read_biff_record(workbook_stream: &[u8], offset: usize) -> Option<(u16, &[u8])> {
    let mut iter = BiffRecordIter::from_offset(workbook_stream, offset).ok()?;
    match iter.next()? {
        Ok(record) => Some((record.record_id, record.data)),
        Err(_) => None,
    }
}

#[derive(Debug, Clone, Copy, PartialEq, Eq)]
pub(crate) struct BiffRecord<'a> {
    /// Offset of the record header in the parent stream.
    pub(crate) offset: usize,
    pub(crate) record_id: u16,
    pub(crate) data: &'a [u8],
}

/// Iterator over physical BIFF records.
///
/// This performs bounds checking on the record header and length. A truncated
/// header or payload yields an `Err` and terminates iteration.
pub(crate) struct BiffRecordIter<'a> {
    stream: &'a [u8],
    offset: usize,
}

impl<'a> BiffRecordIter<'a> {
    pub(crate) fn from_offset(stream: &'a [u8], offset: usize) -> Result<Self, String> {
        if offset > stream.len() {
            return Err(format!(
                "BIFF record offset {offset} out of bounds (len={})",
                stream.len()
            ));
        }
        Ok(Self { stream, offset })
    }
}

impl<'a> Iterator for BiffRecordIter<'a> {
    type Item = Result<BiffRecord<'a>, String>;

    fn next(&mut self) -> Option<Self::Item> {
        if self.offset >= self.stream.len() {
            return None;
        }

        let remaining = self.stream.len().saturating_sub(self.offset);
        if remaining < 4 {
            self.offset = self.stream.len();
            return Some(Err("truncated BIFF record header".to_string()));
        }

        let header = &self.stream[self.offset..self.offset + 4];
        let record_id = u16::from_le_bytes([header[0], header[1]]);
        let len = u16::from_le_bytes([header[2], header[3]]) as usize;

        let data_start = match self.offset.checked_add(4) {
            Some(v) => v,
            None => {
                self.offset = self.stream.len();
                return Some(Err("BIFF record offset overflow".to_string()));
            }
        };
        let data_end = match data_start.checked_add(len) {
            Some(v) => v,
            None => {
                self.offset = self.stream.len();
                return Some(Err("BIFF record length overflow".to_string()));
            }
        };

        let data = match self.stream.get(data_start..data_end) {
            Some(data) => data,
            None => {
                let offset = self.offset;
                self.offset = self.stream.len();
                return Some(Err(format!(
                    "BIFF record 0x{record_id:04X} at offset {offset} extends past end of stream (len={}, end={data_end})",
                    self.stream.len()
                )));
            }
        };

        let offset = self.offset;
        self.offset = data_end;
        Some(Ok(BiffRecord {
            offset,
            record_id,
            data,
        }))
    }
}

/// Best-effort iterator over a BIFF substream.
///
/// This is a convenience wrapper around [`BiffRecordIter`] for BIFF sections where we want to
/// recover as much metadata as possible (e.g. BoundSheet, ROW/COLINFO, cell XF indices).
///
/// - Stops before yielding a *subsequent* `BOF` record (since that indicates the start of the next
///   substream; truncated streams sometimes omit the expected `EOF`).
/// - Stops on a malformed/truncated physical record instead of returning an error.
pub(crate) struct BestEffortSubstreamIter<'a> {
    iter: BiffRecordIter<'a>,
    start_offset: usize,
    finished: bool,
}

impl<'a> BestEffortSubstreamIter<'a> {
    pub(crate) fn from_offset(stream: &'a [u8], start_offset: usize) -> Result<Self, String> {
        Ok(Self {
            iter: BiffRecordIter::from_offset(stream, start_offset)?,
            start_offset,
            finished: false,
        })
    }
}

impl<'a> Iterator for BestEffortSubstreamIter<'a> {
    type Item = BiffRecord<'a>;

    fn next(&mut self) -> Option<Self::Item> {
        if self.finished {
            return None;
        }

        let record = match self.iter.next()? {
            Ok(record) => record,
            Err(_) => {
                self.finished = true;
                return None;
            }
        };

        if record.offset != self.start_offset && is_bof_record(record.record_id) {
            self.finished = true;
            return None;
        }

        Some(record)
    }
}

/// A logical BIFF record. Some BIFF record types may be split across one or more
/// physical `CONTINUE` records; those fragments are concatenated into `data`.
///
/// `fragment_sizes` stores the size of each physical fragment in `data` order,
/// allowing parsers to reason about `CONTINUE` boundaries when needed (e.g.
/// continued BIFF8 strings).
#[derive(Debug, Clone)]
pub(crate) struct LogicalBiffRecord<'a> {
    /// Byte offset of the physical record header in the parent stream.
    pub(crate) offset: usize,
    pub(crate) record_id: u16,
    pub(crate) data: Cow<'a, [u8]>,
    pub(crate) fragment_sizes: Vec<usize>,
}

impl<'a> LogicalBiffRecord<'a> {
    pub(crate) fn is_continued(&self) -> bool {
        self.fragment_sizes.len() > 1
    }

    pub(crate) fn first_fragment(&self) -> &[u8] {
        let first_len = self.fragment_sizes.first().copied().unwrap_or(0);
        self.data.get(0..first_len).unwrap_or_default()
    }

    pub(crate) fn fragments(&self) -> FragmentIter<'_> {
        FragmentIter {
            data: self.data.as_ref(),
            sizes: &self.fragment_sizes,
            idx: 0,
            offset: 0,
        }
    }
}

pub(crate) struct FragmentIter<'a> {
    data: &'a [u8],
    sizes: &'a [usize],
    idx: usize,
    offset: usize,
}

impl<'a> Iterator for FragmentIter<'a> {
    type Item = &'a [u8];

    fn next(&mut self) -> Option<Self::Item> {
        let size = *self.sizes.get(self.idx)?;
        let start = self.offset;
        let end = start.checked_add(size)?;
        let out = self.data.get(start..end)?;
        self.idx = self.idx.saturating_add(1);
        self.offset = end;
        Some(out)
    }
}

/// Iterates over BIFF records, combining `CONTINUE` fragments for record ids for
/// which `allows_continuation(record_id) == true`.
pub(crate) struct LogicalBiffRecordIter<'a> {
    iter: std::iter::Peekable<BiffRecordIter<'a>>,
    allows_continuation: fn(u16) -> bool,
}

impl<'a> LogicalBiffRecordIter<'a> {
    pub(crate) fn new(workbook_stream: &'a [u8], allows_continuation: fn(u16) -> bool) -> Self {
        Self::from_offset(workbook_stream, 0, allows_continuation)
            .expect("offset 0 should always be in bounds for BIFF streams")
    }

    pub(crate) fn from_offset(
        workbook_stream: &'a [u8],
        offset: usize,
        allows_continuation: fn(u16) -> bool,
    ) -> Result<Self, String> {
        Ok(Self {
            iter: BiffRecordIter::from_offset(workbook_stream, offset)?.peekable(),
            allows_continuation,
        })
    }
}

impl<'a> Iterator for LogicalBiffRecordIter<'a> {
    type Item = Result<LogicalBiffRecord<'a>, String>;

    fn next(&mut self) -> Option<Self::Item> {
        let first = match self.iter.next()? {
            Ok(record) => record,
            Err(err) => return Some(Err(err)),
        };

        let start_offset = first.offset;
        let record_id = first.record_id;
        let data = first.data;

        if !(self.allows_continuation)(record_id) {
            return Some(Ok(LogicalBiffRecord {
                offset: start_offset,
                record_id,
                data: Cow::Borrowed(data),
                fragment_sizes: vec![data.len()],
            }));
        }

        // Only allocate/copy when we actually see a CONTINUE record.
        match self.iter.peek() {
            Some(Ok(next)) if next.record_id == RECORD_CONTINUE => {}
            _ => {
                return Some(Ok(LogicalBiffRecord {
                    offset: start_offset,
                    record_id,
                    data: Cow::Borrowed(data),
                    fragment_sizes: vec![data.len()],
                }))
            }
        };

        let mut fragment_sizes = vec![data.len()];
        let mut combined: Vec<u8> = data.to_vec();

        // Collect subsequent CONTINUE records into one logical payload.
        while let Some(peek) = self.iter.peek() {
            let next = match peek {
                Ok(next) => next,
                // Leave the malformed record to be surfaced on the next iteration.
                Err(_) => break,
            };
            if next.record_id != RECORD_CONTINUE {
                break;
            }

            let next = match self
                .iter
                .next()
                .expect("peek verified there is another record")
            {
                Ok(record) => record,
                Err(err) => return Some(Err(err)),
            };
            fragment_sizes.push(next.data.len());
            combined.extend_from_slice(next.data);
        }

        Some(Ok(LogicalBiffRecord {
            offset: start_offset,
            record_id,
            data: Cow::Owned(combined),
            fragment_sizes,
        }))
    }
}

#[cfg(test)]
mod tests {
    use super::*;

    fn record(id: u16, payload: &[u8]) -> Vec<u8> {
        let mut out = Vec::with_capacity(4 + payload.len());
        out.extend_from_slice(&id.to_le_bytes());
        out.extend_from_slice(&(payload.len() as u16).to_le_bytes());
        out.extend_from_slice(payload);
        out
    }

    #[test]
    fn best_effort_iter_stops_at_next_bof() {
        let prefix = record(0x0001, &[0xAA]);
        let start_offset = prefix.len();

        let stream = [
            prefix,
            record(RECORD_BOF_BIFF8, &[0u8; 16]),
            record(0x0002, &[0xBB]),
            record(RECORD_BOF_BIFF8, &[0u8; 16]),
            record(0x0003, &[0xCC]),
        ]
        .concat();

        let iter = BestEffortSubstreamIter::from_offset(&stream, start_offset).unwrap();
        let ids: Vec<u16> = iter.map(|r| r.record_id).collect();
        assert_eq!(ids, vec![RECORD_BOF_BIFF8, 0x0002]);
    }

    #[test]
    fn best_effort_iter_stops_on_malformed_record() {
        // Truncated record: declares 4 bytes but only provides 2.
        let mut truncated = Vec::new();
        truncated.extend_from_slice(&0x0002u16.to_le_bytes());
        truncated.extend_from_slice(&4u16.to_le_bytes());
        truncated.extend_from_slice(&[0xAA, 0xBB]);

        // The truncated record must be at the end of the stream so the physical iterator detects
        // that it extends past the end of the buffer.
        let stream = [record(0x0001, &[1]), truncated].concat();

        let iter = BestEffortSubstreamIter::from_offset(&stream, 0).unwrap();
        let ids: Vec<u16> = iter.map(|r| r.record_id).collect();
        assert_eq!(ids, vec![0x0001]);
    }

    #[test]
    fn filepass_scan_requires_bof_at_offset_zero() {
        // Construct a stream that *contains* a FILEPASS record, but does not start with a BOF
        // record. This should not be treated as an encrypted workbook stream.
        let stream = [
            record(0x0001, &[0xAA]),
            record(RECORD_FILEPASS, &[0x00, 0x00]),
            record(RECORD_EOF, &[]),
        ]
        .concat();

        assert!(!workbook_globals_has_filepass_record(&stream));
    }

    #[test]
    fn filepass_scan_detects_encryption_when_bof_present() {
        let stream = [
            record(RECORD_BOF_BIFF8, &[0u8; 16]),
            record(RECORD_FILEPASS, &[0x00, 0x00]),
            record(RECORD_EOF, &[]),
        ]
        .concat();

        assert!(workbook_globals_has_filepass_record(&stream));
    }

    #[test]
    fn iterates_physical_records_with_bounds_checks() {
        let stream = [record(0x0001, &[1, 2, 3]), record(0x0002, &[4])].concat();
        let mut iter = BiffRecordIter::from_offset(&stream, 0).unwrap();

        let r1 = iter.next().unwrap().unwrap();
        assert_eq!(r1.offset, 0);
        assert_eq!(r1.record_id, 0x0001);
        assert_eq!(r1.data, &[1, 2, 3]);

        let r2 = iter.next().unwrap().unwrap();
        assert_eq!(r2.record_id, 0x0002);
        assert_eq!(r2.data, &[4]);

        assert!(iter.next().is_none());
    }

    #[test]
    fn physical_iter_errors_on_truncated_header() {
        let stream = vec![0x01, 0x02, 0x03];
        let mut iter = BiffRecordIter::from_offset(&stream, 0).unwrap();
        let err = iter.next().unwrap().unwrap_err();
        assert!(err.contains("truncated BIFF record header"), "err={err}");
        assert!(iter.next().is_none());
    }

    #[test]
    fn physical_iter_errors_on_truncated_payload() {
        let mut stream = Vec::new();
        stream.extend_from_slice(&0x0001u16.to_le_bytes());
        stream.extend_from_slice(&4u16.to_le_bytes());
        stream.extend_from_slice(&[1, 2]);

        let mut iter = BiffRecordIter::from_offset(&stream, 0).unwrap();
        let err = iter.next().unwrap().unwrap_err();
        assert!(err.contains("extends past end of stream"), "err={err}");
        assert!(iter.next().is_none());
    }

    #[test]
    fn coalesces_continues_for_allowed_record_ids() {
        let stream = [
            record(0x00AA, &[1, 2]),
            record(RECORD_CONTINUE, &[3]),
            record(RECORD_CONTINUE, &[4, 5]),
            record(0x00BB, &[9]),
        ]
        .concat();

        let allows = |id: u16| id == 0x00AA;
        let mut iter = LogicalBiffRecordIter::new(&stream, allows);

        let first = iter.next().unwrap().unwrap();
        assert_eq!(first.record_id, 0x00AA);
        assert_eq!(first.data.as_ref(), &[1, 2, 3, 4, 5]);
        assert_eq!(first.fragment_sizes, vec![2, 1, 2]);

        let second = iter.next().unwrap().unwrap();
        assert_eq!(second.record_id, 0x00BB);
        assert_eq!(second.data.as_ref(), &[9]);
        assert_eq!(second.fragment_sizes, vec![1]);

        assert!(iter.next().is_none());
    }

    #[test]
    fn does_not_coalesce_when_continuation_is_disallowed() {
        let stream = [record(0x00AA, &[1, 2]), record(RECORD_CONTINUE, &[3])].concat();
        let mut iter = LogicalBiffRecordIter::new(&stream, |_| false);

        let first = iter.next().unwrap().unwrap();
        assert_eq!(first.data.as_ref(), &[1, 2]);

        // CONTINUE becomes its own logical record when the parent doesn't allow continuation.
        let second = iter.next().unwrap().unwrap();
        assert_eq!(second.record_id, RECORD_CONTINUE);
        assert_eq!(second.data.as_ref(), &[3]);
    }

    #[test]
    fn logical_iter_from_offset_starts_at_record_boundary() {
        let prefix = record(0x0001, &[0xAA]);
        let start_offset = prefix.len();

        let stream = [
            prefix,
            record(0x00AA, &[1, 2]),
            record(RECORD_CONTINUE, &[3]),
            record(0x00BB, &[9]),
        ]
        .concat();

        let allows = |id: u16| id == 0x00AA;
        let mut iter = LogicalBiffRecordIter::from_offset(&stream, start_offset, allows).unwrap();

        let first = iter.next().unwrap().unwrap();
        assert_eq!(first.record_id, 0x00AA);
        assert_eq!(first.data.as_ref(), &[1, 2, 3]);

        let second = iter.next().unwrap().unwrap();
        assert_eq!(second.record_id, 0x00BB);
        assert_eq!(second.data.as_ref(), &[9]);
        assert!(iter.next().is_none());
    }
}
