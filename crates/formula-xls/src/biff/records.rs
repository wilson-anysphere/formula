use std::borrow::Cow;

/// BIFF `CONTINUE` record id.
pub(crate) const RECORD_CONTINUE: u16 = 0x003C;

pub(crate) fn is_bof_record(record_id: u16) -> bool {
    record_id == 0x0809 || record_id == 0x0009
}

/// Read a single physical BIFF record at `offset`.
pub(crate) fn read_biff_record(workbook_stream: &[u8], offset: usize) -> Option<(u16, &[u8])> {
    let header = workbook_stream.get(offset..offset + 4)?;
    let record_id = u16::from_le_bytes([header[0], header[1]]);
    let len = u16::from_le_bytes([header[2], header[3]]) as usize;
    let data_start = offset + 4;
    let data_end = data_start.checked_add(len)?;
    let data = workbook_stream.get(data_start..data_end)?;
    Some((record_id, data))
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

    #[allow(dead_code)]
    pub(crate) fn offset(&self) -> usize {
        self.offset
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

    pub(crate) fn fragments(&'a self) -> FragmentIter<'a> {
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
    workbook_stream: &'a [u8],
    offset: usize,
    allows_continuation: fn(u16) -> bool,
}

impl<'a> LogicalBiffRecordIter<'a> {
    pub(crate) fn new(workbook_stream: &'a [u8], allows_continuation: fn(u16) -> bool) -> Self {
        Self {
            workbook_stream,
            offset: 0,
            allows_continuation,
        }
    }

    fn read_next_physical(&self, offset: usize) -> Option<(u16, &'a [u8])> {
        read_biff_record(self.workbook_stream, offset)
    }
}

impl<'a> Iterator for LogicalBiffRecordIter<'a> {
    type Item = Result<LogicalBiffRecord<'a>, String>;

    fn next(&mut self) -> Option<Self::Item> {
        let start_offset = self.offset;
        let (record_id, data) = self.read_next_physical(start_offset)?;

        let mut next_offset = match start_offset
            .checked_add(4)
            .and_then(|o| o.checked_add(data.len()))
        {
            Some(offset) => offset,
            None => {
                self.offset = self.workbook_stream.len();
                return Some(Err("BIFF record offset overflow".to_string()));
            }
        };

        if !(self.allows_continuation)(record_id) {
            self.offset = next_offset;
            return Some(Ok(LogicalBiffRecord {
                offset: start_offset,
                record_id,
                data: Cow::Borrowed(data),
                fragment_sizes: vec![data.len()],
            }));
        }

        // Only allocate/copy when we actually see a CONTINUE record.
        let Some((peek_id, _)) = self.read_next_physical(next_offset) else {
            self.offset = next_offset;
            return Some(Ok(LogicalBiffRecord {
                offset: start_offset,
                record_id,
                data: Cow::Borrowed(data),
                fragment_sizes: vec![data.len()],
            }));
        };
        if peek_id != RECORD_CONTINUE {
            self.offset = next_offset;
            return Some(Ok(LogicalBiffRecord {
                offset: start_offset,
                record_id,
                data: Cow::Borrowed(data),
                fragment_sizes: vec![data.len()],
            }));
        }

        let mut fragment_sizes = vec![data.len()];
        let mut combined: Vec<u8> = data.to_vec();

        // Collect subsequent CONTINUE records into one logical payload.
        while let Some((next_id, next_data)) = self.read_next_physical(next_offset) {
            if next_id != RECORD_CONTINUE {
                break;
            }

            fragment_sizes.push(next_data.len());
            combined.extend_from_slice(next_data);
            next_offset = match next_offset
                .checked_add(4)
                .and_then(|o| o.checked_add(next_data.len()))
            {
                Some(offset) => offset,
                None => {
                    self.offset = self.workbook_stream.len();
                    return Some(Err("BIFF record offset overflow".to_string()));
                }
            };
        }

        self.offset = next_offset;
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
}
