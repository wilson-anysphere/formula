use std::borrow::Cow;

/// BIFF `CONTINUE` record id.
pub(crate) const RECORD_CONTINUE: u16 = 0x003C;

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
        super::read_biff_record(self.workbook_stream, offset)
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
            }
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
