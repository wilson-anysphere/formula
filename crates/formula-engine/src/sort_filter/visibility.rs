use std::collections::BTreeSet;

#[derive(Debug, Clone, Default, PartialEq, Eq)]
pub struct HiddenRows {
    rows: BTreeSet<usize>,
}

impl HiddenRows {
    pub fn new() -> Self {
        Self::default()
    }

    pub fn insert(&mut self, row: usize) {
        self.rows.insert(row);
    }

    pub fn remove(&mut self, row: usize) {
        self.rows.remove(&row);
    }

    pub fn contains(&self, row: usize) -> bool {
        self.rows.contains(&row)
    }

    pub fn iter(&self) -> impl Iterator<Item = usize> + '_ {
        self.rows.iter().copied()
    }
}

#[derive(Debug, Clone, Copy, PartialEq, Eq)]
pub enum RowVisibility {
    Visible,
    Hidden,
}

