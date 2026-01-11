pub(crate) mod number;
pub mod datetime;

#[derive(Debug, Clone, Copy, PartialEq, Eq)]
pub enum DateOrder {
    /// Month / day / year (e.g. `12/31/2024`).
    Mdy,
    /// Day / month / year (e.g. `31/12/2024`).
    Dmy,
}

impl Default for DateOrder {
    fn default() -> Self {
        Self::Mdy
    }
}

#[derive(Debug, Clone, Copy, PartialEq, Eq)]
pub struct ValueLocaleConfig {
    pub decimal_separator: char,
    pub group_separator: char,
    pub date_order: DateOrder,
}

impl ValueLocaleConfig {
    #[must_use]
    pub const fn en_us() -> Self {
        Self {
            decimal_separator: '.',
            group_separator: ',',
            date_order: DateOrder::Mdy,
        }
    }
}
