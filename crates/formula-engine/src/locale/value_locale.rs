use formula_format::Locale;

/// Date component order to use when parsing ambiguous numeric dates like `1/2/2024`.
#[derive(Debug, Clone, Copy, PartialEq, Eq)]
pub enum DateOrder {
    /// Month-Day-Year (`1/2/2024` -> Jan 2, 2024).
    MDY,
    /// Day-Month-Year (`1/2/2024` -> Feb 1, 2024).
    DMY,
    /// Year-Month-Day (`2024/1/2` -> Jan 2, 2024).
    YMD,
}

/// Locale configuration used when parsing *values* (text -> number/date/time).
///
/// This is distinct from the formula parsing locale ([`crate::LocaleConfig`]), which controls
/// tokenization/serialization of the formula language (argument separators, decimal separators in
/// numeric literals, etc.).
#[derive(Debug, Clone, Copy, PartialEq, Eq)]
pub struct ValueLocaleConfig {
    pub separators: Locale,
    pub date_order: DateOrder,
}

impl Default for ValueLocaleConfig {
    fn default() -> Self {
        Self::en_us()
    }
}

impl ValueLocaleConfig {
    #[must_use]
    pub const fn new(separators: Locale, date_order: DateOrder) -> Self {
        Self {
            separators,
            date_order,
        }
    }

    #[must_use]
    pub const fn en_us() -> Self {
        Self::new(Locale::en_us(), DateOrder::MDY)
    }

    #[must_use]
    pub const fn de_de() -> Self {
        Self::new(Locale::de_de(), DateOrder::DMY)
    }

    #[must_use]
    pub const fn fr_fr() -> Self {
        Self::new(Locale::fr_fr(), DateOrder::DMY)
    }

    #[must_use]
    pub const fn es_es() -> Self {
        Self::new(Locale::es_es(), DateOrder::DMY)
    }

    #[must_use]
    pub fn for_locale_id(id: &str) -> Option<Self> {
        match super::normalize_locale_id(id)? {
            "en-US" => Some(Self::en_us()),
            "de-DE" => Some(Self::de_de()),
            "fr-FR" => Some(Self::fr_fr()),
            "es-ES" => Some(Self::es_es()),
            _ => None,
        }
    }
}
