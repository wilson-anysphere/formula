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

    /// British English uses the same separators as `en-US`, but dates are typically interpreted as
    /// day-month-year when parsing ambiguous numeric dates.
    #[must_use]
    pub const fn en_gb() -> Self {
        Self::new(Locale::en_us(), DateOrder::DMY)
    }

    /// Japanese (Japan).
    ///
    /// Assumption: Excel commonly displays numeric dates in YMD order in Japanese locales.
    ///
    /// Note: `formula-format` does not currently ship a `ja-JP` separator preset, so we
    /// temporarily reuse `en-US` punctuation for parsing values.
    #[must_use]
    pub const fn ja_jp() -> Self {
        Self::new(Locale::en_us(), DateOrder::YMD)
    }

    /// Chinese (Simplified, China).
    ///
    /// Assumption: YMD date order is the common default.
    ///
    /// Note: `formula-format` does not currently ship a `zh-CN` separator preset, so we
    /// temporarily reuse `en-US` punctuation for parsing values.
    #[must_use]
    pub const fn zh_cn() -> Self {
        Self::new(Locale::en_us(), DateOrder::YMD)
    }

    /// Korean (Korea).
    ///
    /// Assumption: YMD date order is the common default.
    ///
    /// Note: `formula-format` does not currently ship a `ko-KR` separator preset, so we
    /// temporarily reuse `en-US` punctuation for parsing values.
    #[must_use]
    pub const fn ko_kr() -> Self {
        Self::new(Locale::en_us(), DateOrder::YMD)
    }

    /// Chinese (Traditional, Taiwan).
    ///
    /// Assumption: YMD date order is the common default.
    ///
    /// Note: `formula-format` does not currently ship a `zh-TW` separator preset, so we
    /// temporarily reuse `en-US` punctuation for parsing values.
    #[must_use]
    pub const fn zh_tw() -> Self {
        Self::new(Locale::en_us(), DateOrder::YMD)
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
        let trimmed = id.trim();
        if trimmed.is_empty() {
            return None;
        }

        // Fast path for common canonical ids (avoid allocation + splitting).
        match trimmed {
            "en-US" => return Some(Self::en_us()),
            "en-GB" | "en-UK" | "en-AU" | "en-NZ" | "en-IE" | "en-ZA" => return Some(Self::en_gb()),
            "de-DE" => return Some(Self::de_de()),
            "fr-FR" => return Some(Self::fr_fr()),
            "es-ES" => return Some(Self::es_es()),
            "ja-JP" => return Some(Self::ja_jp()),
            "zh-CN" => return Some(Self::zh_cn()),
            "zh-TW" => return Some(Self::zh_tw()),
            "ko-KR" => return Some(Self::ko_kr()),
            "C" | "POSIX" => return Some(Self::en_us()),
            _ => {}
        }

        let key = super::normalize_locale_key(id)?;
        let parts = super::parse_locale_key(key.as_ref())?;

        match parts.lang {
            // Many POSIX environments report locale as `C` / `POSIX` for the default "C locale".
            // Treat these as `en-US` so callers don't need to special-case.
            "c" | "posix" => Some(Self::en_us()),
            "en" => match parts.region {
                // Note: the formula parsing locale still maps `en-GB` to `en-US` (English function
                // names + `,` argument separators), but value parsing needs the date-order tweak.
                Some("gb") | Some("uk") | Some("au") | Some("nz") | Some("ie") | Some("za") => {
                    Some(Self::en_gb())
                }
                _ => Some(Self::en_us()),
            },
            "ja" => Some(Self::ja_jp()),
            "zh" => {
                // Prefer explicit region codes when present.
                //
                // Otherwise, use the BCP-47 script subtag:
                // - `zh-Hant` is Traditional Chinese, commonly associated with `zh-TW`.
                // - `zh-Hans` is Simplified Chinese, commonly associated with `zh-CN`.
                match parts.region {
                    Some("tw") | Some("hk") | Some("mo") => Some(Self::zh_tw()),
                    Some(_) => Some(Self::zh_cn()),
                    None => match parts.script {
                        Some("hant") => Some(Self::zh_tw()),
                        Some("hans") => Some(Self::zh_cn()),
                        _ => Some(Self::zh_cn()),
                    },
                }
            }
            "ko" => Some(Self::ko_kr()),
            "de" => Some(Self::de_de()),
            "fr" => Some(Self::fr_fr()),
            "es" => Some(Self::es_es()),
            _ => None,
        }
    }
}
