use formula_model::{ArgbColor, ThemePalette as ModelThemePalette};

use super::ThemePalette as XlsxThemePalette;

pub(crate) fn to_model_theme_palette(palette: XlsxThemePalette) -> ModelThemePalette {
    ModelThemePalette {
        dk1: ArgbColor(palette.dk1),
        lt1: ArgbColor(palette.lt1),
        dk2: ArgbColor(palette.dk2),
        lt2: ArgbColor(palette.lt2),
        accent1: ArgbColor(palette.accent1),
        accent2: ArgbColor(palette.accent2),
        accent3: ArgbColor(palette.accent3),
        accent4: ArgbColor(palette.accent4),
        accent5: ArgbColor(palette.accent5),
        accent6: ArgbColor(palette.accent6),
        hlink: ArgbColor(palette.hlink),
        fol_hlink: ArgbColor(palette.followed_hlink),
    }
}

