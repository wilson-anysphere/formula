use std::collections::BTreeMap;

use formula_model::{
    indexed_color_argb, Color, ColorContext, ThemeColorSlot, ThemePalette, Workbook,
};

#[test]
fn tint_math_matches_excel() {
    // Base = 0x64 (100) so the expected values are easy to reason about.
    let base = 0xFF646464;
    let mut palette = ThemePalette::default();
    palette.lt1 = formula_model::ArgbColor(base);

    let theme_idx = ThemeColorSlot::Lt1.theme_index();

    assert_eq!(
        Color::Theme {
            theme: theme_idx,
            tint: None
        }
        .resolve(Some(&palette)),
        Some(base)
    );

    assert_eq!(
        Color::Theme {
            theme: theme_idx,
            tint: Some(0)
        }
        .resolve(Some(&palette)),
        Some(base)
    );

    // +0.5 (lighten): 100 * (1-0.5) + 255 * 0.5 = 177.5 -> 178
    assert_eq!(
        Color::Theme {
            theme: theme_idx,
            tint: Some(500)
        }
        .resolve(Some(&palette)),
        Some(0xFFB2_B2_B2)
    );

    // -0.5 (darken): 100 * (1-0.5) = 50
    assert_eq!(
        Color::Theme {
            theme: theme_idx,
            tint: Some(-500)
        }
        .resolve(Some(&palette)),
        Some(0xFF323232)
    );

    assert_eq!(
        Color::Theme {
            theme: theme_idx,
            tint: Some(1000)
        }
        .resolve(Some(&palette)),
        Some(0xFFFF_FFFF)
    );

    assert_eq!(
        Color::Theme {
            theme: theme_idx,
            tint: Some(-1000)
        }
        .resolve(Some(&palette)),
        Some(0xFF00_00_00)
    );

    // Clamp outside the OOXML (-1..=1) range.
    assert_eq!(
        Color::Theme {
            theme: theme_idx,
            tint: Some(2000)
        }
        .resolve(Some(&palette)),
        Some(0xFFFF_FFFF)
    );
    assert_eq!(
        Color::Theme {
            theme: theme_idx,
            tint: Some(-2000)
        }
        .resolve(Some(&palette)),
        Some(0xFF00_00_00)
    );
}

#[test]
fn indexed_palette_matches_reference_table() {
    // ECMA-376 / SpreadsheetML default indexedColors palette (0..=63), encoded as ARGB.
    const EXPECTED: [u32; 64] = [
        0xFF000000, 0xFFFFFFFF, 0xFFFF0000, 0xFF00FF00, 0xFF0000FF, 0xFFFFFF00, 0xFFFF00FF,
        0xFF00FFFF, 0xFF000000, 0xFFFFFFFF, 0xFFFF0000, 0xFF00FF00, 0xFF0000FF, 0xFFFFFF00,
        0xFFFF00FF, 0xFF00FFFF, 0xFF800000, 0xFF008000, 0xFF000080, 0xFF808000, 0xFF800080,
        0xFF008080, 0xFFC0C0C0, 0xFF808080, 0xFF9999FF, 0xFF993366, 0xFFFFFFCC, 0xFFCCFFFF,
        0xFF660066, 0xFFFF8080, 0xFF0066CC, 0xFFCCCCFF, 0xFF000080, 0xFFFF00FF, 0xFFFFFF00,
        0xFF00FFFF, 0xFF800080, 0xFF800000, 0xFF008080, 0xFF0000FF, 0xFF00CCFF, 0xFFCCFFFF,
        0xFFCCFFCC, 0xFFFFFF99, 0xFF99CCFF, 0xFFFF99CC, 0xFFCC99FF, 0xFFFFCC99, 0xFF3366FF,
        0xFF33CCCC, 0xFF99CC00, 0xFFFFCC00, 0xFFFF9900, 0xFFFF6600, 0xFF666699, 0xFF969696,
        0xFF003366, 0xFF339966, 0xFF003300, 0xFF333300, 0xFF993300, 0xFF993366, 0xFF333399,
        0xFF333333,
    ];

    for (idx, expected) in EXPECTED.iter().enumerate() {
        assert_eq!(
            indexed_color_argb(idx as u16),
            Some(*expected),
            "index {idx}"
        );
    }

    // 64+ are not part of the default palette.
    assert_eq!(indexed_color_argb(64), None);

    // But Excel treats indexed=64 as "automatic" (context-dependent).
    assert_eq!(
        Color::Indexed(64).resolve_in_context(None, ColorContext::Font),
        Some(0xFF000000)
    );
}

#[test]
fn theme_slot_mapping_is_stable() {
    let palette = ThemePalette {
        lt1: formula_model::ArgbColor(0xFF000001),
        dk1: formula_model::ArgbColor(0xFF000002),
        lt2: formula_model::ArgbColor(0xFF000003),
        dk2: formula_model::ArgbColor(0xFF000004),
        accent1: formula_model::ArgbColor(0xFF000005),
        accent2: formula_model::ArgbColor(0xFF000006),
        accent3: formula_model::ArgbColor(0xFF000007),
        accent4: formula_model::ArgbColor(0xFF000008),
        accent5: formula_model::ArgbColor(0xFF000009),
        accent6: formula_model::ArgbColor(0xFF00000A),
        hlink: formula_model::ArgbColor(0xFF00000B),
        fol_hlink: formula_model::ArgbColor(0xFF00000C),
    };

    let expected = [
        palette.lt1.0,
        palette.dk1.0,
        palette.lt2.0,
        palette.dk2.0,
        palette.accent1.0,
        palette.accent2.0,
        palette.accent3.0,
        palette.accent4.0,
        palette.accent5.0,
        palette.accent6.0,
        palette.hlink.0,
        palette.fol_hlink.0,
    ];

    for (idx, expected) in expected.iter().enumerate() {
        assert_eq!(
            Color::Theme {
                theme: idx as u16,
                tint: None
            }
            .resolve(Some(&palette)),
            Some(*expected),
            "theme index {idx}"
        );
    }

    assert!(ThemeColorSlot::from_theme_index(12).is_none());
    assert_eq!(
        Color::Theme {
            theme: 12,
            tint: None
        }
        .resolve(Some(&palette)),
        None
    );
}

#[test]
fn theme_palette_roundtrips_via_workbook_fixture() {
    let palettes: BTreeMap<String, ThemePalette> =
        serde_json::from_str(include_str!("fixtures/theme_palettes.json")).unwrap();

    for (name, palette) in palettes {
        let mut wb = Workbook::new();
        wb.theme = palette.clone();

        let json = serde_json::to_string(&wb).unwrap();
        let reparsed: Workbook = serde_json::from_str(&json).unwrap();
        assert_eq!(reparsed.theme, palette, "palette {name}");
    }
}
