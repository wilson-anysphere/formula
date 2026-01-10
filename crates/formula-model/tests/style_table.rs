use formula_model::{Color, Font, Style, StyleTable};

#[test]
fn style_table_intern_deduplicates() {
    let mut table = StyleTable::new();

    let style = Style {
        font: Some(Font {
            bold: true,
            color: Some(Color::new_argb(0xFFFF0000)),
            ..Default::default()
        }),
        number_format: Some("0%".to_string()),
        ..Default::default()
    };

    let a = table.intern(style.clone());
    let b = table.intern(style);
    assert_eq!(a, b, "identical styles should reuse the same id");
}
