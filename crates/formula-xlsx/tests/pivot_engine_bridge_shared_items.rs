use chrono::NaiveDate;
use formula_engine::pivot::PivotValue;
use formula_xlsx::pivots::engine_bridge::pivot_cache_to_engine_source;
use formula_xlsx::pivots::{PivotCacheDefinition, PivotCacheField, PivotCacheValue};

use pretty_assertions::assert_eq;

#[test]
fn resolves_shared_item_indices_and_dates() {
    let def = PivotCacheDefinition {
        cache_fields: vec![
            PivotCacheField {
                name: "Name".to_string(),
                shared_items: Some(vec![
                    PivotCacheValue::String("Alice".to_string()),
                    PivotCacheValue::String("Bob".to_string()),
                ]),
                ..Default::default()
            },
            PivotCacheField {
                name: "Amount".to_string(),
                shared_items: Some(vec![
                    PivotCacheValue::Number(10.0),
                    PivotCacheValue::Number(20.0),
                ]),
                ..Default::default()
            },
            PivotCacheField {
                name: "When".to_string(),
                shared_items: Some(vec![PivotCacheValue::DateTime(
                    "2024-01-15T00:00:00Z".to_string(),
                )]),
                ..Default::default()
            },
        ],
        ..Default::default()
    };

    let records = vec![
        // All values use shared item indices.
        vec![
            PivotCacheValue::Index(0),
            PivotCacheValue::Index(1),
            PivotCacheValue::Index(0),
        ],
        // Mixed record: third value uses a `<d v="..."/>` datetime string.
        vec![
            PivotCacheValue::Index(1),
            PivotCacheValue::Index(0),
            PivotCacheValue::DateTime("2024-01-15T00:00:00Z".to_string()),
        ],
        // Out-of-range index should behave like blank.
        vec![
            PivotCacheValue::Index(42),
            PivotCacheValue::Index(0),
            PivotCacheValue::Index(0),
        ],
    ];

    let source = pivot_cache_to_engine_source(&def, records.into_iter());

    assert_eq!(
        source[0],
        vec![
            PivotValue::Text("Name".to_string()),
            PivotValue::Text("Amount".to_string()),
            PivotValue::Text("When".to_string()),
        ]
    );

    let date = NaiveDate::from_ymd_opt(2024, 1, 15).expect("valid date");

    assert_eq!(
        source[1],
        vec![
            PivotValue::Text("Alice".to_string()),
            PivotValue::Number(20.0),
            PivotValue::Date(date),
        ]
    );

    // `<d v="2024-01-15T00:00:00Z"/>` should map to `PivotValue::Date(2024-01-15)`.
    assert_eq!(source[2][2], PivotValue::Date(date));

    // Unresolvable shared-item indices are treated as blank.
    assert_eq!(source[3][0], PivotValue::Blank);
}
