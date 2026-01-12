use formula_model::drawings::ImageId;
use formula_model::{CellValue, ImageValue};

#[test]
fn cell_value_image_json_roundtrip_is_stable() {
    let value = CellValue::Image(ImageValue {
        image_id: ImageId::new("image1.png"),
        alt_text: Some("Logo".to_string()),
        width: Some(128),
        height: Some(64),
    });

    let json = serde_json::to_value(&value).expect("serialize CellValue::Image");
    assert_eq!(
        json,
        serde_json::json!({
            "type": "image",
            "value": {
                "imageId": "image1.png",
                "altText": "Logo",
                "width": 128,
                "height": 64
            }
        })
    );

    let roundtrip: CellValue = serde_json::from_value(json).expect("deserialize CellValue::Image");
    assert_eq!(roundtrip, value);
}

#[test]
fn cell_value_image_missing_optional_fields_default_to_none() {
    let json = serde_json::json!({
        "type": "image",
        "value": {
            "imageId": "image1.png"
        }
    });

    let parsed: CellValue = serde_json::from_value(json).expect("deserialize minimal image value");
    assert_eq!(
        parsed,
        CellValue::Image(ImageValue {
            image_id: ImageId::new("image1.png"),
            alt_text: None,
            width: None,
            height: None,
        })
    );
}

#[test]
fn cell_value_image_deserializes_snake_case_aliases() {
    // Some legacy IPC payloads used snake_case keys.
    let json = serde_json::json!({
        "type": "image",
        "value": {
            "image_id": "image1.png",
            "alt_text": "Logo",
            "width": 128,
            "height": 64
        }
    });

    let parsed: CellValue =
        serde_json::from_value(json).expect("deserialize snake_case image value");
    assert_eq!(
        parsed,
        CellValue::Image(ImageValue {
            image_id: ImageId::new("image1.png"),
            alt_text: Some("Logo".to_string()),
            width: Some(128),
            height: Some(64),
        })
    );
}
