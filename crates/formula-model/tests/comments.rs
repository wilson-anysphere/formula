use formula_model::{CellRef, Comment, CommentKind, CommentPatch, Range, Reply, Worksheet};

#[test]
fn create_note_and_threaded_comments_on_same_cell() {
    let mut sheet = Worksheet::new(1, "Sheet1");
    let cell = CellRef::new(0, 0);

    let note = Comment {
        kind: CommentKind::Note,
        content: "legacy note".to_string(),
        ..Default::default()
    };
    let threaded = Comment {
        kind: CommentKind::Threaded,
        content: "threaded comment".to_string(),
        ..Default::default()
    };

    let note_id = sheet.add_comment(cell, note).unwrap();
    let threaded_id = sheet.add_comment(cell, threaded).unwrap();

    let comments = sheet.comments_for_cell(cell);
    assert_eq!(comments.len(), 2);
    assert_eq!(comments[0].id, note_id);
    assert_eq!(comments[0].kind, CommentKind::Note);
    assert_eq!(comments[1].id, threaded_id);
    assert_eq!(comments[1].kind, CommentKind::Threaded);
}

#[test]
fn reply_add_and_remove() {
    let mut sheet = Worksheet::new(1, "Sheet1");
    let cell = CellRef::new(0, 0);

    let comment_id = sheet
        .add_comment(
            cell,
            Comment {
                content: "parent".to_string(),
                ..Default::default()
            },
        )
        .unwrap();

    let reply_id = sheet
        .add_reply(
            &comment_id,
            Reply {
                content: "reply".to_string(),
                ..Default::default()
            },
        )
        .unwrap();

    let comments = sheet.comments_for_cell(cell);
    assert_eq!(comments.len(), 1);
    assert_eq!(comments[0].replies.len(), 1);
    assert_eq!(comments[0].replies[0].id, reply_id);

    sheet.delete_reply(&reply_id).unwrap();
    let comments = sheet.comments_for_cell(cell);
    assert_eq!(comments[0].replies.len(), 0);
}

#[test]
fn merge_cell_anchor_normalization() {
    let mut sheet = Worksheet::new(1, "Sheet1");
    sheet.merge_range(Range::from_a1("A1:B2").unwrap()).unwrap();

    // B2 is inside the merged region anchored at A1.
    let merged_cell = CellRef::new(1, 1);
    let anchor = CellRef::new(0, 0);

    let comment_id = sheet
        .add_comment(
            merged_cell,
            Comment {
                content: "anchored".to_string(),
                ..Default::default()
            },
        )
        .unwrap();

    // Both A1 and B2 should resolve to the same anchored comments.
    assert_eq!(sheet.comments_for_cell(anchor).len(), 1);
    assert_eq!(sheet.comments_for_cell(merged_cell).len(), 1);
    assert_eq!(sheet.comments_for_cell(anchor)[0].id, comment_id);
    assert_eq!(sheet.comments_for_cell(anchor)[0].cell_ref, anchor);
}

#[test]
fn comment_updates_and_serde_roundtrip() {
    let mut sheet = Worksheet::new(1, "Sheet1");
    let cell = CellRef::new(0, 0);

    let comment_id = sheet
        .add_comment(
            cell,
            Comment {
                content: "before".to_string(),
                ..Default::default()
            },
        )
        .unwrap();

    sheet
        .update_comment(
            &comment_id,
            CommentPatch {
                content: Some("after".to_string()),
                resolved: Some(true),
                ..Default::default()
            },
        )
        .unwrap();

    let json = serde_json::to_value(&sheet).unwrap();
    let roundtrip: Worksheet = serde_json::from_value(json).unwrap();

    let comments = roundtrip.comments_for_cell(cell);
    assert_eq!(comments.len(), 1);
    assert_eq!(comments[0].id, comment_id);
    assert_eq!(comments[0].content, "after");
    assert!(comments[0].resolved);
}
