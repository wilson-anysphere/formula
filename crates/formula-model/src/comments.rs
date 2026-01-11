use serde::{Deserialize, Serialize};
use thiserror::Error;

use crate::CellRef;

pub type TimestampMs = i64;

fn default_cell_ref() -> CellRef {
    CellRef::new(0, 0)
}

#[derive(Clone, Debug, Default, PartialEq, Eq, Serialize, Deserialize)]
pub struct CommentAuthor {
    #[serde(default)]
    pub id: String,
    #[serde(default)]
    pub name: String,
}

#[derive(Clone, Debug, Default, PartialEq, Eq, Serialize, Deserialize)]
pub struct Mention {
    #[serde(default)]
    pub user_id: String,
    #[serde(default)]
    pub display: String,
}

#[derive(Clone, Debug, Default, PartialEq, Eq, Serialize, Deserialize)]
pub struct Reply {
    #[serde(default)]
    pub id: String,
    #[serde(default)]
    pub author: CommentAuthor,
    #[serde(default)]
    pub created_at: TimestampMs,
    #[serde(default)]
    pub updated_at: TimestampMs,
    #[serde(default)]
    pub content: String,
    #[serde(default)]
    pub mentions: Vec<Mention>,
}

#[derive(Copy, Clone, Debug, PartialEq, Eq, Serialize, Deserialize)]
pub enum CommentKind {
    Note,
    Threaded,
}

impl Default for CommentKind {
    fn default() -> Self {
        CommentKind::Threaded
    }
}

#[derive(Clone, Debug, PartialEq, Eq, Serialize, Deserialize)]
pub struct Comment {
    #[serde(default)]
    pub id: String,
    #[serde(default = "default_cell_ref")]
    pub cell_ref: CellRef,
    #[serde(default)]
    pub author: CommentAuthor,
    #[serde(default)]
    pub created_at: TimestampMs,
    #[serde(default)]
    pub updated_at: TimestampMs,
    #[serde(default)]
    pub resolved: bool,
    #[serde(default)]
    pub kind: CommentKind,
    #[serde(default)]
    pub content: String,
    #[serde(default)]
    pub mentions: Vec<Mention>,
    #[serde(default)]
    pub replies: Vec<Reply>,
}

/// Partial update payload for editing an existing [`Comment`].
#[derive(Clone, Debug, Default, PartialEq, Eq, Serialize, Deserialize)]
pub struct CommentPatch {
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub author: Option<CommentAuthor>,
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub updated_at: Option<TimestampMs>,
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub resolved: Option<bool>,
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub kind: Option<CommentKind>,
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub content: Option<String>,
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub mentions: Option<Vec<Mention>>,
}

#[derive(Clone, Debug, PartialEq, Eq, Error)]
pub enum CommentError {
    #[error("comment not found: {0}")]
    CommentNotFound(String),
    #[error("reply not found: {0}")]
    ReplyNotFound(String),
    #[error("duplicate comment id: {0}")]
    DuplicateCommentId(String),
    #[error("duplicate reply id: {0}")]
    DuplicateReplyId(String),
}

impl Default for Comment {
    fn default() -> Self {
        Self {
            id: String::new(),
            cell_ref: default_cell_ref(),
            author: CommentAuthor::default(),
            created_at: 0,
            updated_at: 0,
            resolved: false,
            kind: CommentKind::default(),
            content: String::new(),
            mentions: Vec::new(),
            replies: Vec::new(),
        }
    }
}
