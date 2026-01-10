use serde::{Deserialize, Serialize};

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
