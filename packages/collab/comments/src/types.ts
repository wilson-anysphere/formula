export type CommentKind = "note" | "threaded";

export type TimestampMs = number;

export interface CommentAuthor {
  id: string;
  name: string;
}

export interface Mention {
  userId: string;
  display: string;
}

export interface Reply {
  id: string;
  author: CommentAuthor;
  createdAt: TimestampMs;
  updatedAt: TimestampMs;
  content: string;
  mentions: Mention[];
}

export interface Comment {
  id: string;
  cellRef: string;
  kind: CommentKind;
  author: CommentAuthor;
  createdAt: TimestampMs;
  updatedAt: TimestampMs;
  resolved: boolean;
  content: string;
  mentions: Mention[];
  replies: Reply[];
}

