export type { CommentsRoot, YCommentsArray, YCommentsMap } from "./manager.ts";
export {
  createYComment,
  createYReply,
  getCommentsMap,
  getCommentsRoot,
  migrateCommentsArrayToMap,
  yCommentToComment,
  yReplyToReply,
} from "./manager.ts";
