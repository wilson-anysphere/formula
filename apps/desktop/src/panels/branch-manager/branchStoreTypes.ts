import type { CollabSession } from "@formula/collab-session";

export type MaybePromise<T> = T | Promise<T>;

/**
 * Minimal interface implemented by BranchService stores.
 *
 * The canonical JS typedef lives alongside the branch store implementations in
 * `packages/versioning/branches`, but that package does not currently export a
 * first-class TS type. We define the structural contract here so desktop
 * deployments can cleanly provide out-of-doc stores without writing to reserved
 * Yjs roots (`branching:*`).
 */
export interface BranchStore {
  ensureDocument(docId: string, actor: any, initialState: any): Promise<void>;
  hasDocument?(docId: string): Promise<boolean>;
  listBranches(docId: string): Promise<any[]>;
  getBranch(docId: string, name: string): Promise<any | null>;
  createBranch(input: {
    docId: string;
    name: string;
    createdBy: string;
    createdAt: number;
    description: string | null;
    headCommitId: string;
  }): Promise<any>;
  renameBranch(docId: string, oldName: string, newName: string): Promise<void>;
  deleteBranch(docId: string, name: string): Promise<void>;
  updateBranchHead(docId: string, name: string, headCommitId: string): Promise<void>;
  getCurrentBranchName?(docId: string): Promise<string>;
  setCurrentBranchName?(docId: string, name: string): Promise<void>;
  createCommit(input: {
    docId: string;
    parentCommitId: string | null;
    mergeParentCommitId: string | null;
    createdBy: string;
    createdAt: number;
    message: string | null;
    patch: any;
    nextState?: any;
  }): Promise<any>;
  getCommit(commitId: string): Promise<any | null>;
  getDocumentStateAtCommit(commitId: string): Promise<any>;
}

/**
 * Factory hook for providing an alternate branch store implementation (e.g.
 * SQLiteBranchStore, API-backed store) that does not write to reserved Yjs roots.
 */
export type CreateBranchStore = (session: CollabSession) => MaybePromise<BranchStore>;

