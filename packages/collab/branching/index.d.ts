import type { CollabSession } from "@formula/collab-session";

export type Actor = any;
export type BranchService = any;
export type ConflictResolution = any;

export declare class CollabBranchingWorkflow {
  constructor(input: {
    session: CollabSession;
    branchService: BranchService;
    rootName?: string;
    applyOrigin?: any;
    applyWithSessionOrigin?: boolean;
  });

  getCurrentBranchName(): string;
  getCurrentBranchNameAsync(): Promise<string>;

  listBranches(): Promise<any>;
  createBranch(actor: Actor, input: { name: string; description?: string }): Promise<any>;
  renameBranch(actor: Actor, input: { oldName: string; newName: string }): Promise<void>;
  deleteBranch(actor: Actor, input: { name: string }): Promise<void>;

  getCurrentBranch(): Promise<any>;
  getCurrentState(): Promise<any>;
  commitCurrentState(actor: Actor, message?: string): Promise<any>;

  checkoutBranch(actor: Actor, input: { name: string }): Promise<any>;
  previewMerge(actor: Actor, input: { sourceBranch: string }): Promise<any>;
  merge(actor: Actor, input: { sourceBranch: string; resolutions: ConflictResolution[]; message?: string }): Promise<any>;
}

