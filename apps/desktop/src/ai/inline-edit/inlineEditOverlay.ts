import type { ToolPlanPreview } from "../../../../../packages/ai-tools/src/preview/preview-engine.js";

export interface InlineEditOverlayCallbacks {
  onCancel: () => void;
  onRun: (prompt: string) => void;
}

type OverlayMode = "prompt" | "running" | "preview" | "error";

export class InlineEditOverlay {
  readonly element: HTMLDivElement;
  private readonly selectionLabel: HTMLElement;
  private readonly promptInput: HTMLInputElement;
  private readonly statusLabel: HTMLElement;
  private readonly errorLabel: HTMLElement;

  private readonly runButton: HTMLButtonElement;
  private readonly cancelButton: HTMLButtonElement;

  private readonly previewContainer: HTMLDivElement;
  private readonly previewSummary: HTMLElement;
  private readonly previewList: HTMLUListElement;
  private readonly approveButton: HTMLButtonElement;
  private readonly previewCancelButton: HTMLButtonElement;

  private mode: OverlayMode = "prompt";
  private callbacks: InlineEditOverlayCallbacks | null = null;
  private approvalResolver: ((approved: boolean) => void) | null = null;

  constructor(private readonly parent: HTMLElement) {
    const root = document.createElement("div");
    root.className = "ai-inline-edit-overlay dialog";
    root.hidden = true;
    root.dataset.testid = "inline-edit-overlay";
    root.dataset.keybindingBarrier = "true";
    root.setAttribute("role", "dialog");
    root.setAttribute("aria-label", "AI inline edit");

    // Stop key events from reaching the grid while the overlay is open.
    root.addEventListener(
      "keydown",
      (e) => {
        if (!this.isOpen()) return;
        e.stopPropagation();

        if (e.key === "Escape") {
          e.preventDefault();
          this.handleCancel();
          return;
        }

        if (e.key === "Enter") {
          if (this.mode === "prompt") {
            e.preventDefault();
            this.handleRun();
            return;
          }
          if (this.mode === "preview") {
            e.preventDefault();
            this.handleApprove();
          }
        }
      },
      true
    );

    const title = document.createElement("div");
    title.className = "ai-inline-edit-title dialog__title";
    title.textContent = "AI Inline Edit";

    const selection = document.createElement("div");
    selection.className = "ai-inline-edit-selection";
    selection.dataset.testid = "inline-edit-selection";

    const promptRow = document.createElement("div");
    promptRow.className = "ai-inline-edit-prompt-row";

    const promptInput = document.createElement("input");
    promptInput.type = "text";
    promptInput.placeholder = "Describe the transformation…";
    promptInput.className = "ai-inline-edit-prompt";
    promptInput.dataset.testid = "inline-edit-prompt";
    promptInput.autocomplete = "off";
    promptInput.spellcheck = false;
    promptInput.addEventListener("keydown", (e) => {
      // The container listener handles Enter/Escape; prevent duplicate handling here.
      if (e.key === "Enter" || e.key === "Escape") e.stopPropagation();
    });

    promptRow.appendChild(promptInput);

    const status = document.createElement("div");
    status.className = "ai-inline-edit-status";
    status.dataset.testid = "inline-edit-status";

    const error = document.createElement("div");
    error.className = "ai-inline-edit-error";
    error.dataset.testid = "inline-edit-error";
    error.hidden = true;

    const actions = document.createElement("div");
    actions.className = "ai-inline-edit-actions";

    const runButton = document.createElement("button");
    runButton.type = "button";
    runButton.textContent = "Run";
    runButton.dataset.testid = "inline-edit-run";
    runButton.addEventListener("click", () => this.handleRun());

    const cancelButton = document.createElement("button");
    cancelButton.type = "button";
    cancelButton.textContent = "Cancel";
    cancelButton.dataset.testid = "inline-edit-cancel";
    cancelButton.addEventListener("click", () => this.handleCancel());

    actions.appendChild(runButton);
    actions.appendChild(cancelButton);

    const preview = document.createElement("div");
    preview.className = "ai-inline-edit-preview";
    preview.dataset.testid = "inline-edit-preview";
    preview.hidden = true;

    const previewSummary = document.createElement("div");
    previewSummary.className = "ai-inline-edit-preview-summary";
    previewSummary.dataset.testid = "inline-edit-preview-summary";

    const previewList = document.createElement("ul");
    previewList.className = "ai-inline-edit-preview-list";
    previewList.dataset.testid = "inline-edit-preview-changes";

    const previewActions = document.createElement("div");
    previewActions.className = "ai-inline-edit-preview-actions";

    const approveButton = document.createElement("button");
    approveButton.type = "button";
    approveButton.textContent = "Apply";
    approveButton.dataset.testid = "inline-edit-approve";
    approveButton.addEventListener("click", () => this.handleApprove());

    const previewCancel = document.createElement("button");
    previewCancel.type = "button";
    previewCancel.textContent = "Cancel";
    previewCancel.dataset.testid = "inline-edit-preview-cancel";
    previewCancel.addEventListener("click", () => this.handleCancel());

    previewActions.appendChild(approveButton);
    previewActions.appendChild(previewCancel);

    preview.appendChild(previewSummary);
    preview.appendChild(previewList);
    preview.appendChild(previewActions);

    root.appendChild(title);
    root.appendChild(selection);
    root.appendChild(promptRow);
    root.appendChild(status);
    root.appendChild(error);
    root.appendChild(actions);
    root.appendChild(preview);

    this.element = root;
    this.selectionLabel = selection;
    this.promptInput = promptInput;
    this.statusLabel = status;
    this.errorLabel = error;
    this.runButton = runButton;
    this.cancelButton = cancelButton;
    this.previewContainer = preview;
    this.previewSummary = previewSummary;
    this.previewList = previewList;
    this.approveButton = approveButton;
    this.previewCancelButton = previewCancel;

    this.parent.appendChild(root);
  }

  isOpen(): boolean {
    return !this.element.hidden;
  }

  open(selection: string, callbacks: InlineEditOverlayCallbacks): void {
    this.callbacks = callbacks;
    this.selectionLabel.textContent = selection;
    this.promptInput.value = "";
    this.errorLabel.textContent = "";
    this.errorLabel.hidden = true;
    this.statusLabel.textContent = "";
    this.previewContainer.hidden = true;
    this.runButton.disabled = false;
    this.cancelButton.disabled = false;
    this.mode = "prompt";
    this.element.hidden = false;
    this.promptInput.disabled = false;
    this.promptInput.focus();
  }

  close(): void {
    if (this.approvalResolver) {
      // Ensure pending approval prompts never hang if the overlay is closed externally
      // (e.g. user cancels while the tool loop is still running).
      const resolve = this.approvalResolver;
      this.approvalResolver = null;
      resolve(false);
    }
    this.callbacks = null;
    this.mode = "prompt";
    this.element.hidden = true;
    this.promptInput.value = "";
    this.previewContainer.hidden = true;
  }

  getPrompt(): string {
    return this.promptInput.value;
  }

  setRunning(message: string): void {
    this.mode = "running";
    this.statusLabel.textContent = message;
    this.errorLabel.hidden = true;
    this.promptInput.disabled = true;
    this.runButton.disabled = true;
    this.cancelButton.disabled = false;
  }

  showError(message: string): void {
    this.mode = "error";
    this.statusLabel.textContent = "";
    this.errorLabel.textContent = message;
    this.errorLabel.hidden = false;
    this.promptInput.disabled = false;
    this.runButton.disabled = false;
    this.cancelButton.disabled = false;
    this.previewContainer.hidden = true;
    this.promptInput.focus();
  }

  async requestApproval(preview: ToolPlanPreview): Promise<boolean> {
    // If the overlay is not visible, we can't surface an approval UI.
    // Fail closed (deny) to avoid hanging the tool loop.
    if (!this.isOpen()) return false;

    if (this.approvalResolver) {
      // Shouldn't happen in normal usage, but ensure we never leave a pending promise.
      this.approvalResolver(false);
      this.approvalResolver = null;
    }

    this.mode = "preview";
    this.statusLabel.textContent = "Preview changes and approve to apply.";
    this.errorLabel.hidden = true;

    this.previewSummary.textContent = formatPreviewSummary(preview);
    this.previewList.replaceChildren();
    for (const change of preview.changes) {
      const item = document.createElement("li");
      item.textContent = `${change.cell}: ${formatCellData(change.before)} → ${formatCellData(change.after)}`;
      this.previewList.appendChild(item);
    }

    this.previewContainer.hidden = false;
    this.approveButton.disabled = false;
    this.previewCancelButton.disabled = false;
    this.approveButton.focus();

    return new Promise<boolean>((resolve) => {
      this.approvalResolver = resolve;
    });
  }

  private handleRun(): void {
    if (!this.callbacks) return;
    if (this.mode !== "prompt" && this.mode !== "error") return;
    const prompt = this.promptInput.value.trim();
    if (!prompt) return;
    this.callbacks.onRun(prompt);
  }

  private handleApprove(): void {
    if (this.mode !== "preview") return;
    if (!this.approvalResolver) return;
    this.approveButton.disabled = true;
    this.previewCancelButton.disabled = true;
    const resolve = this.approvalResolver;
    this.approvalResolver = null;
    resolve(true);
  }

  private handleCancel(): void {
    if (this.mode === "preview" && this.approvalResolver) {
      const resolve = this.approvalResolver;
      this.approvalResolver = null;
      resolve(false);
      return;
    }
    this.callbacks?.onCancel();
  }
}

function formatPreviewSummary(preview: ToolPlanPreview): string {
  const { creates, modifies, deletes, total_changes } = preview.summary;
  return `Changes: ${total_changes} (creates ${creates}, modifies ${modifies}, deletes ${deletes})`;
}

function formatCellData(cell: { value?: unknown; formula?: string | null } | null | undefined): string {
  if (!cell) return "∅";
  if (cell.formula) return String(cell.formula);
  if (cell.value == null) return "∅";
  if (typeof cell.value === "string") return JSON.stringify(cell.value);
  try {
    return JSON.stringify(cell.value);
  } catch {
    return String(cell.value);
  }
}
