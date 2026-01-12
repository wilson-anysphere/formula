export type MessageType = "info" | "warning" | "error";

export function showToast(message: string, type?: MessageType, options?: { timeoutMs?: number }): void;

export type InputBoxOptions = {
  prompt?: string;
  value?: string;
  placeHolder?: string;
  type?: "text" | "password" | "textarea";
  rows?: number;
  okLabel?: string;
  cancelLabel?: string;
};

export function showInputBox(options?: InputBoxOptions): Promise<string | null>;

export type QuickPickItem<T> = {
  label: string;
  value: T;
  description?: string;
  detail?: string;
};

export function showQuickPick<T>(items: QuickPickItem<T>[], options?: { placeHolder?: string }): Promise<T | null>;

