export function mountScriptEditorPanel(params: {
  workbook: any;
  container: HTMLElement;
  monaco?: any;
  isEditing?: () => boolean;
  isReadOnly?: () => boolean;
}): { dispose: () => void };
