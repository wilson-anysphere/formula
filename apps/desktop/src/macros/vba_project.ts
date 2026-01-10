export type VbaReferenceSummary = {
  name: string | null;
  guid: string | null;
  major: number | null;
  minor: number | null;
  path: string | null;
  raw: string;
};

export type VbaModuleSummary = {
  name: string;
  moduleType: string;
  code: string;
};

export type VbaProjectSummary = {
  name: string | null;
  constants: string | null;
  references: VbaReferenceSummary[];
  modules: VbaModuleSummary[];
};

type TauriInvoke = (cmd: string, args?: any) => Promise<any>;

function getTauriInvoke(): TauriInvoke {
  const invoke = (globalThis as any).__TAURI__?.core?.invoke as TauriInvoke | undefined;
  if (!invoke) {
    throw new Error("Tauri invoke API not available");
  }
  return invoke;
}

function normalizeReference(raw: any): VbaReferenceSummary {
  return {
    name: raw?.name ?? null,
    guid: raw?.guid ?? null,
    major: typeof raw?.major === "number" ? raw.major : raw?.major != null ? Number(raw.major) : null,
    minor: typeof raw?.minor === "number" ? raw.minor : raw?.minor != null ? Number(raw.minor) : null,
    path: raw?.path ?? null,
    raw: String(raw?.raw ?? ""),
  };
}

function normalizeModule(raw: any): VbaModuleSummary {
  return {
    name: String(raw?.name ?? ""),
    moduleType: String(raw?.module_type ?? raw?.moduleType ?? ""),
    code: String(raw?.code ?? ""),
  };
}

/**
 * Fetch the parsed VBA project (modules + code) for the active workbook.
 *
 * The backend returns `null` if the workbook has no `xl/vbaProject.bin`.
 */
export async function getVbaProject(workbookId: string): Promise<VbaProjectSummary | null> {
  const invoke = getTauriInvoke();
  const raw = await invoke("get_vba_project", { workbook_id: workbookId });
  if (!raw) return null;

  return {
    name: raw?.name ?? null,
    constants: raw?.constants ?? null,
    references: Array.isArray(raw?.references) ? raw.references.map(normalizeReference) : [],
    modules: Array.isArray(raw?.modules) ? raw.modules.map(normalizeModule) : [],
  };
}

