import type {
  MacroBackend,
  MacroCellUpdate,
  MacroInfo,
  MacroPermission,
  MacroRunResult,
  MacroSecurityStatus,
  MacroSignatureStatus,
} from "./types";
import type { MacroSecurityController } from "./security";

export interface RunMacroOptions {
  workbookId: string;
  macroId: string;
  timeoutMs?: number;
}

export class MacroRunner {
  constructor(
    private readonly backend: MacroBackend,
    private readonly security: MacroSecurityController
  ) {}

  async list(workbookId: string): Promise<MacroInfo[]> {
    return await this.backend.listMacros(workbookId);
  }

  async run(options: RunMacroOptions): Promise<MacroRunResult> {
    let status: MacroSecurityStatus;
    try {
      status = await this.backend.getMacroSecurityStatus(options.workbookId);
    } catch (err) {
      return { ok: false, output: [], error: { message: `Failed to read macro security status: ${String(err)}` } };
    }

    if (status.hasMacros && !macroTrustAllowsRun(status)) {
      const decision = await this.security.requestTrustDecision({
        workbookId: options.workbookId,
        macroId: options.macroId,
        status,
      });

      if (!decision || decision === "blocked") {
        return { ok: false, output: [], error: { message: "User declined to trust macros." } };
      }

      try {
        status = await this.backend.setMacroTrust(options.workbookId, decision);
      } catch (err) {
        return { ok: false, output: [], error: { message: `Failed to update Trust Center decision: ${String(err)}` } };
      }
    }

    const maxPermissionEscalations = 2;
    let permissions: MacroPermission[] = [];
    let mergedOutput: string[] = [];
    let mergedUpdates: MacroCellUpdate[] | undefined;

    for (let escalation = 0; escalation <= maxPermissionEscalations; escalation++) {
      const result = await this.backend.runMacro({
        workbookId: options.workbookId,
        macroId: options.macroId,
        permissions,
        timeoutMs: options.timeoutMs,
      });

      mergedOutput.push(...result.output);
      mergedUpdates = mergeUpdates(mergedUpdates, result.updates);

      if (result.ok) {
        return { ...result, output: mergedOutput, updates: mergedUpdates };
      }

      if (!result.permissionRequest) {
        return { ...result, output: mergedOutput, updates: mergedUpdates };
      }

      if (escalation === maxPermissionEscalations) {
        return {
          ...result,
          output: mergedOutput,
          updates: mergedUpdates,
          error: result.error ?? { message: "Macro requested additional permissions too many times." },
        };
      }

      const granted = await this.security.requestPermissions({
        workbookId: options.workbookId,
        request: result.permissionRequest,
        alreadyGranted: permissions,
      });

      if (!granted || granted.length === 0) {
        return {
          ...result,
          output: mergedOutput,
          updates: mergedUpdates,
          error: { message: "User declined to grant requested permissions." },
        };
      }

      const next = mergePermissions(permissions, granted);
      const newlyGranted = next.filter((p) => !permissions.includes(p));
      permissions = next;

      if (newlyGranted.length > 0) {
        mergedOutput.push(`[macro] Granted permissions: ${newlyGranted.join(", ")}. Retrying...`);
      } else {
        mergedOutput.push(`[macro] Retrying with existing permissions...`);
      }
    }

    // Unreachable: loop returns in all branches.
    return { ok: false, output: [], error: { message: "Macro runner exhausted retries." } };
  }
}

function isCryptographicallyVerifiedSignature(status: MacroSignatureStatus | undefined): boolean {
  if (!status) return false;
  // `signed_untrusted` is reserved for a future "valid signature but untrusted chain" state.
  // It should still satisfy the "signed only" policy because the signature verifies.
  return status === "signed_verified" || status === "signed_untrusted";
}

function macroTrustAllowsRun(status: MacroSecurityStatus): boolean {
  switch (status.trust) {
    case "trusted_always":
    case "trusted_once":
      return true;
    case "trusted_signed_only":
      return isCryptographicallyVerifiedSignature(status.signature?.status);
    case "blocked":
    default:
      return false;
  }
}

function mergePermissions(existing: MacroPermission[], granted: MacroPermission[]): MacroPermission[] {
  const out = new Set<MacroPermission>(existing);
  for (const perm of granted) out.add(perm);
  return Array.from(out.values());
}

function mergeUpdates(
  existing: MacroCellUpdate[] | undefined,
  next: MacroCellUpdate[] | undefined
): MacroCellUpdate[] | undefined {
  if (!next || next.length === 0) return existing;
  if (!existing || existing.length === 0) return [...next];

  const out = [...existing];
  const index = new Map<string, number>();
  for (let i = 0; i < out.length; i++) {
    const u = out[i];
    index.set(`${u.sheetId}:${u.row}:${u.col}`, i);
  }
  for (const u of next) {
    const key = `${u.sheetId}:${u.row}:${u.col}`;
    const idx = index.get(key);
    if (idx == null) {
      index.set(key, out.length);
      out.push(u);
    } else {
      out[idx] = u;
    }
  }
  return out;
}
