import { describe, expect, it, vi } from "vitest";

import type { MacroBackend, MacroRunResult, MacroSecurityStatus } from "../types";
import type { MacroSecurityController } from "../security";
import { MacroRunner } from "../runner";

function makeSecurityController(overrides: Partial<MacroSecurityController> = {}): MacroSecurityController {
  return {
    requestTrustDecision: vi.fn(async (_prompt: any) => null),
    requestPermissions: vi.fn(async (_prompt: any) => null),
    ...overrides,
  };
}

describe("MacroRunner", () => {
  it("prompts for a Trust Center decision when macros are blocked, then runs the macro", async () => {
    const blockedStatus: MacroSecurityStatus = {
      hasMacros: true,
      trust: "blocked",
      originPath: "/tmp/test.xlsx",
      workbookFingerprint: "fp",
      signature: { status: "unsigned" },
    };

    const runMacro = vi.fn<MacroBackend["runMacro"]>(async () => ({ ok: true, output: ["ok"] }));

    const backend: MacroBackend = {
      listMacros: vi.fn(async () => []),
      getMacroSecurityStatus: vi.fn(async () => blockedStatus),
      setMacroTrust: vi.fn(async (_workbookId, decision) => ({ ...blockedStatus, trust: decision })),
      runMacro,
    };

    const security = makeSecurityController({
      requestTrustDecision: vi.fn(async (_prompt: any) => "trusted_once" as const),
    });

    const runner = new MacroRunner(backend, security);
    const result = await runner.run({ workbookId: "wb1", macroId: "Macro1" });

    expect(result.ok).toBe(true);
    expect(security.requestTrustDecision).toHaveBeenCalledTimes(1);
    expect(backend.setMacroTrust).toHaveBeenCalledWith("wb1", "trusted_once");
    expect(runMacro).toHaveBeenCalledWith({
      workbookId: "wb1",
      macroId: "Macro1",
      permissions: [],
      timeoutMs: undefined,
    });

    const setOrder = (backend.setMacroTrust as any).mock.invocationCallOrder[0];
    const runOrder = (runMacro as any).mock.invocationCallOrder[0];
    expect(setOrder).toBeLessThan(runOrder);
  });

  it("escalates permissions once when the backend returns a permissionRequest", async () => {
    const status: MacroSecurityStatus = {
      hasMacros: true,
      trust: "trusted_once",
      originPath: "/tmp/test.xlsx",
      workbookFingerprint: "fp",
      signature: { status: "unsigned" },
    };

    const first: MacroRunResult = {
      ok: false,
      output: ["before"],
      permissionRequest: {
        reason: "permission: Network",
        macroId: "Macro1",
        workbookOriginPath: "/tmp/test.xlsx",
        requested: ["network"],
      },
      error: { message: "sandbox blocked" },
    };
    const second: MacroRunResult = { ok: true, output: ["after"] };

    const runMacro = vi
      .fn<MacroBackend["runMacro"]>()
      .mockResolvedValueOnce(first)
      .mockResolvedValueOnce(second);

    const backend: MacroBackend = {
      listMacros: vi.fn(async () => []),
      getMacroSecurityStatus: vi.fn(async () => status),
      setMacroTrust: vi.fn(async () => status),
      runMacro,
    };

    const security = makeSecurityController({
      requestPermissions: vi.fn(async (_prompt: any) => ["network" as const]),
    });

    const runner = new MacroRunner(backend, security);
    const result = await runner.run({ workbookId: "wb1", macroId: "Macro1" });

    expect(result.ok).toBe(true);
    expect(security.requestPermissions).toHaveBeenCalledTimes(1);
    expect(runMacro).toHaveBeenCalledTimes(2);
    expect(runMacro).toHaveBeenNthCalledWith(1, {
      workbookId: "wb1",
      macroId: "Macro1",
      permissions: [],
      timeoutMs: undefined,
    });
    expect(runMacro).toHaveBeenNthCalledWith(2, {
      workbookId: "wb1",
      macroId: "Macro1",
      permissions: ["network"],
      timeoutMs: undefined,
    });
    expect(result.output.join("\n")).toContain("Granted permissions: network");
  });

  it("returns a deterministic error when the user declines trust", async () => {
    const status: MacroSecurityStatus = {
      hasMacros: true,
      trust: "blocked",
      originPath: "/tmp/test.xlsx",
      workbookFingerprint: "fp",
      signature: { status: "unsigned" },
    };

    const backend: MacroBackend = {
      listMacros: vi.fn(async () => []),
      getMacroSecurityStatus: vi.fn(async () => status),
      setMacroTrust: vi.fn(async () => status),
      runMacro: vi.fn(async () => ({ ok: true, output: [] })),
    };

    const security = makeSecurityController({
      requestTrustDecision: vi.fn(async (_prompt: any) => null),
    });

    const runner = new MacroRunner(backend, security);
    const result = await runner.run({ workbookId: "wb1", macroId: "Macro1" });

    expect(result.ok).toBe(false);
    expect(result.error?.message).toBe("User declined to trust macros.");
    expect(backend.runMacro).not.toHaveBeenCalled();
  });

  it("returns a deterministic error when the user declines permissions", async () => {
    const status: MacroSecurityStatus = {
      hasMacros: true,
      trust: "trusted_once",
      originPath: "/tmp/test.xlsx",
      workbookFingerprint: "fp",
      signature: { status: "unsigned" },
    };

    const backend: MacroBackend = {
      listMacros: vi.fn(async () => []),
      getMacroSecurityStatus: vi.fn(async () => status),
      setMacroTrust: vi.fn(async () => status),
      runMacro: vi.fn(async () => ({
        ok: false,
        output: [],
        permissionRequest: {
          reason: "permission: Network",
          macroId: "Macro1",
          workbookOriginPath: "/tmp/test.xlsx",
          requested: ["network" as const],
        },
        error: { message: "sandbox blocked" },
      })),
    };

    const security = makeSecurityController({
      requestPermissions: vi.fn(async (_prompt: any) => null),
    });

    const runner = new MacroRunner(backend, security);
    const result = await runner.run({ workbookId: "wb1", macroId: "Macro1" });

    expect(result.ok).toBe(false);
    expect(result.error?.message).toBe("User declined to grant requested permissions.");
  });

  it("does not prompt for trust when trusted_signed_only and signature is verified", async () => {
    const status: MacroSecurityStatus = {
      hasMacros: true,
      trust: "trusted_signed_only",
      originPath: "/tmp/test.xlsm",
      workbookFingerprint: "fp",
      signature: { status: "signed_verified" },
    };

    const runMacro = vi.fn<MacroBackend["runMacro"]>(async () => ({ ok: true, output: ["ok"] }));

    const backend: MacroBackend = {
      listMacros: vi.fn(async () => []),
      getMacroSecurityStatus: vi.fn(async () => status),
      setMacroTrust: vi.fn(async () => status),
      runMacro,
    };

    const security = makeSecurityController();

    const runner = new MacroRunner(backend, security);
    const result = await runner.run({ workbookId: "wb1", macroId: "Macro1" });

    expect(result.ok).toBe(true);
    expect(security.requestTrustDecision).not.toHaveBeenCalled();
    expect(runMacro).toHaveBeenCalledTimes(1);
  });

  it("prompts for trust when trusted_signed_only but signature is unverified", async () => {
    const status: MacroSecurityStatus = {
      hasMacros: true,
      trust: "trusted_signed_only",
      originPath: "/tmp/test.xlsm",
      workbookFingerprint: "fp",
      signature: { status: "signed_unverified" },
    };

    const runMacro = vi.fn<MacroBackend["runMacro"]>(async () => ({ ok: true, output: ["ok"] }));

    const backend: MacroBackend = {
      listMacros: vi.fn(async () => []),
      getMacroSecurityStatus: vi.fn(async () => status),
      setMacroTrust: vi.fn(async (_workbookId, decision) => ({ ...status, trust: decision })),
      runMacro,
    };

    const security = makeSecurityController({
      requestTrustDecision: vi.fn(async (_prompt: any) => "trusted_once" as const),
    });

    const runner = new MacroRunner(backend, security);
    const result = await runner.run({ workbookId: "wb1", macroId: "Macro1" });

    expect(result.ok).toBe(true);
    expect(security.requestTrustDecision).toHaveBeenCalledTimes(1);
    expect(runMacro).toHaveBeenCalledTimes(1);
  });
});
