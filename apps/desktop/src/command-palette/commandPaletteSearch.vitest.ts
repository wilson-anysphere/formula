import { describe, expect, it } from "vitest";

import type { CommandContribution } from "../extensions/commandRegistry.js";
import { buildCommandPaletteSections } from "./commandPaletteSearch.js";
import { prepareCommandForFuzzy } from "./fuzzy.js";

describe("command palette search", () => {
  it("returns functions in a FUNCTIONS section ranked by fuzzy score", () => {
    const sections = buildCommandPaletteSections({
      query: "sum",
      commands: [],
      limits: { maxResults: 50, maxResultsPerGroup: 50 },
    });
    const functions = sections.find((s) => s.title === "FUNCTIONS");
    expect(functions).toBeTruthy();

    const names = functions!.results.filter((r) => r.kind === "function").map((r) => r.name);
    expect(names.length).toBeGreaterThan(0);

    // Exact match should outrank longer prefix matches.
    expect(names[0]).toBe("SUM");
    expect(names).toContain("SUMIF");
  });

  it("matches functions when users include parentheses (e.g. SUM()", () => {
    const sections = buildCommandPaletteSections({
      query: "sum(",
      commands: [],
      limits: { maxResults: 50, maxResultsPerGroup: 50 },
    });
    const functions = sections.find((s) => s.title === "FUNCTIONS");
    expect(functions).toBeTruthy();
    expect(functions!.results[0]).toMatchObject({ kind: "function", name: "SUM" });
  });

  it("groups recent commands separately from the COMMANDS section", () => {
    const commands: CommandContribution[] = [
      { commandId: "freeze", title: "Freeze Panes", category: "View", source: { kind: "builtin" } },
      { commandId: "pivot", title: "Insert Pivot Table", category: "Insert", source: { kind: "builtin" } },
    ];

    const sections = buildCommandPaletteSections({
      query: "",
      commands: commands.map((cmd) => prepareCommandForFuzzy(cmd)),
      recentIds: ["pivot"],
      limits: { maxResults: 50, maxResultsPerGroup: 50 },
    });

    expect(sections.map((s) => s.title)).toEqual(["RECENT", "COMMANDS"]);
    expect(sections[0]!.results.map((r) => (r.kind === "command" ? r.command.commandId : null))).toEqual(["pivot"]);
    expect(sections[1]!.results.map((r) => (r.kind === "command" ? r.command.commandId : null))).toEqual(["freeze"]);
  });

  it("ranks exact function matches above similarly named commands (sum → SUM before AutoSum)", () => {
    const commands: CommandContribution[] = [
      {
        commandId: "edit.autoSum",
        title: "AutoSum",
        category: "Editing",
        source: { kind: "builtin" },
        description: "Insert a SUM formula based on adjacent cells",
        keywords: ["autosum", "sum"],
      },
    ];

    const sections = buildCommandPaletteSections({
      query: "sum",
      commands: commands.map((cmd) => prepareCommandForFuzzy(cmd)),
      limits: { maxResults: 50, maxResultsPerGroup: 50 },
    });

    expect(sections[0]?.title).toBe("FUNCTIONS");
    expect(sections[0]?.results[0]).toMatchObject({ kind: "function", name: "SUM" });
  });
});
