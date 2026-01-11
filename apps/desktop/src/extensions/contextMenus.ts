import { evaluateWhenClause, type ContextKeyLookup } from "./whenClause.js";

export type ContributedMenuItem = {
  extensionId: string;
  command: string;
  when: string | null;
  group: string | null;
};

export type ResolvedMenuItem = ContributedMenuItem & { enabled: boolean };

function parseGroup(group: string | null): { name: string; order: number } {
  if (!group) return { name: "", order: 0 };
  const raw = String(group);
  const [name, orderRaw] = raw.split("@", 2);
  const order = orderRaw != null ? Number(orderRaw) : 0;
  return { name: name ?? raw, order: Number.isFinite(order) ? order : 0 };
}

/**
 * Applies `when` clauses and basic group sorting to menu items.
 *
 * The desktop UI currently uses this for `cell/context` items.
 */
export function resolveMenuItems(items: ContributedMenuItem[], lookup: ContextKeyLookup): ResolvedMenuItem[] {
  const resolved = items.map((item) => ({
    ...item,
    enabled: evaluateWhenClause(item.when, lookup),
  }));

  resolved.sort((a, b) => {
    const ga = parseGroup(a.group);
    const gb = parseGroup(b.group);
    if (ga.name !== gb.name) return ga.name.localeCompare(gb.name);
    if (ga.order !== gb.order) return ga.order - gb.order;
    return a.command.localeCompare(b.command);
  });

  return resolved;
}
