type ClientId = 'a' | 'b';

type Patch = {
  clientId: ClientId;
  clock: number;
  cell: string;
  value: string;
};

type CellVersion = { clock: number; clientId: ClientId };

type Client = {
  id: ClientId;
  clock: number;
  cells: Map<string, string>;
  versions: Map<string, CellVersion>;
};

function createClient(id: ClientId): Client {
  return { id, clock: 0, cells: new Map(), versions: new Map() };
}

function compareVersions(a: CellVersion, b: CellVersion): number {
  if (a.clock !== b.clock) return a.clock - b.clock;
  return a.clientId < b.clientId ? -1 : a.clientId > b.clientId ? 1 : 0;
}

function makeEdit(client: Client, cell: string, value: string): Patch {
  client.clock += 1;
  client.cells.set(cell, value);
  client.versions.set(cell, { clock: client.clock, clientId: client.id });
  return { clientId: client.id, clock: client.clock, cell, value };
}

function applyPatch(client: Client, patch: Patch): void {
  const incoming: CellVersion = { clock: patch.clock, clientId: patch.clientId };
  const current = client.versions.get(patch.cell);
  if (!current || compareVersions(current, incoming) < 0) {
    client.cells.set(patch.cell, patch.value);
    client.versions.set(patch.cell, incoming);
  }
}

export function createCollaborationBenchmarks(): Array<{
  name: string;
  fn: () => void;
  targetMs: number;
  iterations?: number;
  warmup?: number;
  clock?: 'wall' | 'cpu';
}> {
  const a = createClient('a');
  const b = createClient('b');

  let counter = 0;

  return [
    {
      name: 'collab.edit_to_sync.p95',
      fn: () => {
        counter += 1;
        // “Edit” a single cell and sync to the other client via a serialized patch.
        const patch = makeEdit(a, 'A1', `v${counter}`);
        const encoded = JSON.stringify(patch);
        const decoded = JSON.parse(encoded) as Patch;
        applyPatch(b, decoded);
      },
      // Target per docs/16-performance-targets.md (<100ms). We use a stricter
      // guardrail since this is an in-memory simulation.
      targetMs: 5,
    },
    {
      name: 'collab.conflict_resolution.p95',
      fn: () => {
        counter += 1;
        // Simulate concurrent edits (offline) then merge updates.
        const p1 = makeEdit(a, 'B2', `a${counter}`);
        const p2 = makeEdit(b, 'B2', `b${counter}`);

        // Exchange patches in opposite order to force resolution.
        applyPatch(a, p2);
        applyPatch(b, p1);

        // Sanity check: both clients converge.
        const va = a.cells.get('B2');
        const vb = b.cells.get('B2');
        if (va !== vb) throw new Error(`diverged: a=${va} b=${vb}`);
      },
      targetMs: 10,
    },
  ];
}
