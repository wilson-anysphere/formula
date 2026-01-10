import { InMemoryAwarenessHub, PresenceManager } from "../../../collab/presence/index.js";
import { PresenceRenderer } from "../presenceRenderer.js";

const cellWidth = 60;
const cellHeight = 24;

function drawBaseGrid(ctx) {
  ctx.clearRect(0, 0, ctx.canvas.width, ctx.canvas.height);
  ctx.fillStyle = "#ffffff";
  ctx.fillRect(0, 0, ctx.canvas.width, ctx.canvas.height);

  ctx.strokeStyle = "#e5e7eb";
  ctx.lineWidth = 1;

  ctx.beginPath();
  for (let x = 0; x <= ctx.canvas.width; x += cellWidth) {
    ctx.moveTo(x + 0.5, 0);
    ctx.lineTo(x + 0.5, ctx.canvas.height);
  }
  for (let y = 0; y <= ctx.canvas.height; y += cellHeight) {
    ctx.moveTo(0, y + 0.5);
    ctx.lineTo(ctx.canvas.width, y + 0.5);
  }
  ctx.stroke();
}

function createCellRectFn() {
  return (row, col) => {
    const x = col * cellWidth;
    const y = row * cellHeight;
    if (x + cellWidth < 0 || y + cellHeight < 0) return null;
    if (x > 600 || y > 360) return null;
    return { x, y, width: cellWidth, height: cellHeight };
  };
}

function clamp(value, min, max) {
  return Math.max(min, Math.min(max, value));
}

function selectionAround(cursor) {
  return [
    {
      start: { row: cursor.row, col: cursor.col },
      end: { row: cursor.row + 1, col: cursor.col + 2 },
    },
  ];
}

const hub = new InMemoryAwarenessHub();
const awarenessA = hub.createAwareness(1);
const awarenessB = hub.createAwareness(2);

const presenceA = new PresenceManager(awarenessA, {
  user: { id: "ada", name: "Ada", color: "#ff2d55" },
  activeSheet: "Sheet1",
});

const presenceB = new PresenceManager(awarenessB, {
  user: { id: "grace", name: "Grace", color: "#4c8bf5" },
  activeSheet: "Sheet1",
});

const rendererA = new PresenceRenderer();
const rendererB = new PresenceRenderer();

const gridCtxA = document.getElementById("grid-a").getContext("2d");
const gridCtxB = document.getElementById("grid-b").getContext("2d");
const overlayCtxA = document.getElementById("overlay-a").getContext("2d");
const overlayCtxB = document.getElementById("overlay-b").getContext("2d");

drawBaseGrid(gridCtxA);
drawBaseGrid(gridCtxB);

const getCellRect = createCellRectFn();

let tick = 0;

function mountOverlay(presenceManager, renderer, ctx) {
  let scheduled = false;
  let latest = [];

  return presenceManager.subscribe((presences) => {
    latest = presences;
    if (scheduled) return;
    scheduled = true;
    requestAnimationFrame(() => {
      scheduled = false;
      renderer.clear(ctx);
      renderer.render(ctx, latest, { getCellRect });
    });
  });
}

mountOverlay(presenceA, rendererA, overlayCtxA);
mountOverlay(presenceB, rendererB, overlayCtxB);

function updateSimulatedUsers() {
  tick += 1;

  const columns = Math.floor(600 / cellWidth) - 1;
  const rows = Math.floor(360 / cellHeight) - 1;

  const cursorA = {
    row: clamp(Math.floor((Math.sin(tick / 30) + 1) * 0.5 * rows), 0, rows),
    col: clamp(Math.floor((Math.cos(tick / 40) + 1) * 0.5 * columns), 0, columns),
  };

  const cursorB = {
    row: clamp(Math.floor((Math.cos(tick / 28) + 1) * 0.5 * rows), 0, rows),
    col: clamp(Math.floor((Math.sin(tick / 36) + 1) * 0.5 * columns), 0, columns),
  };

  presenceA.setCursor(cursorA);
  presenceA.setSelections(selectionAround(cursorA));

  presenceB.setCursor(cursorB);
  presenceB.setSelections(selectionAround(cursorB));

  if (tick % 300 === 0) {
    presenceB.setActiveSheet(presenceB.localPresence.activeSheet === "Sheet1" ? "Sheet2" : "Sheet1");
  }
}

function tickLoop() {
  updateSimulatedUsers();
  requestAnimationFrame(tickLoop);
}

tickLoop();
