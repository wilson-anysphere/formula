import { expect, test } from "@playwright/test";

test.describe("grid scroll performance", () => {
  test.skip(process.env.PERF_TESTS !== "1", "Set PERF_TESTS=1 to run performance benchmarks");

  test("maintains 60fps-class scrolling on a large grid", async ({ page }) => {
    await page.goto("/?perf=1");

    await expect(page.getByTestId("engine-status")).toContainText("ready");

    await page.waitForFunction(() => {
      const api = (window as any).__gridApi as { scrollBy?: (dx: number, dy: number) => void } | undefined;
      return typeof api?.scrollBy === "function";
    });

    const result = await page.evaluate(
      ({ frames, deltaY }) =>
        new Promise<{ avg: number; p95: number; samples: number[] }>((resolve) => {
          const api = (window as any).__gridApi as { scrollBy: (dx: number, dy: number) => void } | undefined;
          if (!api) throw new Error("Missing __gridApi – did you load with ?perf=1?");

          let remaining = frames;
          let last = performance.now();
          const samples: number[] = [];

          const tick = (now: number) => {
            const dt = now - last;
            last = now;
            samples.push(dt);

            api.scrollBy(0, deltaY);
            remaining -= 1;

            if (remaining > 0) {
              requestAnimationFrame(tick);
              return;
            }

            const trimmed = samples.slice(1);
            const avg = trimmed.reduce((sum, value) => sum + value, 0) / Math.max(1, trimmed.length);
            const sorted = [...trimmed].sort((a, b) => a - b);
            const p95Index = Math.min(sorted.length - 1, Math.floor((sorted.length - 1) * 0.95));
            const p95 = sorted[p95Index] ?? 0;
            resolve({ avg, p95, samples: trimmed });
          };

          requestAnimationFrame((now) => {
            last = now;
            requestAnimationFrame(tick);
          });
        }),
      { frames: 120, deltaY: 120 }
    );

    expect(result.avg).toBeLessThan(20);
    expect(result.p95).toBeLessThan(30);
  });
});
