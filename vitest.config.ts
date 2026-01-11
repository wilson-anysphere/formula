import { defineConfig } from "vitest/config";

export default defineConfig({
  test: {
    include: [
      "packages/**/*.test.ts",
      "packages/**/*.test.tsx",
      "apps/**/*.test.ts",
      "apps/**/*.test.tsx",
      "services/api/src/__tests__/**/*.test.ts"
    ],
    environment: "node",
    setupFiles: ["./vitest.setup.ts"]
  }
});
