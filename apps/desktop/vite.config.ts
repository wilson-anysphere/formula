import { defineConfig } from "vite";

export default defineConfig({
  root: ".",
  server: {
    port: 4173,
    strictPort: true
  },
  test: {
    environment: "node",
    include: ["src/**/*.vitest.ts"],
    exclude: ["tests/**", "node_modules/**"]
  }
});
