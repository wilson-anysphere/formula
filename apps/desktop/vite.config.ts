import { defineConfig } from "vite";

export default defineConfig({
  root: ".",
  server: {
    port: 4174,
    strictPort: true
  },
  test: {
    environment: "node",
    include: ["src/**/*.vitest.ts"],
    exclude: ["tests/**", "node_modules/**"]
  }
});
