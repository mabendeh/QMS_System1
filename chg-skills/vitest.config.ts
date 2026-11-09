import { defineConfig } from "vitest/config";
import react from "@vitejs/plugin-react";
import path from "node:path";

export default defineConfig({
  plugins: [react()],
  resolve: { alias: { "@": path.resolve(__dirname, "src") } },
  test: {
    include: ["tests/unit/**/*.test.{ts,tsx}"],
    environment: "node",
    environmentMatchGlobs: [["tests/unit/**/*.test.tsx", "jsdom"]],
    setupFiles: ["tests/unit/setup.ts"],
    fileParallelism: false,
  },
});
