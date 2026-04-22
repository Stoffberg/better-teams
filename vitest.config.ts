import path from "node:path";
import tailwindcss from "@tailwindcss/vite";
import react from "@vitejs/plugin-react";
import { defineConfig } from "vitest/config";

export default defineConfig({
  define: {
    "import.meta.env.VITE_ELECTRON_MAIN": JSON.stringify(false),
  },
  plugins: [react(), tailwindcss()],
  resolve: {
    alias: {
      "@better-teams/app": path.resolve(__dirname, "./packages/app/src"),
      "@better-teams/core": path.resolve(__dirname, "./packages/core/src"),
      "@better-teams/desktop-electron/preload": path.resolve(
        __dirname,
        "./packages/desktop-electron/src/preload/contracts.ts",
      ),
      "@better-teams/ui": path.resolve(__dirname, "./packages/ui/src"),
    },
  },
  test: {
    globals: true,
    environment: "happy-dom",
    setupFiles: ["./vitest.setup.ts"],
    include: ["packages/**/*.test.{ts,tsx}"],
  },
});
