import path from "node:path";
import { fileURLToPath } from "node:url";
import tailwindcss from "@tailwindcss/vite";
import react from "@vitejs/plugin-react";
import { defineConfig } from "vite";

const __dirname = path.dirname(fileURLToPath(import.meta.url));
const repoRoot = path.resolve(__dirname, "../..");

export default defineConfig({
  root: repoRoot,
  plugins: [react(), tailwindcss()],
  resolve: {
    preserveSymlinks: true,
    alias: {
      "@better-teams/app": path.resolve(__dirname, "./src"),
      "@better-teams/core": path.resolve(repoRoot, "./packages/core/src"),
      "@better-teams/desktop-electron/preload": path.resolve(
        repoRoot,
        "./packages/desktop-electron/src/preload/contracts.ts",
      ),
      "@better-teams/ui": path.resolve(repoRoot, "./packages/ui/src"),
    },
  },
  clearScreen: false,
  server: {
    port: 5173,
    strictPort: true,
    host: "127.0.0.1",
    watch: {
      ignored: [
        "**/coverage/**",
        "**/packages/desktop-electron/**",
        "**/*.test.ts",
        "**/*.test.tsx",
        "**/e2e/**",
        "**/playwright-report/**",
      ],
    },
  },
  envPrefix: ["VITE_"],
  build: {
    outDir: "dist",
    target: "chrome120",
    minify: "esbuild",
    sourcemap: false,
  },
});
