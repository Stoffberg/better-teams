import path from "node:path";
import { fileURLToPath } from "node:url";
import { defineConfig } from "vite";

const __dirname = path.dirname(fileURLToPath(import.meta.url));
const repoRoot = path.resolve(__dirname, "../..");

export default defineConfig({
  resolve: {
    alias: {
      "@better-teams/core": path.resolve(repoRoot, "./packages/core/src"),
      "@better-teams/desktop-electron": path.resolve(__dirname, "./src"),
      "@better-teams/desktop-electron/preload": path.resolve(
        __dirname,
        "./src/preload/contracts.ts",
      ),
    },
  },
});
