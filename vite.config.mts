import path from "node:path";
import { fileURLToPath } from "node:url";
import tailwindcss from "@tailwindcss/vite";
import react from "@vitejs/plugin-react";
import { defineConfig } from "vite";

const __dirname = path.dirname(fileURLToPath(import.meta.url));

export default defineConfig({
  plugins: [react(), tailwindcss()],
  resolve: {
    alias: {
      "@": path.resolve(__dirname, "./src"),
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
        "**/electron/**",
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
