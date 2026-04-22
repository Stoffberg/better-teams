import { cpSync, mkdirSync, rmSync } from "node:fs";
import { createRequire } from "node:module";
import path from "node:path";
import { MakerDMG } from "@electron-forge/maker-dmg";
import { MakerZIP } from "@electron-forge/maker-zip";
import { VitePlugin } from "@electron-forge/plugin-vite";

const require = createRequire(import.meta.url);

const nativeRuntimeModules = ["better-sqlite3", "bindings", "file-uri-to-path"];

const copyNativeRuntimeModule = (moduleName: string, buildPath: string) => {
  const packageJsonPath = require.resolve(`${moduleName}/package.json`);
  const modulePath = path.dirname(packageJsonPath);
  const targetPath = path.join(buildPath, "node_modules", moduleName);

  rmSync(targetPath, { force: true, recursive: true });
  mkdirSync(path.dirname(targetPath), { recursive: true });
  cpSync(modulePath, targetPath, { dereference: true, recursive: true });
};

const config = {
  packagerConfig: {
    asar: {
      unpack: "**/*.node",
    },
    appBundleId: "com.betterteams.app",
    executableName: "Better Teams",
    icon: "resources/icon",
  },
  hooks: {
    packageAfterCopy: async (_config, buildPath) => {
      for (const moduleName of nativeRuntimeModules) {
        copyNativeRuntimeModule(moduleName, buildPath);
      }
    },
  },
  rebuildConfig: {},
  makers: [new MakerZIP({}, ["darwin"]), new MakerDMG({})],
  plugins: [
    new VitePlugin({
      build: [
        {
          entry: "packages/desktop-electron/src/main/main.ts",
          config: "packages/desktop-electron/vite.main.config.mts",
          target: "main",
        },
        {
          entry: "packages/desktop-electron/src/preload/preload.ts",
          config: "packages/desktop-electron/vite.preload.config.mts",
          target: "preload",
        },
      ],
      renderer: [
        {
          name: "main_window",
          config: "packages/app/vite.config.mts",
        },
      ],
    }),
  ],
};

export default config;
