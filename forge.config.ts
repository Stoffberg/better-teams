import { MakerDMG } from "@electron-forge/maker-dmg";
import { MakerZIP } from "@electron-forge/maker-zip";
import { VitePlugin } from "@electron-forge/plugin-vite";

const config = {
  packagerConfig: {
    asar: true,
    appBundleId: "com.betterteams.app",
    executableName: "Better Teams",
    icon: "resources/icon",
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
