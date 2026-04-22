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
          entry: "electron/main.ts",
          config: "vite.main.config.mts",
          target: "main",
        },
        {
          entry: "electron/preload.ts",
          config: "vite.preload.config.mts",
          target: "preload",
        },
      ],
      renderer: [
        {
          name: "main_window",
          config: "vite.config.mts",
        },
      ],
    }),
  ],
};

export default config;
