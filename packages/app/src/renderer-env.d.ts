/// <reference types="vite/client" />

import type { BetterTeamsDesktopApi } from "@better-teams/desktop-electron/preload";

declare global {
  interface Window {
    betterTeams?: BetterTeamsDesktopApi;
  }
}
