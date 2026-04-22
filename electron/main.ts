import path from "node:path";
import { pathToFileURL } from "node:url";
import {
  app,
  BrowserWindow,
  ipcMain,
  Menu,
  nativeImage,
  net,
  protocol,
  shell,
  Tray,
} from "electron";
import started from "electron-squirrel-startup";
import { performFetch } from "./http";
import { cacheImageFile, removeCachedImageFiles } from "./image-cache";
import { execute, select } from "./sqlite";
import {
  extractTokens,
  getAuthToken,
  getAvailableAccounts,
  getCachedPresence,
} from "./token-store";

if (started) {
  app.quit();
}

protocol.registerSchemesAsPrivileged([
  {
    scheme: "better-teams-asset",
    privileges: {
      standard: true,
      secure: true,
      supportFetchAPI: true,
      corsEnabled: false,
    },
  },
]);

app.setName("Better Teams");
if (process.platform === "darwin") {
  app.setPath(
    "userData",
    path.join(app.getPath("appData"), "com.betterteams.app"),
  );
}

let mainWindow: BrowserWindow | null = null;
let tray: Tray | null = null;
let isQuitting = false;

const singleInstanceLock = app.requestSingleInstanceLock();
if (!singleInstanceLock) {
  app.quit();
}

registerIpc();

app.whenReady().then(() => {
  protocol.handle("better-teams-asset", (request) => {
    const filePath = filePathFromAssetUrl(request.url);
    return net.fetch(pathToFileURL(filePath).toString());
  });
  createWindow();
  createTray();
});

app.on("second-instance", () => {
  showMainWindow();
});

app.on("before-quit", () => {
  isQuitting = true;
});

app.on("window-all-closed", () => {
  if (process.platform !== "darwin") {
    app.quit();
  }
});

app.on("activate", () => {
  if (BrowserWindow.getAllWindows().length === 0) {
    createWindow();
  } else {
    showMainWindow();
  }
});

function createWindow(): void {
  mainWindow = new BrowserWindow({
    width: 1280,
    height: 860,
    minWidth: 980,
    minHeight: 640,
    title: "Better Teams",
    icon: resourcePath("resources/icon.png"),
    webPreferences: {
      preload: path.join(__dirname, "preload.js"),
      contextIsolation: true,
      nodeIntegration: false,
    },
  });

  mainWindow.on("close", (event) => {
    if (isQuitting) return;
    event.preventDefault();
    mainWindow?.hide();
  });

  const devServerUrl = MAIN_WINDOW_VITE_DEV_SERVER_URL?.replace(
    "://localhost:",
    "://127.0.0.1:",
  );
  if (devServerUrl) {
    const windowRef = mainWindow;
    void waitForUrl(devServerUrl).then(() => {
      if (!windowRef.isDestroyed()) {
        void windowRef.loadURL(devServerUrl);
      }
    });
  } else {
    void mainWindow.loadFile(
      path.join(__dirname, `../renderer/${MAIN_WINDOW_VITE_NAME}/index.html`),
    );
  }
}

async function waitForUrl(url: string): Promise<void> {
  for (let attempt = 0; attempt < 80; attempt += 1) {
    try {
      const response = await fetch(url);
      if (response.ok) return;
    } catch {
      await sleep(250);
    }
  }
  throw new Error(`Timed out waiting for ${url}`);
}

async function sleep(ms: number): Promise<void> {
  await new Promise((resolve) => setTimeout(resolve, ms));
}

function createTray(): void {
  const icon = nativeImage.createFromPath(resourcePath("resources/icon.png"));
  tray = new Tray(icon);
  tray.setToolTip("Better Teams");
  tray.setContextMenu(
    Menu.buildFromTemplate([
      { label: "Show Better Teams", click: showMainWindow },
      {
        label: "Quit",
        click: () => {
          isQuitting = true;
          app.quit();
        },
      },
    ]),
  );
  tray.on("click", showMainWindow);
}

function showMainWindow(): void {
  if (!mainWindow) {
    createWindow();
  }
  mainWindow?.show();
  if (mainWindow?.isMinimized()) mainWindow.restore();
  mainWindow?.focus();
}

function registerIpc(): void {
  ipcMain.handle("teams:extractTokens", () => extractTokens());
  ipcMain.handle("teams:getAuthToken", (_event, tenantId: string | null) =>
    getAuthToken(tenantId),
  );
  ipcMain.handle("teams:getAvailableAccounts", () => getAvailableAccounts());
  ipcMain.handle("teams:getCachedPresence", (_event, userMris: string[]) =>
    getCachedPresence(userMris),
  );
  ipcMain.handle(
    "images:cacheFile",
    (_event, cacheKey: string, bytes: number[], extension: string | null) =>
      cacheImageFile(cacheKey, Uint8Array.from(bytes), extension),
  );
  ipcMain.handle("images:removeFiles", (_event, paths: string[]) =>
    removeCachedImageFiles(paths),
  );
  ipcMain.handle("sqlite:execute", (_event, sql: string, bindValues = []) =>
    execute(sql, bindValues),
  );
  ipcMain.handle("sqlite:select", (_event, sql: string, bindValues = []) =>
    select(sql, bindValues),
  );
  ipcMain.handle("http:fetch", (_event, request) => performFetch(request));
  ipcMain.handle("shell:openExternal", (_event, url: string) =>
    shell.openExternal(url),
  );
}

function resourcePath(relativePath: string): string {
  return path.join(app.getAppPath(), relativePath);
}

function filePathFromAssetUrl(assetUrl: string): string {
  const parsed = new URL(assetUrl);
  if (parsed.hostname !== "file") {
    throw new Error("Invalid asset host");
  }
  return decodeURIComponent(parsed.pathname.slice(1));
}
