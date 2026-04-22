import path from "node:path";
import { pathToFileURL } from "node:url";
import { Worker } from "node:worker_threads";
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
import {
  FetchRequestSchema,
  ImageCacheIpcRequestSchema,
  ImageCachePathSchema,
  PresenceRequestSchema,
  ProfilePresentationRequestSchema,
  ShellOpenExternalUrlSchema,
  TenantIdSchema,
} from "../preload/contracts";
import { performFetch } from "./http";
import {
  cachedImagePathFromFileName,
  cacheImageFile,
  getCachedImageFile,
  hasCachedImageFile,
} from "./image-cache";
import {
  getPersistedAccounts,
  getPersistedConversations,
  getPersistedMessages,
  getPersistedProfilePresentation,
  getPersistedSession,
} from "./persisted-teams-cache";
import type {
  AccountOption,
  CachedPresenceEntry,
  ExtractedToken,
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

type TokenWorkerRequest =
  | { operation: "extractTokens" }
  | { operation: "getAuthToken"; tenantId?: string | null }
  | { operation: "getAvailableAccounts" }
  | { operation: "getCachedPresence"; userMris: string[] };

type TokenWorkerResponse =
  | { ok: true; value: unknown }
  | { ok: false; error: string };

const singleInstanceLock = app.requestSingleInstanceLock();
if (!singleInstanceLock) {
  app.quit();
}

registerIpc();

app.whenReady().then(() => {
  protocol.handle("better-teams-asset", (request) => {
    try {
      const filePath = filePathFromAssetUrl(request.url);
      if (!hasCachedImageFile(filePath)) {
        return new Response(null, { status: 404 });
      }
      return net.fetch(pathToFileURL(filePath).toString());
    } catch {
      return new Response(null, { status: 404 });
    }
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
  ipcMain.handle("teams:extractTokens", () =>
    runTokenWorker<ExtractedToken[]>({ operation: "extractTokens" }),
  );
  ipcMain.handle("teams:getAuthToken", (_event, tenantId) =>
    runTokenWorker<ExtractedToken | null>({
      operation: "getAuthToken",
      tenantId: TenantIdSchema.parse(tenantId),
    }),
  );
  ipcMain.handle("teams:getAvailableAccounts", () =>
    runTokenWorker<AccountOption[]>({ operation: "getAvailableAccounts" }),
  );
  ipcMain.handle("teams:getCachedAccounts", () => getPersistedAccounts());
  ipcMain.handle("teams:getCachedSession", (_event, tenantId) =>
    getPersistedSession(TenantIdSchema.parse(tenantId)),
  );
  ipcMain.handle("teams:getCachedPresence", (_event, userMris) =>
    runTokenWorker<CachedPresenceEntry[]>({
      operation: "getCachedPresence",
      userMris: PresenceRequestSchema.parse(userMris),
    }),
  );
  ipcMain.handle("teams:getCachedProfilePresentation", (_event, mris) =>
    getPersistedProfilePresentation(
      ProfilePresentationRequestSchema.parse(mris),
    ),
  );
  ipcMain.handle("teams:getCachedConversations", (_event, tenantId) =>
    getPersistedConversations(TenantIdSchema.parse(tenantId)),
  );
  ipcMain.handle(
    "teams:getCachedMessages",
    (_event, tenantId, conversationId) =>
      getPersistedMessages(
        TenantIdSchema.parse(tenantId),
        ImageCachePathSchema.parse(conversationId),
      ),
  );
  ipcMain.handle("images:cacheFile", (_event, cacheKey, bytes, extension) => {
    const request = ImageCacheIpcRequestSchema.parse({
      cacheKey,
      bytes,
      extension,
    });
    return cacheImageFile(
      request.cacheKey,
      Uint8Array.from(request.bytes),
      request.extension,
    );
  });
  ipcMain.handle("images:getCachedFile", (_event, cacheKey) =>
    getCachedImageFile(ImageCachePathSchema.parse(cacheKey)),
  );
  ipcMain.handle("images:hasCachedFile", (_event, filePath) =>
    hasCachedImageFile(ImageCachePathSchema.parse(filePath)),
  );
  ipcMain.handle("http:fetch", (_event, request) =>
    performFetch(FetchRequestSchema.parse(request)),
  );
  ipcMain.handle("shell:openExternal", (_event, url) =>
    shell.openExternal(ShellOpenExternalUrlSchema.parse(url)),
  );
}

function resourcePath(relativePath: string): string {
  return path.join(app.getAppPath(), relativePath);
}

function runTokenWorker<T>(request: TokenWorkerRequest): Promise<T> {
  return new Promise((resolve, reject) => {
    const worker = new Worker(path.join(__dirname, "token-worker.js"), {
      workerData: request,
    });
    let settled = false;
    const timeout = setTimeout(() => {
      if (settled) return;
      settled = true;
      void worker.terminate();
      reject(new Error("Timed out reading Teams token store"));
    }, 10_000);
    const finish = (callback: () => void) => {
      if (settled) return;
      settled = true;
      clearTimeout(timeout);
      callback();
    };

    worker.once("message", (message: TokenWorkerResponse) => {
      finish(() => {
        if (message.ok) {
          resolve(message.value as T);
          return;
        }
        reject(new Error(message.error));
      });
    });
    worker.once("error", (error) => {
      finish(() => reject(error));
    });
    worker.once("exit", (code) => {
      if (code === 0) return;
      finish(() => reject(new Error(`Teams token worker exited with ${code}`)));
    });
  });
}

function filePathFromAssetUrl(assetUrl: string): string {
  const parsed = new URL(assetUrl);
  if (parsed.hostname === "cache") {
    return cachedImagePathFromFileName(decodeURIComponent(parsed.pathname));
  }
  if (parsed.hostname !== "file") {
    throw new Error("Invalid asset host");
  }
  return decodeURIComponent(parsed.pathname.slice(1));
}
