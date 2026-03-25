import fs from "node:fs";
import path from "node:path";

const EXECUTABLE_BASENAME = "better-teams";

export function resolvePackagedElectronBinary(repoRoot: string): string {
  const outDir = path.join(repoRoot, "out");
  if (!fs.existsSync(outDir)) {
    throw new Error("Missing out/. Run: bun run package");
  }
  const subs = fs.readdirSync(outDir);
  if (process.platform === "darwin") {
    const dir = subs.find((d) => d.includes("darwin") && !d.endsWith(".zip"));
    if (!dir) {
      throw new Error("No darwin build under out/");
    }
    const platformRoot = path.join(outDir, dir);
    const appBundles = fs
      .readdirSync(platformRoot)
      .filter((name) => name.endsWith(".app"));
    if (appBundles.length === 0) {
      throw new Error(`No .app bundle in ${platformRoot}`);
    }
    const exe = path.join(
      platformRoot,
      appBundles[0],
      "Contents",
      "MacOS",
      EXECUTABLE_BASENAME,
    );
    if (!fs.existsSync(exe)) {
      throw new Error(`Missing ${exe}`);
    }
    return exe;
  }
  if (process.platform === "linux") {
    const dir = subs.find((d) => d.includes("linux"));
    if (!dir) {
      throw new Error("No linux build under out/");
    }
    const exe = path.join(outDir, dir, EXECUTABLE_BASENAME);
    if (!fs.existsSync(exe)) {
      throw new Error(`Missing ${exe}`);
    }
    return exe;
  }
  const dir = subs.find((d) => d.includes("win32"));
  if (!dir) {
    throw new Error("No win32 build under out/");
  }
  const exe = path.join(outDir, dir, `${EXECUTABLE_BASENAME}.exe`);
  if (!fs.existsSync(exe)) {
    throw new Error(`Missing ${exe}`);
  }
  return exe;
}
