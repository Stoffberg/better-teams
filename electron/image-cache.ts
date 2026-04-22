import { createHash } from "node:crypto";
import fs from "node:fs";
import os from "node:os";
import path from "node:path";
import { app } from "electron";

export function cacheImageFile(
  cacheKey: string,
  bytes: Uint8Array,
  extension?: string | null,
): string {
  const dir = imageCacheDir();
  const filePath = path.join(dir, hashedFilename(cacheKey, extension));
  fs.writeFileSync(filePath, bytes);
  return filePath;
}

export function removeCachedImageFiles(paths: string[]): void {
  const dir = imageCacheDir();
  for (const filePath of paths) {
    const candidate = path.resolve(filePath);
    if (!isWithinDir(candidate, dir)) continue;
    try {
      fs.rmSync(candidate, { force: true });
    } catch (error) {
      throw new Error(`Failed to remove cached image file: ${String(error)}`);
    }
  }
}

function imageCacheDir(): string {
  const dir =
    process.platform === "darwin"
      ? path.join(os.homedir(), "Library/Caches/com.betterteams.app/images")
      : path.join(app.getPath("userData"), "Cache", "images");
  fs.mkdirSync(dir, { recursive: true });
  return dir;
}

function hashedFilename(cacheKey: string, extension?: string | null): string {
  const digest = createHash("sha1").update(cacheKey).digest("hex");
  return `${digest}.${normalizedExtension(extension)}`;
}

function normalizedExtension(extension?: string | null): string {
  const value = extension?.trim().toLowerCase();
  if (value === "jpg" || value === "jpeg") return "jpg";
  if (value === "png") return "png";
  if (value === "gif") return "gif";
  if (value === "webp") return "webp";
  if (value === "avif") return "avif";
  return "img";
}

function isWithinDir(filePath: string, dir: string): boolean {
  const relative = path.relative(dir, filePath);
  return (
    Boolean(relative) &&
    !relative.startsWith("..") &&
    !path.isAbsolute(relative)
  );
}
