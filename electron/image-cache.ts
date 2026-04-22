import { createHash } from "node:crypto";
import fs from "node:fs";
import os from "node:os";
import path from "node:path";
import { app } from "electron";

const IMAGE_EXTENSIONS = ["jpg", "png", "gif", "webp", "avif", "img"] as const;

export function cacheImageFile(
  cacheKey: string,
  bytes: Uint8Array,
  extension?: string | null,
): string {
  const dir = imageCacheDir();
  const filePath = path.join(dir, hashedFilename(cacheKey, extension));
  removeCachedImageVariants(cacheKey, filePath);
  fs.writeFileSync(filePath, bytes);
  return filePath;
}

export function getCachedImageFile(cacheKey: string): string | null {
  const dir = imageCacheDir();
  for (const extension of IMAGE_EXTENSIONS) {
    const filePath = path.join(dir, hashedFilename(cacheKey, extension));
    if (isFile(filePath)) return filePath;
  }
  return null;
}

export function hasCachedImageFile(filePath: string): boolean {
  const candidate = path.resolve(filePath);
  return isWithinDir(candidate, imageCacheDir()) && isFile(candidate);
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

function removeCachedImageVariants(cacheKey: string, keepPath: string): void {
  const dir = imageCacheDir();
  const keep = path.resolve(keepPath);
  for (const extension of IMAGE_EXTENSIONS) {
    const filePath = path.resolve(
      path.join(dir, hashedFilename(cacheKey, extension)),
    );
    if (filePath === keep) continue;
    try {
      fs.rmSync(filePath, { force: true });
    } catch {}
  }
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

function isFile(filePath: string): boolean {
  try {
    return fs.statSync(filePath).isFile();
  } catch {
    return false;
  }
}
