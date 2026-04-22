import fs from "node:fs";
import path from "node:path";
import { fileURLToPath } from "node:url";
import { describe, expect, it } from "vitest";

const repoRoot = path.resolve(
  path.dirname(fileURLToPath(import.meta.url)),
  "../../..",
);

function sourceFiles(dir: string): string[] {
  if (!fs.existsSync(dir)) return [];
  return fs.readdirSync(dir, { withFileTypes: true }).flatMap((entry) => {
    const fullPath = path.join(dir, entry.name);
    if (entry.isDirectory()) return sourceFiles(fullPath);
    return /\.(ts|tsx)$/.test(entry.name) ? [fullPath] : [];
  });
}

function relative(filePath: string): string {
  return path.relative(repoRoot, filePath);
}

function filesMatching(
  dir: string,
  predicate: (filePath: string, source: string) => boolean,
): string[] {
  return sourceFiles(path.join(repoRoot, dir))
    .filter((filePath) =>
      predicate(filePath, fs.readFileSync(filePath, "utf8")),
    )
    .map(relative);
}

describe("architecture boundaries", () => {
  it("keeps core independent from app, ui, React, and Electron", () => {
    expect(
      filesMatching("packages/core/src", (_filePath, source) =>
        /from ["'](@better-teams\/app|@better-teams\/ui|@better-teams\/desktop-electron|react|electron)(\/|["'])/.test(
          source,
        ),
      ),
    ).toEqual([]);
  });

  it("keeps app desktop preload access inside the desktop service boundary", () => {
    expect(
      filesMatching("packages/app/src", (filePath, source) => {
        if (relative(filePath).includes(".test.")) return false;
        if (relative(filePath) === "packages/app/src/renderer-env.d.ts") {
          return false;
        }
        if (
          relative(filePath).startsWith("packages/app/src/services/desktop/")
        ) {
          return false;
        }
        return source.includes("@better-teams/desktop-electron/preload");
      }),
    ).toEqual([]);
  });

  it("keeps raw Teams client access inside app services and tests", () => {
    expect(
      filesMatching("packages/app/src", (filePath, source) => {
        const rel = relative(filePath);
        if (rel.includes(".test.")) return false;
        if (rel.startsWith("packages/app/src/services/teams/")) return false;
        return source.includes("@better-teams/core/teams/client/factory");
      }),
    ).toEqual([]);
  });

  it("does not reintroduce the app lib junk drawer", () => {
    expect(fs.existsSync(path.join(repoRoot, "packages/app/src/lib"))).toBe(
      false,
    );
  });
});
