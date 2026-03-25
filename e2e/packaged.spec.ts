import { constants } from "node:fs";
import { access } from "node:fs/promises";
import path from "node:path";
import { expect, test } from "@playwright/test";
import { resolvePackagedElectronBinary } from "./packaged-binary";

test("packaged app output is present", async () => {
  const repoRoot = process.cwd();
  const outDir = path.join(repoRoot, "out");
  await access(outDir, constants.R_OK);

  try {
    const executablePath = resolvePackagedElectronBinary(repoRoot);
    await access(executablePath, constants.X_OK);
    expect(executablePath).toContain("Better Teams");
    return;
  } catch {
    const versionFile = path.join(
      outDir,
      "Better Teams-darwin-arm64",
      "version",
    );
    await access(versionFile, constants.R_OK);
    expect(versionFile).toContain("Better Teams-darwin-arm64");
  }
});
