import { execFileSync } from "node:child_process";
import { cpSync, mkdirSync, rmSync } from "node:fs";
import path from "node:path";
import { fileURLToPath } from "node:url";

const rootDir = path.resolve(
  path.dirname(fileURLToPath(import.meta.url)),
  "..",
);
const brandingDir = path.join(rootDir, "assets", "branding");
const generatedDir = path.join(brandingDir, "generated");
const publicDir = path.join(rootDir, "public");
const iconSvg = path.join(brandingDir, "icon.svg");
const traySvg = path.join(brandingDir, "tray-template.svg");
const socialSvg = path.join(brandingDir, "social-card.svg");
const iconsetDir = path.join(generatedDir, "better-teams.iconset");

function run(command, args) {
  execFileSync(command, args, { stdio: "inherit" });
}

function ensureDir(dir) {
  mkdirSync(dir, { recursive: true });
}

/** Render SVG to PNG using rsvg-convert for proper gradient/filter support */
function renderPng(svgPath, outputPath, width, height = width) {
  run("rsvg-convert", [
    "-w",
    String(width),
    "-h",
    String(height),
    "-o",
    outputPath,
    svgPath,
  ]);
}

ensureDir(generatedDir);
ensureDir(publicDir);
rmSync(iconsetDir, { recursive: true, force: true });
ensureDir(iconsetDir);

const squareSizes = [16, 32, 48, 64, 128, 180, 192, 256, 512, 1024];

for (const size of squareSizes) {
  renderPng(iconSvg, path.join(generatedDir, `icon-${size}.png`), size);
}

cpSync(
  path.join(generatedDir, "icon-512.png"),
  path.join(generatedDir, "icon.png"),
);
cpSync(
  path.join(generatedDir, "icon-512.png"),
  path.join(publicDir, "icon-512.png"),
);
cpSync(
  path.join(generatedDir, "icon-192.png"),
  path.join(publicDir, "icon-192.png"),
);
cpSync(
  path.join(generatedDir, "icon-180.png"),
  path.join(publicDir, "apple-touch-icon.png"),
);
cpSync(
  path.join(generatedDir, "icon-32.png"),
  path.join(publicDir, "favicon-32x32.png"),
);
cpSync(
  path.join(generatedDir, "icon-16.png"),
  path.join(publicDir, "favicon-16x16.png"),
);
cpSync(iconSvg, path.join(publicDir, "favicon.svg"));

// Generate Windows ICO from multiple sizes
run("magick", [
  path.join(generatedDir, "icon-16.png"),
  path.join(generatedDir, "icon-32.png"),
  path.join(generatedDir, "icon-48.png"),
  path.join(generatedDir, "icon-64.png"),
  path.join(generatedDir, "icon-128.png"),
  path.join(generatedDir, "icon-256.png"),
  path.join(generatedDir, "icon.ico"),
]);

// Generate macOS iconset
const iconsetMappings = [
  ["icon_16x16.png", 16],
  ["icon_16x16@2x.png", 32],
  ["icon_32x32.png", 32],
  ["icon_32x32@2x.png", 64],
  ["icon_128x128.png", 128],
  ["icon_128x128@2x.png", 256],
  ["icon_256x256.png", 256],
  ["icon_256x256@2x.png", 512],
  ["icon_512x512.png", 512],
  ["icon_512x512@2x.png", 1024],
];

for (const [filename, size] of iconsetMappings) {
  cpSync(
    path.join(generatedDir, `icon-${size}.png`),
    path.join(iconsetDir, filename),
  );
}

run("iconutil", [
  "-c",
  "icns",
  iconsetDir,
  "-o",
  path.join(generatedDir, "icon.icns"),
]);

// Generate tray icons
renderPng(iconSvg, path.join(generatedDir, "tray.png"), 32);
renderPng(traySvg, path.join(generatedDir, "trayTemplate.png"), 32);

// Generate social card PNG
renderPng(socialSvg, path.join(publicDir, "social-card.png"), 1200, 630);
