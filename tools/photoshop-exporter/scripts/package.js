/* Copyright (C) 2026 Tencent. Licensed under the Apache License, Version 2.0. */
"use strict";
const fs = require("node:fs");
const path = require("node:path");
const { execFileSync } = require("node:child_process");
const plugin = path.resolve(__dirname, "..");
const dist = path.join(plugin, "dist");
const staging = path.join(dist, "plugin");
const manifest = require("../manifest.json");
const platforms = ["mac/arm64", "mac/x64", "win/x64"].filter(dir =>
  fs.existsSync(path.join(plugin, dir, manifest.addon.name)));
if (!platforms.length) throw new Error("Build the native addon first (npm run build:native)");
fs.rmSync(staging, { recursive: true, force: true });
fs.mkdirSync(staging, { recursive: true });
for (const name of ["manifest.json", "index.html", "styles.css", "main.js", "src"]) {
  fs.cpSync(path.join(plugin, name), path.join(staging, name), { recursive: true });
}
for (const dir of platforms) {
  fs.mkdirSync(path.join(staging, dir), { recursive: true });
  fs.copyFileSync(path.join(plugin, dir, manifest.addon.name), path.join(staging, dir, manifest.addon.name));
}
const output = path.join(dist, `pagx-photoshop-${manifest.version}-dev.ccx`);
fs.rmSync(output, { force: true });
if (process.platform === "win32") {
  throw new Error("Package dist/plugin using UXP Developer Tool on Windows");
}
execFileSync("zip", ["-q", "-r", "-X", output, "."], { cwd: staging, stdio: "inherit" });
console.log(`Development package: ${output}\nIncluded architectures: ${platforms.join(", ")}`);
console.log("This is a local development build, not a signed/notarized Marketplace release.");
