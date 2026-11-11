/* Copyright (C) 2026 Tencent. Licensed under the Apache License, Version 2.0. */
"use strict";
const fs = require("node:fs");
const path = require("node:path");
const { execFileSync } = require("node:child_process");
const plugin = path.resolve(__dirname, "..");
const repo = path.resolve(plugin, "../..");
const sdk = process.env.UXP_SDK_DIR;
if (!sdk || !fs.existsSync(path.join(sdk, "src/utilities/UxpAddon.h"))) {
  throw new Error("Set UXP_SDK_DIR to the extracted official Adobe UXP Hybrid Plugin SDK");
}
if (!["darwin", "win32"].includes(process.platform)) throw new Error("Photoshop requires macOS or Windows");
const architecture = process.env.PAGX_ARCH || process.arch;
if (!["arm64", "x64"].includes(architecture) || (process.platform === "win32" && architecture !== "x64")) {
  throw new Error("Unsupported Photoshop addon architecture");
}
const build = process.env.PAGX_BUILD_DIR || path.join(repo, `cmake-build-photoshop-${architecture}`);
const args = ["-S", repo, "-B", build, "-DCMAKE_BUILD_TYPE=Release", "-DPAG_BUILD_PHOTOSHOP_PLUGIN=ON",
  "-DPAG_BUILD_SHARED=OFF", "-DPAG_BUILD_FRAMEWORK=OFF", "-DPAG_BUILD_TESTS=OFF", "-DPAG_USE_LIBAVC=OFF",
  `-DUXP_SDK_DIR=${path.resolve(sdk)}`];
if (process.platform === "darwin") args.push(`-DCMAKE_OSX_ARCHITECTURES=${architecture === "x64" ? "x86_64" : "arm64"}`);
else args.push("-A", "x64");
execFileSync("cmake", args, { stdio: "inherit" });
execFileSync("cmake", ["--build", build, "--config", "Release", "--target", "pagx-photoshop-addon", "-j", "8"], { stdio: "inherit" });
const output = path.join(plugin, process.platform === "darwin" ? "mac" : "win", architecture);
fs.mkdirSync(output, { recursive: true });
const filename = "pagx-photoshop.uxpaddon";
const candidates = [path.join(build, "tools/photoshop-exporter/native", filename),
  path.join(build, "tools/photoshop-exporter/native/Release", filename)];
const binary = candidates.find(file => fs.existsSync(file));
if (!binary) throw new Error("CMake did not produce the addon");
const destination = path.join(output, filename);
fs.copyFileSync(binary, destination);
if (process.platform === "darwin") {
  // Development signature only. Marketplace distribution needs Developer ID + notarization.
  execFileSync("codesign", ["--force", "--sign", "-", destination], { stdio: "inherit" });
}
console.log(`Built development addon: ${destination}`);
