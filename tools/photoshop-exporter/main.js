/* Copyright (C) 2026 Tencent. Licensed under the Apache License, Version 2.0. */
"use strict";
const ps = require("photoshop");
const { storage, host } = require("uxp");
const { NativeBridge } = require("./src/native-bridge.js");
const { PhotoshopHost } = require("./src/photoshop-host.js");
const { OutputFiles } = require("./src/output-files.js");
const { ExportService, ExportCancelled } = require("./src/export-service.js");
const status = document.getElementById("status");
const details = document.getElementById("details");
const current = document.getElementById("current");
const fileButton = document.getElementById("file");
const cancel = document.getElementById("cancel");
let service;
let activeSignal;

function busy(value) {
  current.disabled = value;
  fileButton.disabled = value;
  cancel.disabled = !value;
}
async function run(kind) {
  if (activeSignal) return;
  activeSignal = { cancelled: false };
  busy(true);
  details.textContent = "";
  try {
    let input;
    if (kind === "current") {
      if (!ps.app.documents.length) throw new Error("请先在 Photoshop 中打开文档");
      input = { kind, documentId: ps.app.activeDocument.id };
    } else {
      const file = await storage.localFileSystem.getFileForOpening({ types: ["psd"], allowMultiple: false });
      if (!file) throw new ExportCancelled();
      input = { kind, file };
    }
    status.textContent = "请选择 PAGX 保存目录…";
    const folder = await storage.localFileSystem.getFolder();
    if (!folder) throw new ExportCancelled();
    const result = await service.run({ input, folder, signal: activeSignal,
      progress: message => { status.textContent = message; } });
    status.textContent = `已导出：${result.file.name}`;
    details.textContent = result.file.nativePath;
  } catch (error) {
    status.textContent = error instanceof ExportCancelled || activeSignal.cancelled || (error && error.number === 8007)
      ? "已取消导出" : `导出失败：${error && error.message || String(error)}`;
    console.error(error);
  } finally { activeSignal = undefined; busy(false); }
}
cancel.addEventListener("click", () => { if (activeSignal) activeSignal.cancelled = true; });
current.addEventListener("click", () => run("current"));
fileButton.addEventListener("click", () => run("file"));
(async function initialize() {
  try {
    // Hybrid modules load asynchronously in Photoshop, unlike regular CommonJS modules.
    const bridge = new NativeBridge(await require("pagx-photoshop.uxpaddon"));
    service = new ExportService({ bridge, host: new PhotoshopHost(ps, bridge.info.maxPixels, host.version),
      files: new OutputFiles(storage.formats.utf8) });
    status.textContent = "准备就绪";
    busy(false);
  } catch (error) {
    status.textContent = `导出模块加载失败：${error && error.message || String(error)}`;
    details.textContent = "请使用与系统架构匹配、包含原生模块的完整插件。";
    console.error(error);
  }
})();
