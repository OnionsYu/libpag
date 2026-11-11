/* Copyright (C) 2026 Tencent. Licensed under the Apache License, Version 2.0. */
"use strict";
// Run from UDT's Photoshop console, not Node. Test code is excluded from the .ccx package.
const ps = require("photoshop");
const uxp = require("uxp");
const { PhotoshopHost } = require("../src/photoshop-host.js");
const { NativeBridge } = require("../src/native-bridge.js");
const { OutputFiles } = require("../src/output-files.js");
const { ExportService } = require("../src/export-service.js");

function documentState() {
  return JSON.stringify({ active: ps.app.documents.length ? ps.app.activeDocument.id : null,
    documents: Array.from(ps.app.documents).map(doc => ({ id: doc.id, name: doc.name,
      saved: doc.saved, history: doc.historyStates.length,
      selected: doc.activeLayers.map(layer => layer.id) })) });
}

async function run(folder) {
  folder = folder || await uxp.storage.localFileSystem.getFolder();
  if (!folder) return;
  const bridge = new NativeBridge(await require("pagx-photoshop.uxpaddon"));
  const service = new ExportService({ bridge,
    host: new PhotoshopHost(ps, bridge.info.maxPixels, uxp.host.version),
    files: new OutputFiles(uxp.storage.formats.utf8) });
  const before = documentState();
  const entries = await folder.getEntries();
  const report = { hostVersion: uxp.host.version, startedAt: new Date().toISOString(), cases: [] };
  const inputFile = name => {
    const file = entries.find(entry => entry.name === name);
    if (!file) throw new Error(`Missing fixture: ${name}; run make-fixtures.py first`);
    return { kind: "file", file };
  };
  const checkState = () => {
    if (documentState() !== before) throw new Error("Source document state changed or temporary document leaked");
    if (service.busy) throw new Error("Export stayed busy");
  };
  // Let Photoshop produce the 16-bit PSD, including its format-specific layer records.
  const fixture16 = await folder.createFile(`photoshop16-${Date.now()}.psd`, { overwrite: false });
  await service.host.withDocument(inputFile("selected-fixture.psd"), {}, async ({ doc }) => {
    doc.bitsPerChannel = ps.constants.BitsPerChannelType.SIXTEEN;
    await doc.saveAs.psd(fixture16, { embedColorProfile: true }, true);
  });
  checkState();
  entries.push(fixture16);
  for (const name of ["selected-fixture.psd", fixture16.name, "empty-fixture.psd",
    ...(entries.some(entry => entry.name === "all-layer-types.psd") ? ["all-layer-types.psd"] : [])]) {
    console.log("PAGX_HOST_CASE", name);
    const result = await service.run({ input: inputFile(name), folder });
    checkState();
    report.cases.push({ name, passed: true, output: result.file.name, composite: result.composite });
  }
  // Cancel after begin(), while an actual Photoshop document and native session are both owned.
  for (const mode of ["cancel", "native-error"]) {
    const filesBefore = (await folder.getEntries()).map(entry => entry.name).sort().join("\n");
    const signal = {};
    const addRaster = bridge.addRaster;
    if (mode === "native-error") bridge.addRaster = function(id, metadata, pixels) {
      return this.addon.addRaster(id, JSON.stringify({ ...metadata, componentSize: 32 }), pixels.buffer);
    };
    let caught;
    try {
      await service.run({ input: inputFile("selected-fixture.psd"), folder, signal,
        progress(message) { if (mode === "cancel" && message.startsWith("采集图像")) signal.cancelled = true; } });
    } catch (error) { caught = error; }
    finally { bridge.addRaster = addRaster; }
    // executeAsModal rewraps thrown errors; the original JS prototype/message may be unavailable.
    const message = String(caught);
    if (!caught || (mode === "cancel" && (!signal.cancelled || !message.includes("已取消导出"))) ||
        (mode === "native-error" && !message.includes("Invalid capture pixel contract"))) {
      throw new Error(`Unexpected ${mode} result: ${message}`);
    }
    checkState();
    if ((await folder.getEntries()).map(entry => entry.name).sort().join("\n") !== filesBefore) {
      throw new Error("Failed export published or left a temporary output file");
    }
    report.cases.push({ name: mode, passed: true });
  }
  const retry = await service.run({ input: inputFile("selected-fixture.psd"), folder });
  checkState();
  report.cases.push({ name: "retry-after-failures", passed: true, output: retry.file.name });
  const file = await folder.createFile(`host-report-${Date.now()}.json`, { overwrite: false });
  await file.write(JSON.stringify(report, null, 2), { format: uxp.storage.formats.utf8 });
  console.log("PAGX_HOST_SMOKE", JSON.stringify(report));
  return report;
}

module.exports = { run };
