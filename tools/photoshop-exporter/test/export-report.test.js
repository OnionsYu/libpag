/* Copyright (C) 2026 Tencent. Licensed under the Apache License, Version 2.0. */
"use strict";
const test = require("node:test");
const assert = require("node:assert/strict");
const { makeExportReport, summarizeReport } = require("../src/export-report.js");
const { OutputFiles } = require("../src/output-files.js");

test("reports name replacement roots and their parent without exposing source text or PSD paths", () => {
  const source = { name: "示例.psd", schemaVersion: 2, hostVersion: "27.4.0",
    layers: [{ id: "g", name: "Group", children: [
      { id: "base", name: "基底", children: [], text: "source text" },
      { id: "clip", name: "剪贴", children: [] }] }] };
  const plan = { composite: false,
    statistics: { vectorShapes: 2, outlinedTexts: 1, sourceImages: 0, fallbackScopes: 1 },
    rasters: [{ id: "range:base", parentId: "g", layerIds: ["base", "clip"] }],
    diagnostics: [{ code: "clipping-range", layerId: "clip", message: "clipped" }] };
  const report = makeExportReport(source, plan, false, { version: "0.2.0", protocolVersion: 2 });
  assert.deepEqual(report.captures[0].layers, [{ id: "base", name: "基底" }, { id: "clip", name: "剪贴" }]);
  assert.deepEqual(report.captures[0].parent, { id: "g", name: "Group" });
  assert.equal(report.diagnostics[0].layerName, "剪贴");
  assert.equal(JSON.stringify(report).includes("source text"), false);
  assert.match(summarizeReport(report), /矢量形状 2/);
  assert.match(summarizeReport(report), /剪贴：剪贴图层与基底/);
  const composite = makeExportReport(source, { ...plan, composite: true,
    rasters: [{ id: "composite", parentId: "", layerIds: [] }] }, true, {});
  assert.equal(composite.requestedMode, "appearance");
  assert.equal(composite.captures[0].scope, "document");
  assert.match(summarizeReport(composite), /整幅位图/);
});

function storage(failure) {
  const files = new Map([["Sample.pagx.report.json", "existing report"]]);
  const signal = {};
  const check = phase => { if (phase === failure) throw new Error(phase); };
  const folder = {
    async getEntries() { return [...files.keys()].map(name => ({ name })); },
    async createFile(name) {
      const isReport = name.endsWith(".report");
      check(isReport ? "create-report" : "create-pagx");
      assert(!files.has(name));
      files.set(name, "");
      return {
        name,
        get nativePath() { return `/out/${this.name}`; },
        async write(data) {
          check(isReport ? "write-report" : "write-pagx");
          files.set(this.name, data);
        },
        async moveTo(parent, options) {
          assert.equal(options.overwrite, false);
          check(isReport ? "rename-report" : "rename-pagx");
          assert(!files.has(options.newName));
          files.set(options.newName, files.get(this.name));
          files.delete(this.name);
          this.name = options.newName;
          if (isReport && failure === "cancel-after-report") signal.cancelled = true;
        },
        async delete() { files.delete(this.name); }
      };
    }
  };
  return { files, signal, folder };
}

test("PAGX and report share a collision-free name and include the exact published filename", async () => {
  const f = storage();
  const result = await new OutputFiles("utf8").publish(f.folder, "Sample.psd", "<pagx/>", f.signal, { reportVersion: 1 });
  assert.equal(result.name, "Sample (2).pagx");
  assert.equal(result.reportPath, "/out/Sample (2).pagx.report.json");
  assert.equal(f.files.size, 3);
  assert.equal(f.files.get("Sample.pagx.report.json"), "existing report");
  assert.equal(f.files.get(result.name), "<pagx/>");
  assert.deepEqual(JSON.parse(f.files.get(`${result.name}.report.json`)),
    { reportVersion: 1, outputFile: result.name });
});

test("failed publication and cancellation roll back both owned files, including a renamed report", async () => {
  for (const phase of ["create-report", "write-report", "rename-report", "rename-pagx", "cancel-after-report"]) {
    const f = storage(phase);
    await assert.rejects(new OutputFiles("utf8").publish(f.folder, "Sample.psd", "<pagx/>", f.signal, {}));
    assert.deepEqual([...f.files], [["Sample.pagx.report.json", "existing report"]], phase);
  }
});
