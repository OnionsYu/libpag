/* Copyright (C) 2026 Tencent. Licensed under the Apache License, Version 2.0. */
"use strict";
const test = require("node:test");
const assert = require("node:assert/strict");
const { ExportService, ExportCancelled } = require("../src/export-service.js");
const { NativeBridge } = require("../src/native-bridge.js");
const { OutputFiles } = require("../src/output-files.js");
const { PhotoshopHost } = require("../src/photoshop-host.js");

function serviceFixture() {
  const calls = [];
  const host = {
    async withDocument(input, signal, task) {
      calls.push(input.kind);
      try { return await task({}, { reportProgress() {} }); }
      finally { calls.push("close"); }
    },
    async snapshot() { return { name: "sample.psd" }; },
    async prepareCaptures() {},
    async capture(doc, source, request, consume) { consume({}, new Uint8Array([1, 2, 3, 255])); }
  };
  const bridge = {
    begin() { return { sessionId: "s", composite: true, rasters: [{ id: "composite" }], diagnostics: [] }; },
    addRaster() { calls.push("pixels"); },
    finish() { calls.push("finish"); return "<pagx/>"; },
    discard() { calls.push("discard"); }
  };
  const files = { async publish() { calls.push("publish"); return { name: "sample.pagx" }; } };
  return { calls, host, bridge, files, service: new ExportService({ host, bridge, files }) };
}

test("both input modes share a transaction and close host documents before publication", async () => {
  for (const kind of ["current", "file"]) {
    const f = serviceFixture();
    const result = await f.service.run({ input: { kind }, folder: {} });
    assert.equal(result.file.name, "sample.pagx");
    assert.deepEqual(f.calls, [kind, "pixels", "close", "finish", "publish"]);
  }
});

test("capture failure and cancellation discard native session and never publish", async () => {
  for (const cancel of [false, true]) {
    const f = serviceFixture();
    const signal = {};
    f.host.capture = async () => {
      if (cancel) signal.cancelled = true;
      else throw new Error("capture failed");
    };
    await assert.rejects(f.service.run({ input: { kind: "current" }, signal }),
      cancel ? ExportCancelled : /capture failed/);
    assert.deepEqual(f.calls, ["current", "close", "discard"]);
    assert.equal(f.service.busy, false);
  }
});

test("native validation failure never publishes and allows a later export", async () => {
  const f = serviceFixture();
  f.bridge.finish = () => { throw new Error("bad xml"); };
  await assert.rejects(f.service.run({ input: { kind: "current" } }), /bad xml/);
  assert.equal(f.service.busy, false);
  assert(!f.calls.includes("publish"));
  assert(f.calls.includes("discard"));
});

test("simultaneous exports are rejected", async () => {
  const f = serviceFixture();
  let resume;
  f.host.snapshot = () => new Promise(resolve => { resume = resolve; });
  const first = f.service.run({ input: { kind: "current" } });
  await assert.rejects(f.service.run({ input: { kind: "file" } }), /已有导出/);
  resume({ name: "sample" });
  await first;
});

test("bridge sends only the view bytes, not adjacent pixels in its backing buffer", () => {
  let captured;
  const bridge = new NativeBridge({ info: () => '{"protocolVersion":1,"schemaVersion":1}',
    addRaster: (id, metadata, bytes) => { captured = new Uint8Array(bytes); } });
  bridge.addRaster("s", { componentSize: 8 }, new Uint8Array([99, 1, 2, 3, 4, 88]).subarray(1, 5));
  assert.deepEqual([...captured], [1, 2, 3, 4]);
  bridge.addRaster("s", { componentSize: 16 }, new Uint16Array([99, 65535, 0, 32895, 88]).subarray(1, 4));
  assert.deepEqual([...new Uint16Array(captured.buffer)], [65535, 0, 32895]);
  assert.throws(() => bridge.addRaster("s", { componentSize: 8 }, new Uint16Array(4)), /位深不匹配/);
  assert.throws(() => new NativeBridge({ info: () => '{"protocolVersion":2}' }), /不匹配/);
});

function storageFixture(signal = {}) {
  const calls = [];
  const temp = {
    nativePath: "/out/temp",
    async write(xml) { calls.push(["write", xml]); },
    async moveTo(folder, options) { calls.push(["move", options]); this.nativePath = `/out/${options.newName}`; },
    async delete() { calls.push(["delete"]); }
  };
  const folder = {
    async getEntries() { return [{ name: "Sample.pagx" }]; },
    async createFile(name, options) { assert.equal(options.overwrite, false); return temp; }
  };
  return { calls, temp, folder, signal, files: new OutputFiles("utf8") };
}

test("publish uses a new name and a complete temporary file, preserving existing exports", async () => {
  const f = storageFixture();
  const result = await f.files.publish(f.folder, "Sample.psd", "XML", f.signal);
  assert.equal(result.nativePath, "/out/Sample (2).pagx");
  assert.deepEqual(f.calls, [["write", "XML"], ["move", { newName: "Sample (2).pagx", overwrite: false }]]);
});

test("write errors, a racing filename and cancellation remove the temp without overwriting", async () => {
  for (const failure of ["write", "move", "cancel"]) {
    const f = storageFixture();
    if (failure === "cancel") f.temp.write = async () => { f.signal.cancelled = true; };
    else f.temp[failure === "write" ? "write" : "moveTo"] = async () => { throw new Error(failure); };
    await assert.rejects(f.files.publish(f.folder, "Sample.psd", "XML", f.signal));
    assert.deepEqual(f.calls.at(-1), ["delete"]);
  }
});

function hostFixture() {
  const original = { id: 1, name: "unsaved.psd", saved: false, history: ["paint"],
    activeLayers: [17], async duplicate() { ps.app.documents.push(copy); ps.app.activeDocument = copy; return copy; } };
  const copy = { id: 2, name: "temp" };
  const ps = { app: { documents: [original], activeDocument: original, async open() {
    ps.app.documents.push(copy); ps.app.activeDocument = copy; return copy;
  } }, core: { async executeAsModal(fn, options) {
    assert.equal(options.timeOut, 1, "modal acquisition timeout is expressed in seconds");
    const closing = [];
    try { return await fn({ hostControl: { async registerAutoCloseDocument(id) { closing.push(id); } } }); }
    finally { ps.app.documents = ps.app.documents.filter(doc => !closing.includes(doc.id)); }
  } } };
  return { original, copy, ps, host: new PhotoshopHost(ps, 64 * 1024 * 1024) };
}

test("current document is duplicated and original dirty state, selection and history survive errors", async () => {
  const f = hostFixture();
  await assert.rejects(f.host.withDocument({ kind: "current", documentId: 1 }, {}, async ({ doc }) => {
    assert.equal(doc.id, 2);
    throw new Error("stop");
  }), /stop/);
  assert.deepEqual(f.ps.app.documents, [f.original]);
  assert.equal(f.ps.app.activeDocument, f.original);
  assert.equal(f.original.saved, false);
  assert.deepEqual(f.original.history, ["paint"]);
  assert.deepEqual(f.original.activeLayers, [17]);
});

test("selected PSD is owned only when newly opened; an already-open document is never closed", async () => {
  const f = hostFixture();
  await f.host.withDocument({ kind: "file", file: { name: "input.psd" } }, {}, async ({ doc }) => {
    assert.equal(doc.id, 2);
  });
  assert.deepEqual(f.ps.app.documents, [f.original]);
  f.ps.app.open = async () => f.original;
  await assert.rejects(f.host.withDocument({ kind: "file", file: {} }, {}, () => {}), /已在 Photoshop/);
  assert.deepEqual(f.ps.app.documents, [f.original]);
});

test("capture uses returned bounds and disposes imaging resources even when native conversion throws", async () => {
  let disposed = false;
  const host = new PhotoshopHost({ imaging: { async getPixels(options) {
    assert.equal(options.applyAlpha, false);
    assert.equal(options.layerID, undefined);
    assert.equal(options.componentSize, 8);
    assert.deepEqual(options.sourceBounds, { left: 0, top: 0, right: 10, bottom: 20 });
    return { level: 0, sourceBounds: { left: -3, top: 8, right: -1, bottom: 9 }, imageData: {
      width: 2, height: 1, colorSpace: "RGB", colorProfile: "sRGB IEC61966-2.1", componentSize: 8, components: 4,
      async getData() { return new Uint8Array(8); }, dispose() { disposed = true; }
    } };
  } } }, 100);
  await assert.rejects(host.capture({ doc: { id: 1 } },
    { bitDepth: 8, bounds: { left: 0, top: 0, width: 10, height: 20 } }, { id: "composite" }, metadata => {
    assert.equal(metadata.left, -3);
    assert.equal(metadata.top, 8);
    throw new Error("encoding failed");
  }), /encoding failed/);
  assert.equal(disposed, true);
});

test("snapshot preserves the UXP host version and handles nested empty Photoshop groups", async () => {
  const constants = { BitsPerChannelType: { EIGHT: 8, SIXTEEN: 16 },
    LayerKind: { GROUP: "group", NORMAL: "pixel" },
    BlendMode: { NORMAL: "normal", PASSTHROUGH: "passThrough" }, DocumentMode: { RGB: "RGB" } };
  const group = { id: 2, name: "empty", kind: "group", visible: true, opacity: 100,
    fillOpacity: 100, blendMode: "passThrough", bounds: {}, layers: [] };
  const pixel = { ...group, id: 3, name: "pixels", kind: "pixel",
    bounds: { left: 7, top: 9, right: 37, bottom: 31 } };
  const doc = { bitsPerChannel: 8, pixelAspectRatio: 1, artboards: [], width: 64, height: 48,
    resolution: 72, mode: "RGB", colorProfileName: "sRGB IEC61966-2.1",
    layers: [pixel, { ...group, id: 1, layers: [group] }] };
  const host = new PhotoshopHost({ constants }, 10000, "27.4.0");
  const snapshot = await host.snapshot({ doc, sourceName: "input.psd" });
  assert.equal(snapshot.hostVersion, "27.4.0");
  assert.deepEqual(snapshot.layers.map(layer => layer.id), ["ps:1", "ps:3"]);
  assert.deepEqual(snapshot.layers[0].bounds, { left: 0, top: 0, width: 0, height: 0 });
  assert.deepEqual(snapshot.layers[0].children[0].bounds, snapshot.layers[0].bounds);
  assert.deepEqual(snapshot.layers[1].bounds, { left: 7, top: 9, width: 30, height: 22 });
  pixel.bounds = {};
  await assert.rejects(host.snapshot({ doc }), /无效的图层边界/);
  doc.artboards = [1];
  await assert.rejects(host.snapshot({ doc }), /画板文档/);
  doc.pixelAspectRatio = 2;
  await assert.rejects(host.snapshot({ doc }), /方形像素/);
});

test("panel waits for the real asynchronous Hybrid loader before enabling export", async () => {
  const vm = require("node:vm");
  const fs = require("node:fs");
  const elements = Object.fromEntries(["status", "details", "current", "file", "cancel"].map(id =>
    [id, { disabled: true, addEventListener() {} }]));
  let resolveAddon;
  const addon = new Promise(resolve => { resolveAddon = resolve; });
  const errors = [];
  const imports = { photoshop: {}, uxp: { storage: { formats: { utf8: "utf8" } }, host: { version: "27.4.0" } },
    "pagx-photoshop.uxpaddon": addon };
  vm.runInNewContext(fs.readFileSync(require.resolve("../main.js"), "utf8"), {
    document: { getElementById: id => elements[id] }, console: { error: error => errors.push(error) },
    require: name => name in imports ? imports[name] : require(`../${name}`)
  });
  assert.equal(elements.current.disabled, true);
  resolveAddon({ info: () => JSON.stringify({ protocolVersion: 1, schemaVersion: 1, maxPixels: 100 }) });
  await new Promise(resolve => setImmediate(resolve));
  assert.deepEqual(errors, []);
  assert.equal(elements.current.disabled, false);
  assert.equal(elements.file.disabled, false);
  assert.equal(elements.status.textContent, "准备就绪");
});
