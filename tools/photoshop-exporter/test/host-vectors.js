/* Copyright (C) 2026 Tencent. Licensed under the Apache License, Version 2.0. */
"use strict";
// Run inside UDT. Generated fixtures and independent Photoshop references stay outside the plugin.
const ps = require("photoshop");
const uxp = require("uxp");
const unit = value => ({ _unit: "pixelsUnit", _value: value });
async function action(descriptor) {
  const [result] = await ps.action.batchPlay([{ ...descriptor,
    _options: { dialogOptions: "silent" } }], {});
  if (result._obj === "error") throw new Error(JSON.stringify(result));
  return result;
}
async function shape(bounds, color = [230, 40, 70], type = "rectangle") {
  await action({ _obj: "make", _target: [{ _ref: "contentLayer" }],
    using: { _obj: "contentLayer", type: { _obj: "solidColorLayer",
      color: { _obj: "RGBColor", red: color[0], grain: color[1], blue: color[2] } },
    shape: { _obj: type, top: unit(bounds[1]), left: unit(bounds[0]),
      bottom: unit(bounds[3]), right: unit(bounds[2]) } } });
  return ps.app.activeDocument.activeLayers[0];
}
async function descriptor(doc, layer) {
  const target = [{ _ref: "layer", _id: layer.id }, { _ref: "document", _id: doc.id }];
  const value = await action({ _obj: "get", _target: target });
  let path;
  if (value.hasVectorMask) {
    path = await action({ _obj: "get", _target: [
      { _ref: "path", _enum: "path", _value: "vectorMask" }, ...target] });
  }
  return { id: layer.id, kind: layer.kind, value, path };
}
async function write(folder, name, value) {
  const file = await folder.createFile(name, { overwrite: true });
  await file.write(typeof value === "string" ? value : JSON.stringify(value, null, 2),
    { format: uxp.storage.formats.utf8 });
  return file;
}
async function baseline(doc, folder, name) {
  const result = await ps.imaging.getPixels({ documentID: doc.id,
    sourceBounds: { left: 0, top: 0, right: doc.width, bottom: doc.height },
    colorSpace: "RGB", colorProfile: "sRGB IEC61966-2.1", componentSize: 8, applyAlpha: false });
  try {
    const data = result.imageData;
    const file = await folder.createFile(`${name}.rgba`, { overwrite: true });
    await file.write(await data.getData({ chunky: true }), { format: uxp.storage.formats.binary });
    await write(folder, `${name}.reference.json`, { width: doc.width, height: doc.height,
      sourceBounds: result.sourceBounds, pixelWidth: data.width, pixelHeight: data.height,
      components: data.components, level: result.level });
  } finally { result.imageData.dispose(); }
}
async function probe(folder) {
  folder = folder || await uxp.storage.localFileSystem.getTemporaryFolder();
  await ps.core.executeAsModal(async context => {
    const doc = await ps.app.createDocument({ name: "PAGX vector fixture", width: 256, height: 192,
      resolution: 72, mode: ps.constants.NewDocumentMode.RGB,
      fill: ps.constants.DocumentFill.TRANSPARENT, profile: "sRGB IEC61966-2.1" });
    await context.hostControl.registerAutoCloseDocument(doc.id);
    const initial = await descriptor(doc, doc.layers[0]);
    const rectangle = await shape([12, 18, 110, 80]);
    const rectangleData = await descriptor(doc, rectangle);
    const properties = {};
    for (const property of ["blendOptions", "blendRange", "knockout", "blendInterior",
      "blendClipped", "transparencyShapesLayer", "layerEffects", "AGMStrokeStyleInfo"]) {
      properties[property] = (await ps.action.batchPlay([{ _obj: "get", _target: [
        { _property: property }, { _ref: "layer", _id: rectangle.id },
        { _ref: "document", _id: doc.id }], _options: { dialogOptions: "silent" } }], {}))[0];
    }
    const ellipse = await shape([120.5, 30.25, 230.5, 140.25], [20, 170, 220], "ellipse");
    const ellipseData = await descriptor(doc, ellipse);
    await baseline(doc, folder, "shapes");
    const file = await folder.createFile("shapes.psd", { overwrite: true });
    await doc.saveAs.psd(file, { embedColorProfile: true }, true);
    const text = await doc.createTextLayer({ name: "Text", contents: "PAGX 中文", fontSize: 24,
      position: { x: 12, y: 175 } });
    const textData = await descriptor(doc, text);
    await text.textItem.convertToShape();
    const outlinedData = await descriptor(doc, doc.activeLayers[0]);
    await write(folder, "descriptor-probe.json", { hostVersion: uxp.host.version,
      initial, rectangle: rectangleData, ellipse: ellipseData, text: textData, outlined: outlinedData,
      properties });
  }, { commandName: "验证 PAGX 矢量描述符", timeOut: 1 });
  console.log("PAGX_DESCRIPTOR_PROBE", folder.nativePath);
}
async function exportFixture(folder, name = "shapes") {
  const { PhotoshopHost } = require("../src/photoshop-host.js");
  const { NativeBridge } = require("../src/native-bridge.js");
  const { ExportService } = require("../src/export-service.js");
  const { OutputFiles } = require("../src/output-files.js");
  const bridge = new NativeBridge(await require("pagx-photoshop.uxpaddon"));
  const host = new PhotoshopHost(ps, bridge.info.maxPixels, uxp.host.version);
  const snapshot = host.snapshot.bind(host);
  host.snapshot = async document => {
    const source = await snapshot(document);
    await write(folder, `${name}.snapshot.json`, source);
    return source;
  };
  const service = new ExportService({ host, bridge, files: new OutputFiles(uxp.storage.formats.utf8) });
  const before = Array.from(ps.app.documents).map(doc => doc.id).join(",");
  const result = await service.run({ input: { kind: "file", file: await folder.getEntry(`${name}.psd`) }, folder });
  if (Array.from(ps.app.documents).map(doc => doc.id).join(",") !== before) throw new Error("Leaked document");
  await write(folder, `${name}.export.json`, { output: result.file.name, composite: result.composite,
    diagnostics: result.diagnostics });
  console.log("PAGX_VECTOR_EXPORT", result.file.nativePath, result.composite, JSON.stringify(result.diagnostics));
  return result;
}
async function geometryFixtures(folder) {
  for (const name of ["hole", "stroke", "transformed"]) {
    await ps.core.executeAsModal(async context => {
      const doc = await ps.app.createDocument({ name, width: 256, height: 192, resolution: 72,
        mode: ps.constants.NewDocumentMode.RGB, fill: ps.constants.DocumentFill.TRANSPARENT,
        profile: "sRGB IEC61966-2.1" });
      await context.hostControl.registerAutoCloseDocument(doc.id);
      let layer = await shape([20.5, 20.25, 200.5, 140.25], [80, 150, 220]);
      if (name === "hole") {
        const paths = [[[20.5, 20.25], [200.5, 20.25], [200.5, 140.25], [20.5, 140.25]],
          [[60, 50], [160, 50], [160, 110], [60, 110]]].map((anchors, index) => {
          const path = new ps.app.SubPathInfo();
          path.closed = true;
          path.operation = index ? ps.constants.ShapeOperation.SHAPESUBTRACT : ps.constants.ShapeOperation.SHAPEADD;
          path.entireSubPath = anchors.map(anchor => {
            const p = new ps.app.PathPointInfo();
            p.anchor = anchor; p.leftDirection = anchor; p.rightDirection = anchor;
            p.kind = ps.constants.PointKind.CORNERPOINT;
            return p;
          });
          return path;
        });
        await doc.pathItems.add("PAGX compound path", paths);
        // Photoshop cannot delete the last layer. Hide the construction helper explicitly.
        layer.visible = false;
        await action({ _obj: "select", _target: [{ _ref: "path", _name: "PAGX compound path" }] });
        await action({ _obj: "make", _target: [{ _ref: "contentLayer" }], using: {
          _obj: "contentLayer", type: { _obj: "solidColorLayer",
            color: { _obj: "RGBColor", red: 80, grain: 150, blue: 220 } } } });
        layer = doc.activeLayers[0];
        if ((await descriptor(doc, layer)).path.pathContents.pathComponents.length !== 2)
          throw new Error("Photoshop did not create the compound-path fixture");
      } else if (name === "stroke") {
        await action({ _obj: "set", _target: [{ _ref: "contentLayer", _id: layer.id }],
          to: { _obj: "shapeStyle", strokeStyle: { _obj: "strokeStyle", strokeStyleVersion: 2,
            strokeEnabled: true, fillEnabled: true, strokeStyleLineWidth: unit(6),
            strokeStyleLineDashOffset: { _unit: "pointsUnit", _value: 0 }, strokeStyleMiterLimit: 4,
            strokeStyleLineCapType: { _enum: "strokeStyleLineCapType", _value: "strokeStyleRoundCap" },
            strokeStyleLineJoinType: { _enum: "strokeStyleLineJoinType", _value: "strokeStyleRoundJoin" },
            strokeStyleLineAlignment: { _enum: "strokeStyleLineAlignment", _value: "strokeStyleAlignInside" },
            strokeStyleScaleLock: false, strokeStyleStrokeAdjust: false, strokeStyleLineDashSet: [],
            strokeStyleBlendMode: { _enum: "blendMode", _value: "normal" },
            strokeStyleOpacity: { _unit: "percentUnit", _value: 100 },
            strokeStyleContent: { _obj: "solidColorLayer", color: { _obj: "RGBColor", red: 240, grain: 40, blue: 20 } },
            strokeStyleResolution: 72 } } });
        layer.opacity = 75; layer.fillOpacity = 60;
      } else {
        await layer.rotate(17, ps.constants.AnchorPosition.MIDDLECENTER);
        await layer.translate(-25.5, 6.25);
      }
      await write(folder, `${name}.descriptor.json`, await descriptor(doc, layer));
      await baseline(doc, folder, name);
      const file = await folder.createFile(`${name}.psd`, { overwrite: true });
      await doc.saveAs.psd(file, { embedColorProfile: true }, true);
    }, { commandName: "创建 PAGX 几何验证样例", timeOut: 1 });
    await exportFixture(folder, name);
  }
  console.log("PAGX_GEOMETRY_FIXTURES_DONE");
}
module.exports = { probe, exportFixture, geometryFixtures, shape, descriptor, action, write, baseline };
