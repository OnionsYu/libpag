/* Copyright (C) 2026 Tencent. Licensed under the Apache License, Version 2.0. */
"use strict";
const ps = require("photoshop");
const { shape, action, descriptor, write, baseline, baselineScales, exportFixture } = require("./host-vectors.js");
const { stroke } = require("./host-appearance.js");
async function run(folder, exportResults = true, names = ["vector-mask", "vector-mask-disabled",
  "vector-mask-feather", "vector-mask-density", "vector-mask-raster-child", "vector-mask-pass-through"]) {
  for (const name of names) {
    await ps.core.executeAsModal(async context => {
      const doc = await ps.app.createDocument({ name, width: 256, height: 192, resolution: 72,
        mode: ps.constants.NewDocumentMode.RGB, fill: ps.constants.DocumentFill.TRANSPARENT,
        profile: "sRGB IEC61966-2.1" });
      await context.hostControl.registerAutoCloseDocument(doc.id);
      if (name === "vector-mask-pass-through") await shape([0, 0, 230, 180], [30, 160, 75]);
      const background = await shape([5, 5, 240, 185], [70, 130, 200]);
      const foreground = await shape([40, 30, 220, 155], [230, 90, 40]);
      foreground.opacity = 70;
      if (name === "vector-mask-raster-child") await stroke(foreground, { opacity: 35 });
      if (name === "vector-mask-pass-through") foreground.blendMode = ps.constants.BlendMode.MULTIPLY;
      const group = await doc.createLayerGroup({ name: "Masked vectors", fromLayers: [background, foreground],
        blendMode: name === "vector-mask-pass-through" ? ps.constants.BlendMode.PASSTHROUGH : ps.constants.BlendMode.NORMAL });
      if (name !== "vector-mask-pass-through") group.opacity = 65;
      const paths = [[[20.25, 20.75], [210.5, 27.25], [201.75, 169.5], [18.25, 159.75]],
        [[65, 50], [160, 50], [160, 130], [65, 130]]].map((anchors, i) => {
        const path = new ps.app.SubPathInfo();
        path.closed = true;
        path.operation = i ? ps.constants.ShapeOperation.SHAPESUBTRACT : ps.constants.ShapeOperation.SHAPEADD;
        path.entireSubPath = anchors.map(anchor => {
          const p = new ps.app.PathPointInfo();
          p.anchor = anchor; p.leftDirection = anchor; p.rightDirection = anchor;
          p.kind = ps.constants.PointKind.CORNERPOINT;
          return p;
        });
        return path;
      });
      await doc.pathItems.add("PAGX mask path", paths);
      await action({ _obj: "select", _target: [{ _ref: "path", _name: "PAGX mask path" }] });
      await action({ _obj: "make", _target: [{ _ref: "path" }],
        at: { _ref: "path", _enum: "path", _value: "vectorMask" },
        using: { _ref: "path", _enum: "ordinal", _value: "targetEnum" } });
      if (name === "vector-mask-disabled") await action({ _obj: "set", _target: [{ _ref: "layer", _id: group.id }],
        to: { _obj: "layer", vectorMaskEnabled: false } });
      if (name === "vector-mask-feather") await action({ _obj: "set", _target: [{ _ref: "layer", _id: group.id }],
        to: { _obj: "layer", vectorMaskFeather: { _unit: "pixelsUnit", _value: 3 } } });
      if (name === "vector-mask-density") await action({ _obj: "set", _target: [{ _ref: "layer", _id: group.id }],
        to: { _obj: "layer", vectorMaskDensity: { _unit: "percentUnit", _value: 50 } } });
      const raw = await descriptor(doc, group);
      if (!raw.value.hasVectorMask) throw new Error("Photoshop did not create the vector mask");
      await write(folder, `${name}.descriptor.json`, raw);
      await shape([222, 15, 250, 40], [180, 20, 100], "ellipse");
      if (["vector-mask", "vector-mask-disabled"].includes(name)) await baselineScales(doc, folder, name, context);
      else await baseline(doc, folder, name);
      await doc.saveAs.psd(await folder.createFile(`${name}.psd`, { overwrite: true }),
        { embedColorProfile: true }, true);
    }, { commandName: "创建 PAGX 矢量蒙版样例", timeOut: 1 });
    if (exportResults) {
      const result = await exportFixture(folder, name);
      const captures = ["vector-mask", "vector-mask-disabled"].includes(name) ? 0 : 1;
      if (result.composite || result.captures.length !== captures || result.statistics.vectorShapes < 1)
        throw new Error(`${name}: incorrect mask scope or unrelated vector loss`);
    }
  }
  console.log("PAGX_MASKS_DONE");
}
async function inverted(folder, name = "vector-mask-inverted") {
  await ps.core.executeAsModal(async context => {
    const doc = await ps.app.open(await folder.getEntry(`${name}.psd`));
    await context.hostControl.registerAutoCloseDocument(doc.id);
    // The fixture changes layer metadata without updating the merged preview. Force a real
    // layered composite before recording the independent reference.
    for (const layer of doc.layers) { const visible = layer.visible; layer.visible = !visible; layer.visible = visible; }
    if (name === "vector-mask-inverted") await baselineScales(doc, folder, name, context);
    else await baseline(doc, folder, name);
  }, { commandName: "验证反相蒙版样例", timeOut: 1 });
  return exportFixture(folder, name);
}
module.exports = { run, inverted };
