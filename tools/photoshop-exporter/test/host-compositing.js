/* Copyright (C) 2026 Tencent. Licensed under the Apache License, Version 2.0. */
"use strict";
const ps = require("photoshop");
const { shape, descriptor, action, write, baseline, exportFixture } = require("./host-vectors.js");

async function run(folder) {
  for (const name of ["isolated-group", "pass-through", "clipping", "group-mask", "adjustment"]) {
    await ps.core.executeAsModal(async context => {
      const doc = await ps.app.createDocument({ name, width: 256, height: 192, resolution: 72,
        mode: ps.constants.NewDocumentMode.RGB, fill: ps.constants.DocumentFill.TRANSPARENT,
        profile: "sRGB IEC61966-2.1" });
      await context.hostControl.registerAutoCloseDocument(doc.id);
      await shape([5, 5, 180, 180], [70, 130, 200]);
      const base = await shape([30, 25, 160, 145], [230, 120, 40]);
      base.opacity = 60;
      const top = await shape([65, 55, 210, 160], [40, 180, 100]);
      top.opacity = 75;
      if (name === "clipping") {
        top.isClippingMask = true;
      } else {
        let members = [base, top];
        if (name === "adjustment") {
          await action({ _obj: "make", _target: [{ _ref: "adjustmentLayer" }],
            using: { _obj: "adjustmentLayer", type: { _obj: "invert" } } });
          members.push(doc.activeLayers[0]);
        } else top.blendMode = ps.constants.BlendMode.MULTIPLY;
        const group = await doc.createLayerGroup({ name: "Composite boundary", fromLayers: members,
          blendMode: name === "pass-through" ? ps.constants.BlendMode.PASSTHROUGH : ps.constants.BlendMode.NORMAL });
        if (name !== "pass-through") group.opacity = 50;
        if (name === "group-mask") {
          await doc.selection.selectRectangle({ left: 40, top: 35, right: 180, bottom: 135 });
          await action({ _obj: "make", new: { _class: "channel" },
            at: { _ref: "channel", _enum: "channel", _value: "mask" },
            using: { _enum: "userMaskEnabled", _value: "revealSelection" } });
          await doc.selection.deselect();
        }
        await write(folder, `${name}.group-descriptor.json`, await descriptor(doc, group));
      }
      const foreground = await shape([190, 15, 245, 65], [200, 40, 180], "ellipse");
      foreground.name = "Unrelated vector above";
      await baseline(doc, folder, name);
      const file = await folder.createFile(`${name}.psd`, { overwrite: true });
      await doc.saveAs.psd(file, { embedColorProfile: true }, true);
    }, { commandName: "创建 PAGX 合成依赖样例", timeOut: 1 });
    const result = await exportFixture(folder, name);
    if (result.composite || result.statistics.vectorShapes < 1 || result.captures.length !== 1)
      throw new Error(`${name}: unrelated vectors were collapsed or captures were duplicated`);
  }
  console.log("PAGX_COMPOSITING_FIXTURES_DONE");
}
module.exports = { run };
