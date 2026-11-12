/* Copyright (C) 2026 Tencent. Licensed under the Apache License, Version 2.0. */
"use strict";
const ps = require("photoshop");
const { gradient } = require("./host-appearance.js");
const { action, descriptor, write, baselineScales, exportFixture } = require("./host-vectors.js");
async function run(folder, exportResults = true) {
  for (const method of ["classic", "linear", "perceptual"]) {
    const name = `gradient-${method}`;
    await ps.core.executeAsModal(async context => {
      const doc = await ps.app.createDocument({ name, width: 256, height: 192, resolution: 72,
        mode: ps.constants.NewDocumentMode.RGB, fill: ps.constants.DocumentFill.TRANSPARENT,
        profile: "sRGB IEC61966-2.1" });
      await context.hostControl.registerAutoCloseDocument(doc.id);
      const unit = _value => ({ _unit: "pixelsUnit", _value });
      await action({ _obj: "make", _target: [{ _ref: "contentLayer" }], using: {
        _obj: "contentLayer", type: gradient(method), shape: { _obj: "rectangle",
          top: unit(20), left: unit(20), bottom: unit(160), right: unit(220) } } });
      const layer = doc.activeLayers[0];
      await write(folder, `${name}.descriptor.json`, await descriptor(doc, layer));
      await baselineScales(doc, folder, name, context);
      await doc.saveAs.psd(await folder.createFile(`${name}.psd`, { overwrite: true }),
        { embedColorProfile: true }, true);
    }, { commandName: "创建 PAGX 渐变样例", timeOut: 1 });
    if (exportResults) await exportFixture(folder, name);
  }
  console.log("PAGX_GRADIENTS_DONE");
}
module.exports = { run };
