/* Copyright (C) 2026 Tencent. Licensed under the Apache License, Version 2.0. */
"use strict";
const ps = require("photoshop");
const { shape, write, descriptor, baselineScales, exportFixture } = require("./host-vectors.js");
async function run(folder) {
  for (const mode of ["MULTIPLY", "SCREEN", "DARKEN", "LIGHTEN", "OVERLAY", "HARDLIGHT", "SOFTLIGHT", "EXCLUSION"]) {
    if (ps.constants.BlendMode[mode] === undefined) throw new Error(`Unknown blend enum: ${mode}`);
    const name = `blend-${mode.toLowerCase()}`;
    await ps.core.executeAsModal(async context => {
      const doc = await ps.app.createDocument({ name, width: 256, height: 192, resolution: 72,
        mode: ps.constants.NewDocumentMode.RGB, fill: ps.constants.DocumentFill.TRANSPARENT,
        profile: "sRGB IEC61966-2.1" });
      await context.hostControl.registerAutoCloseDocument(doc.id);
      const background = await shape([10, 10, 190, 180], [70, 130, 200]);
      background.opacity = 65;
      const foreground = await shape([60, 30, 230, 150], [230, 90, 40]);
      foreground.opacity = 70; foreground.fillOpacity = 60;
      foreground.blendMode = ps.constants.BlendMode[mode];
      await write(folder, `${name}.descriptor.json`, await descriptor(doc, foreground));
      await baselineScales(doc, folder, name, context);
      await doc.saveAs.psd(await folder.createFile(`${name}.psd`, { overwrite: true }),
        { embedColorProfile: true }, true);
    }, { commandName: "创建 PAGX 原生混合样例", timeOut: 1 });
    const result = await exportFixture(folder, name);
    const captures = mode === "SOFTLIGHT" ? 1 : 0;
    if (result.captures.length !== captures) throw new Error(`${name}: incorrect blend qualification`);
  }
  console.log("PAGX_BLENDS_DONE");
}
module.exports = { run };
