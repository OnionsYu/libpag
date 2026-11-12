/* Copyright (C) 2026 Tencent. Licensed under the Apache License, Version 2.0. */
"use strict";
const ps = require("photoshop");
const { shape, descriptor, action, write, baselineScales, exportFixture } = require("./host-vectors.js");
const unit = value => ({ _unit: "pixelsUnit", _value: value });
const rgb = (red, grain, blue) => ({ _obj: "RGBColor", red, grain, blue });

function gradient(method, options = {}) {
  return { _obj: "gradientLayer", angle: { _unit: "angleUnit", _value: 0 },
    type: { _enum: "gradientType", _value: "linear" }, reverse: false, dither: false,
    align: true, scale: { _unit: "percentUnit", _value: 100 },
    gradient: { _obj: "gradientClassEvent", name: "PAGX test gradient",
      gradientForm: { _enum: "gradientForm", _value: "customStops" }, interfaceIconFrameDimmed: 4096,
      colors: [rgb(230, 40, 70), rgb(20, 170, 220)].map((color, i) => ({ _obj: "colorStop", color,
        type: { _enum: "colorStopType", _value: "userStop" }, location: i * 4096, midpoint: 50 })),
      transparency: [0, 4096].map(location => ({ _obj: "transferSpec", opacity: {
        _unit: "percentUnit", _value: 100 }, location, midpoint: 50 })) },
    gradientsInterpolationMethod: { _enum: "gradientInterpolationMethodType", _value: method }, ...options };
}
async function stroke(layer, { dashed = false, opacity = 100, align = "Center", cap = "Butt", intervals = [3, 2] } = {}) {
  await action({ _obj: "set", _target: [{ _ref: "contentLayer", _id: layer.id }],
    to: { _obj: "shapeStyle", strokeStyle: { _obj: "strokeStyle", strokeStyleVersion: 2,
      strokeEnabled: true, fillEnabled: true, strokeStyleLineWidth: unit(6),
      strokeStyleLineDashOffset: { _unit: "pointsUnit", _value: dashed ? 1 : 0 },
      strokeStyleMiterLimit: 4,
      strokeStyleLineCapType: { _enum: "strokeStyleLineCapType", _value: `strokeStyle${cap}Cap` },
      strokeStyleLineJoinType: { _enum: "strokeStyleLineJoinType", _value: "strokeStyleMiterJoin" },
      strokeStyleLineAlignment: { _enum: "strokeStyleLineAlignment", _value: `strokeStyleAlign${align}` },
      strokeStyleScaleLock: false, strokeStyleStrokeAdjust: false,
      strokeStyleLineDashSet: dashed ? intervals.map(_value => ({ _unit: "pointsUnit", _value })) : [],
      strokeStyleBlendMode: { _enum: "blendMode", _value: "normal" },
      strokeStyleOpacity: { _unit: "percentUnit", _value: opacity },
      strokeStyleContent: { _obj: "solidColorLayer", color: rgb(240, 40, 20) },
      strokeStyleResolution: 72 } } });
}
async function run(folder, exportResults = true) {
  const names = ["solid-fill", "dash", "dash-round", "dash-square", "dash-300", "dash-overlap", "stroke-opacity"];
  for (const name of names) {
    await ps.core.executeAsModal(async context => {
      const doc = await ps.app.createDocument({ name, width: 256, height: 192, resolution: name === "dash-300" ? 300 : 72,
        mode: ps.constants.NewDocumentMode.RGB, fill: ps.constants.DocumentFill.TRANSPARENT,
        profile: "sRGB IEC61966-2.1" });
      await context.hostControl.registerAutoCloseDocument(doc.id);
      let layer;
      if (name === "solid-fill") {
        await action({ _obj: "make", _target: [{ _ref: "contentLayer" }], using: {
          _obj: "contentLayer", type: { _obj: "solidColorLayer", color: rgb(80, 150, 220) } } });
        layer = doc.activeLayers[0];
        await action({ _obj: "delete", _target: [{ _ref: "channel", _enum: "channel", _value: "mask" }] });
        layer.opacity = 75; layer.fillOpacity = 60;
      } else {
        layer = await shape([20, 20, 220, 160], [80, 150, 220]);
        await stroke(layer, { dashed: name.startsWith("dash"), opacity: name === "stroke-opacity" ? 35 : 100,
          cap: ["dash-round", "dash-overlap"].includes(name) ? "Round" : name === "dash-square" ? "Square" : "Butt",
          intervals: ["dash-round", "dash-square"].includes(name) ? [12, 8] : [3, 2] });
        layer.opacity = 75; layer.fillOpacity = 60;
      }
      await write(folder, `${name}.descriptor.json`, await descriptor(doc, layer));
      await baselineScales(doc, folder, name, context);
      await doc.saveAs.psd(await folder.createFile(`${name}.psd`, { overwrite: true }),
        { embedColorProfile: true }, true);
    }, { commandName: "创建 PAGX 填充描边样例", timeOut: 1 });
    if (exportResults) await exportFixture(folder, name);
  }
  console.log("PAGX_APPEARANCE_DONE");
}
module.exports = { run, stroke, gradient };
