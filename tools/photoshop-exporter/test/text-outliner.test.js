/* Copyright (C) 2026 Tencent. Licensed under the Apache License, Version 2.0. */
"use strict";
const test = require("node:test");
const assert = require("node:assert/strict");
const { outlineEligibility, outlineText } = require("../src/text-outliner.js");
const fixture = require("./fixtures/ps27-descriptors.json");
function textDescriptor() {
  return { textWarningLevel: 0, textKey: { textStyleRange: [{ textStyle: {
    fontAvailable: true, fill: true, stroke: false,
    color: { _obj: "RGBColor", red: 0, grain: 0, blue: 0 }
  } }] } };
}
test("outline eligibility rejects missing fonts, multicolor runs and warped text", () => {
  const raw = textDescriptor();
  assert.equal(outlineEligibility(raw), undefined);
  raw.textKey.textStyleRange[0].textStyle.fontAvailable = false;
  assert.equal(outlineEligibility(raw), "missing-font");
  raw.textKey.textStyleRange[0].textStyle.fontAvailable = true;
  const range = structuredClone(raw.textKey.textStyleRange[0]);
  range.textStyle.color.red = 255;
  raw.textKey.textStyleRange.push(range);
  assert.equal(outlineEligibility(raw), "multicolor-text");
  raw.textKey.warp = { warpStyle: { _value: "warpArc" } };
  assert.equal(outlineEligibility(raw), "warped-text");
});
test("temporary text outline is removed on success and on descriptor/conversion errors", async () => {
  for (const failure of [undefined, "convert", "get"]) {
    const layers = [];
    const copy = { id: 2, textItem: { async convertToShape() { if (failure === "convert") throw new Error("convert"); } },
      async delete() { layers.splice(0); } };
    const layer = { id: 1, async duplicate() { layers.push(copy); return copy; } };
    let calls = 0;
    const reader = { doc: { layers: [{ id: 10, layers }] }, target: () => [], async get() {
      if (failure === "get") throw new Error("get");
      return calls++ ? fixture.shapes.rectangle.path : fixture.shapes.rectangle.raw;
    } };
    const result = await outlineText(layer, textDescriptor(), reader);
    assert.equal(layers.length, 0);
    assert.equal(!!result.shape, !failure);
  }
});
test("a failed duplicate cleanup aborts export instead of rasterizing duplicated text", async () => {
  const copy = { id: 2, textItem: { async convertToShape() { throw new Error("unsupported"); } }, async delete() {} };
  await assert.rejects(outlineText({ id: 1, async duplicate() { return copy; } }, textDescriptor(),
    { doc: { layers: [{ id: 10, layers: [copy] }] } }), /临时层/);
});
