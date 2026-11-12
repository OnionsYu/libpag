/* Copyright (C) 2026 Tencent. Licensed under the Apache License, Version 2.0. */
"use strict";
const test = require("node:test");
const assert = require("node:assert/strict");
const { DescriptorReader, shape, pathComponents } = require("../src/descriptor-reader.js");
const fixture = require("./fixtures/ps27-descriptors.json");

test("RGB blending environment must be explicitly qualified, including normal layers", async () => {
  for (const gamma of [false, true, 1, 2.2, undefined, "false"]) {
    const reader = new DescriptorReader({ action: { async batchPlay([request]) {
      assert.equal(request._target[0]._property, "colorSettings");
      assert.equal(request._target[1]._ref, "application");
      return [{ colorSettings: { RGBBlendGamma: gamma, TextBlendGamma: 1.45 } }];
    } } }, { id: 1 }, "27.4.0");
    assert.deepEqual(await reader.documentFeatures(), gamma === false ? [] :
      ["compositing-environment:rgb-blend-gamma"]);
  }
  const reader = new DescriptorReader({}, { id: 1 }, "27.4.0");
  assert.deepEqual(await reader.documentFeatures(), ["compositing-environment:unknown"]);
});

test("recorded Photoshop paths retain fractional geometry and absolute closing handles", () => {
  const ellipse = fixture.shapes.ellipse;
  const value = shape(ellipse.raw, ellipse.path);
  const points = value.components[0].subpaths[0].points;
  assert.deepEqual(points[0].anchor, { x: 175.5, y: 30.25 });
  assert.deepEqual(points[0].incoming, { x: 145.124347, y: 30.25 });
  assert.deepEqual(points[3].outgoing, { x: 120.5, y: 54.874347 });
  assert.equal(value.fill.blue, ellipse.raw.adjustment[0].color.blue / 255);
  assert.equal(value.components[0].winding, false);
});

test("unqualified units, paint, mask and path operations never masquerade as supported geometry", () => {
  for (const mutate of [
    f => { f.raw.vectorMaskFeather = 2; },
    f => { f.raw.adjustment[0]._obj = "gradientLayer"; },
    f => { f.path.pathContents.pathComponents[0].shapeOperation._value = "future"; },
    f => { f.path.pathContents.pathComponents[0].subpathListKey[0].points[0].anchor.horizontal._unit = "percentUnit"; },
    f => { f.path.pathContents.pathComponents[0].subpathListKey[0].points[0].anchor.horizontal._value = true; }
  ]) {
    const value = structuredClone(fixture.shapes.rectangle); mutate(value);
    assert.throws(() => shape(value.raw, value.path));
  }
});

test("advanced blending omitted by ordinary get is detected through recorded action JSON", async () => {
  for (const [feature, json] of Object.entries(fixture.advanced)) {
    const raw = { ...fixture.shapes.rectangle.raw, layerID: 3 };
    const ps = { action: { async batchPlay([request]) {
      return request._target[0]._property === "json" ? [{ json: JSON.stringify(json) }] : [raw];
    } } };
    const reader = new DescriptorReader(ps, { id: json.id }, fixture.hostVersion);
    await reader.init();
    const source = { kind: "pixel", compositingKnown: false, features: [], rasterFeatures: [] };
    await reader.read({ id: 3 }, source);
    assert.equal(source.compositingKnown, true);
    assert(source.features.includes(`blend:${feature}`));
  }
});

test("unavailable or future descriptor contracts keep composition unknown", async () => {
  for (const hostVersion of ["28.0.0", "27.4.0"]) {
    const reader = new DescriptorReader({ action: { async batchPlay() { throw new Error("unavailable"); } } }, { id: 1 }, hostVersion);
    await reader.init();
    const source = { compositingKnown: false };
    await reader.read({ id: 3 }, source);
    assert.equal(source.compositingKnown, false);
  }
});
