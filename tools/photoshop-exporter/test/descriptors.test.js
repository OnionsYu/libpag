/* Copyright (C) 2026 Tencent. Licensed under the Apache License, Version 2.0. */
"use strict";
const test = require("node:test");
const assert = require("node:assert/strict");
const { DescriptorReader, shape, pathComponents } = require("../src/descriptor-reader.js");
const fixture = require("./fixtures/ps27-descriptors.json");
const appearance = require("./fixtures/ps27-appearance.json");

test("recorded dash units and independent stroke opacity become pixel-space v3 paints", () => {
  const dash = shape(appearance.dash.value, appearance.dash.path);
  assert.deepEqual(dash.stroke.dashes, [3, 2]);
  assert.equal(dash.stroke.dashOffset, 1);
  const highResolution = structuredClone(appearance.dash);
  highResolution.value.AGMStrokeStyleInfo.strokeStyleResolution = 300;
  assert.deepEqual(shape(highResolution.value, highResolution.path).stroke.dashes, [12.5, 25 / 3]);
  assert.equal(dash.stroke.paint.kind, "solid");
  const translucent = shape(appearance["stroke-opacity"].value, appearance["stroke-opacity"].path);
  assert.equal(translucent.stroke.opacity, 89 / 255); // Photoshop quantizes 35% to an 8-bit value.
});

test("gradient source retains interpolation, separate stop streams and document coordinates", () => {
  for (const method of ["classic", "linear", "perceptual"]) {
    const raw = appearance[`gradient-${method}`];
    const value = shape(raw.value, raw.path).fill;
    assert.equal(value.kind, "gradient");
    assert.equal(value.gradient.interpolation, method);
    assert.deepEqual(value.gradient.matrix.map(v => v || 0), [200, 0, 0, 200, 20, 90]);
    assert.deepEqual(value.gradient.opacity.map(stop => stop.opacity), [1, 1]);
    assert.equal(value.gradient.colors[0].midpoint, 0.5);
    assert.equal(value.gradient.smoothness, 1);
  }
});

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
  assert.equal(value.fill.color.blue, ellipse.raw.adjustment[0].color.blue / 255);
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
