/* Copyright (C) 2026 Tencent. Licensed under the Apache License, Version 2.0. */
"use strict";
const { shape, color } = require("./descriptor-reader.js");

function outlineEligibility(raw) {
  const text = raw.textKey;
  if (raw.textWarningLevel !== 0 || !text || !Array.isArray(text.textStyleRange) || !text.textStyleRange.length)
    return "text-metadata-unavailable";
  if (text.warp && text.warp.warpStyle._value !== "warpNone") return "warped-text";
  let paint;
  for (const range of text.textStyleRange) {
    const style = { ...range.textStyle.baseParentStyle, ...range.textStyle };
    if (style.fontAvailable !== true) return "missing-font";
    if (style.stroke || style.fill === false) return "text-paint";
    let current;
    try { current = JSON.stringify(color(style.color)); }
    catch (_) { return "text-paint"; }
    if (paint !== undefined && paint !== current) return "multicolor-text";
    paint = current;
  }
  return undefined;
}

// Convert a disposable duplicate, never the source layer. Even a later raster fallback must see
// the original text. The export document itself is owned by registerAutoCloseDocument on cancel.
async function outlineText(layer, raw, reader) {
  const reason = outlineEligibility(raw);
  if (reason) return { reason };
  let copy;
  try {
    copy = await layer.duplicate();
    if (!copy || copy.id === layer.id) throw new Error("Text duplication did not produce a new layer");
    await copy.textItem.convertToShape();
    const descriptor = await reader.get(reader.target(copy.id));
    const path = await reader.get([{ _ref: "path", _enum: "path", _value: "vectorMask" }, ...reader.target(copy.id)]);
    return { shape: shape(descriptor, path) };
  } catch (error) {
    return { reason: `text-outline:${String(error)}` };
  } finally {
    if (copy && copy.id !== layer.id) {
      // Do not turn cleanup failure into a successful fallback with duplicated content.
      await copy.delete();
      const contains = layers => Array.from(layers).some(item => item.id === copy.id ||
        (item.layers && contains(item.layers)));
      if (contains(reader.doc.layers)) throw new Error("文字轮廓临时层未被删除");
    }
  }
}
module.exports = { outlineEligibility, outlineText };
