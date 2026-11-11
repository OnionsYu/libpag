/* Copyright (C) 2026 Tencent. Licensed under the Apache License, Version 2.0. */
"use strict";

function number(value, name) {
  if (typeof value !== "number" || !Number.isFinite(value)) throw new Error(`无效描述符数值：${name}`);
  return value;
}
function unit(value, expected) {
  if (!value || value._unit !== expected) throw new Error(`未支持的描述符单位：${expected}`);
  return number(value._value, expected);
}
function color(value) {
  if (!value || value._obj !== "RGBColor") throw new Error("形状不是 RGB 颜色");
  const channel = key => {
    const v = number(value[key], key);
    if (v < -0.001 || v > 255.001) throw new Error("形状颜色超出范围");
    // Photoshop's 8-bit descriptor can contain a few ulps above 255.
    return Math.min(1, Math.max(0, v / 255));
  };
  return { red: channel("red"), green: channel("grain"), blue: channel("blue"), alpha: 1 };
}
function point(value) {
  return { x: unit(value.horizontal, "pixelsUnit"), y: unit(value.vertical, "pixelsUnit") };
}
function pathComponents(value) {
  const components = value && value.pathContents && value.pathContents.pathComponents;
  if (!Array.isArray(components) || !components.length) throw new Error("缺少矢量路径");
  return components.map(component => {
    const operation = component.shapeOperation && component.shapeOperation._value;
    if (!["add", "subtract", "intersect", "xor"].includes(operation)) throw new Error("未支持的路径运算");
    if (component.windingFill !== undefined && typeof component.windingFill !== "boolean")
      throw new Error("未支持的路径填充规则");
    if (!Array.isArray(component.subpathListKey)) throw new Error("缺少子路径");
    return { operation: operation === "xor" ? "exclude" : operation,
      winding: component.windingFill === true,
      subpaths: component.subpathListKey.map(path => {
        if (typeof path.closedSubpath !== "boolean" || !Array.isArray(path.points) || !path.points.length)
          throw new Error("无效的子路径");
        return { closed: path.closedSubpath, points: path.points.map(p => ({
          anchor: point(p.anchor), incoming: point(p.backward || p.anchor),
          outgoing: point(p.forward || p.anchor) })) };
      }) };
  });
}
function shape(raw, path) {
  if (!raw.hasVectorMask || raw.vectorMaskEnabled !== true || raw.vectorMaskDensity !== 255 ||
      raw.vectorMaskFeather !== 0 || raw.vectorMaskEmpty !== false || raw.useAlignedRendering !== false)
    throw new Error("矢量蒙版或像素对齐需要 Photoshop 渲染");
  const value = { components: pathComponents(path) };
  const style = raw.AGMStrokeStyleInfo;
  const fillEnabled = style ? style.fillEnabled : raw.fillEnabled;
  if (typeof fillEnabled !== "boolean") throw new Error("未知形状填充状态");
  if (fillEnabled) {
    if (!Array.isArray(raw.adjustment) || raw.adjustment.length !== 1 || raw.adjustment[0]._obj !== "solidColorLayer")
      throw new Error("形状填充需要 Photoshop 渲染");
    value.fill = color(raw.adjustment[0].color);
  }
  if (style && style.strokeEnabled) {
    if (style.strokeStyleVersion !== 2 || style.strokeStyleBlendMode._value !== "normal" ||
        style.strokeStyleStrokeAdjust !== false || style.strokeStyleContent._obj !== "solidColorLayer")
      throw new Error("描边样式需要 Photoshop 渲染");
    const caps = { strokeStyleButtCap: "butt", strokeStyleRoundCap: "round", strokeStyleSquareCap: "square" };
    const joins = { strokeStyleMiterJoin: "miter", strokeStyleRoundJoin: "round", strokeStyleBevelJoin: "bevel" };
    const aligns = { strokeStyleAlignCenter: "center", strokeStyleAlignInside: "inner", strokeStyleAlignOutside: "outer" };
    const cap = caps[style.strokeStyleLineCapType._value], join = joins[style.strokeStyleLineJoinType._value];
    const align = aligns[style.strokeStyleLineAlignment._value];
    if (!cap || !join || !align) throw new Error("未知描边枚举");
    const width = unit(style.strokeStyleLineWidth, "pixelsUnit");
    // Photoshop's dash values are multiples of stroke width, despite their pointsUnit label.
    // Dashed strokes remain a host render until a host fixture qualifies that convention.
    if (style.strokeStyleLineDashSet.length) throw new Error("虚线描边尚未验证");
    if (unit(style.strokeStyleOpacity, "percentUnit") !== 100)
      throw new Error("非默认描边透明度需要 Photoshop 渲染");
    value.stroke = { color: color(style.strokeStyleContent.color), width,
      opacity: unit(style.strokeStyleOpacity, "percentUnit") / 100,
      miterLimit: number(style.strokeStyleMiterLimit, "miterLimit"), cap, join, align,
      dashes: [], dashOffset: 0 };
  }
  return value;
}

class DescriptorReader {
  constructor(ps, doc, hostVersion) { this.ps = ps; this.doc = doc; this.hostVersion = hostVersion; this.layers = new Map(); }
  async get(target) {
    const [result] = await this.ps.action.batchPlay([{ _obj: "get", _target: target,
      _options: { dialogOptions: "silent" } }], {});
    if (!result || result._obj === "error") throw new Error("Photoshop 描述符不可用");
    return result;
  }
  target(id) { return [{ _ref: "layer", _id: id }, { _ref: "document", _id: this.doc.id }]; }
  async init() {
    // The action JSON property includes advanced blend settings that ordinary layer 'get' omits.
    // Missing fields mean defaults only within this qualified host / JSON contract.
    if (!/^27\.4(?:\.|$)/.test(this.hostVersion || "") || !this.ps.action) return;
    try {
      const result = await this.get([{ _property: "json" }, { _ref: "document", _id: this.doc.id }]);
      const document = JSON.parse(result.json);
      if (document.version !== "1.6.1" || document.id !== this.doc.id || !Array.isArray(document.layers)) return;
      const walk = layers => { for (const layer of layers) { this.layers.set(layer.id, layer); if (layer.layers) walk(layer.layers); } };
      walk(document.layers);
      this.qualified = true;
    } catch (_) { /* Unknown capability is represented explicitly by compositingKnown=false. */ }
  }
  async read(layer, source) {
    const json = this.layers.get(layer.id);
    if (!this.qualified || !json) return;
    let raw;
    try { raw = await this.get(this.target(layer.id)); }
    catch (_) { return; }
    if (raw.layerID !== layer.id || [raw.hasUserMask, raw.hasVectorMask, raw.hasFilterMask]
      .some(value => typeof value !== "boolean")) return;
    source.compositingKnown = true;
    const blend = json.blendOptions || {};
    for (const key of Object.keys(blend)) {
      if (!["mode", "opacity", "fillOpacity"].includes(key)) source.features.push(`blend:${key}`);
    }
    const channels = (raw.channelRestrictions || []).map(c => c._value).sort().join(",");
    if (channels !== "blue,grain,red") source.features.push("channel-restrictions");
    if (json.layerEffects || raw.layerEffects) source.features.push("layer-effects");
    if (raw.artboardEnabled || json.artboard) source.features.push("artboard");
    if (raw.hasUserMask) source.rasterFeatures.push("pixel-mask");
    if (raw.hasFilterMask) source.rasterFeatures.push("filter-mask");
    if (raw.hasVectorMask && source.kind !== "fill") source.rasterFeatures.push("vector-mask");
    if (source.kind === "fill" && raw.hasVectorMask) {
      source.kind = "shape";
      try {
        source.shape = shape(raw, await this.get([
          { _ref: "path", _enum: "path", _value: "vectorMask" }, ...this.target(layer.id)]));
      } catch (error) { source.rasterFeatures.push(`shape:${String(error)}`); }
    }
  }
}
module.exports = { DescriptorReader, pathComponents, shape, color, unit };
