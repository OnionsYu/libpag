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
function paint(value, raw, documentBounds) {
  if (value && value._obj === "solidColorLayer") return { kind: "solid", color: color(value.color) };
  if (!value || value._obj !== "gradientLayer" || !value.gradient ||
      value.gradient.gradientForm._value !== "customStops" || value.type._value !== "linear")
    throw new Error("未支持的绘制类型");
  const source = value.gradient;
  const boolean = (key, fallback) => {
    if (value[key] === undefined) return fallback;
    if (typeof value[key] !== "boolean") throw new Error(`未知渐变标志：${key}`);
    return value[key];
  };
  const aligned = boolean("align", true);
  const bounds = aligned && raw.hasVectorMask ? raw.pathBounds && raw.pathBounds.pathBounds : documentBounds;
  if (!bounds) throw new Error("缺少渐变参考矩形");
  const left = number(bounds.left, "left"), top = number(bounds.top, "top");
  const width = number(bounds.right, "right") - left, height = number(bounds.bottom, "bottom") - top;
  const angle = value.angle ? unit(value.angle, "angleUnit") * Math.PI / 180 : 0;
  const scale = value.scale ? unit(value.scale, "percentUnit") / 100 : 1;
  const x = Math.cos(angle), y = -Math.sin(angle);
  const length = (Math.abs(width * x) + Math.abs(height * y)) * scale;
  if (!(width > 0 && height > 0 && length > 0)) throw new Error("无效渐变范围");
  const offset = value.offset || { horizontal: { _unit: "percentUnit", _value: 0 },
    vertical: { _unit: "percentUnit", _value: 0 } };
  const cx = left + width * (0.5 + unit(offset.horizontal, "percentUnit") / 100);
  const cy = top + height * (0.5 + unit(offset.vertical, "percentUnit") / 100);
  if (!Array.isArray(source.colors) || !source.colors.length || !Array.isArray(source.transparency) ||
      !source.transparency.length) throw new Error("缺少渐变标记");
  return { kind: "gradient", gradient: {
    type: value.type._value, interpolation: value.gradientsInterpolationMethod._value,
    smoothness: number(source.interfaceIconFrameDimmed, "smoothness") / 4096,
    reverse: boolean("reverse", false), dither: boolean("dither", false),
    matrix: [x * length, y * length, -y * length, x * length, cx - x * length / 2, cy - y * length / 2],
    colors: source.colors.map(stop => {
      if (stop.type._value !== "userStop") throw new Error("动态渐变颜色尚未解析");
      return { position: number(stop.location, "location") / 4096,
        midpoint: number(stop.midpoint, "midpoint") / 100, color: color(stop.color) };
    }),
    opacity: source.transparency.map(stop => ({ position: number(stop.location, "location") / 4096,
      midpoint: number(stop.midpoint, "midpoint") / 100, opacity: unit(stop.opacity, "percentUnit") / 100 }))
  } };
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
function shape(raw, path, documentBounds) {
  if (!raw.hasVectorMask || raw.vectorMaskEnabled !== true || raw.vectorMaskDensity !== 255 ||
      raw.vectorMaskFeather !== 0 || raw.vectorMaskEmpty !== false || raw.useAlignedRendering !== false)
    throw new Error("矢量蒙版或像素对齐需要 Photoshop 渲染");
  const value = { components: pathComponents(path) };
  const style = raw.AGMStrokeStyleInfo;
  const fillEnabled = style ? style.fillEnabled : raw.fillEnabled;
  if (typeof fillEnabled !== "boolean") throw new Error("未知形状填充状态");
  if (fillEnabled) {
    if (!Array.isArray(raw.adjustment) || raw.adjustment.length !== 1)
      throw new Error("形状填充需要 Photoshop 渲染");
    value.fill = paint(raw.adjustment[0], raw, documentBounds);
  }
  if (style && style.strokeEnabled) {
    if (style.strokeStyleVersion !== 2 || style.strokeStyleBlendMode._value !== "normal" ||
        style.strokeStyleStrokeAdjust !== false)
      throw new Error("描边样式需要 Photoshop 渲染");
    const caps = { strokeStyleButtCap: "butt", strokeStyleRoundCap: "round", strokeStyleSquareCap: "square" };
    const joins = { strokeStyleMiterJoin: "miter", strokeStyleRoundJoin: "round", strokeStyleBevelJoin: "bevel" };
    const aligns = { strokeStyleAlignCenter: "center", strokeStyleAlignInside: "inner", strokeStyleAlignOutside: "outer" };
    const cap = caps[style.strokeStyleLineCapType._value], join = joins[style.strokeStyleLineJoinType._value];
    const align = aligns[style.strokeStyleLineAlignment._value];
    if (!cap || !join || !align) throw new Error("未知描边枚举");
    const width = unit(style.strokeStyleLineWidth, "pixelsUnit");
    // The action descriptor stores actual points, even though the Photoshop UI presents
    // width multiples. Recorded 72/300 PPI references qualify resolution / 72 here.
    if (!Array.isArray(style.strokeStyleLineDashSet)) throw new Error("缺少虚线样式");
    const pointsToPixels = number(style.strokeStyleResolution, "stroke resolution") / 72;
    if (!(pointsToPixels > 0)) throw new Error("无效描边分辨率");
    const dashes = style.strokeStyleLineDashSet.map(value => unit(value, "pointsUnit") * pointsToPixels);
    if (dashes.some(value => value < 0) || (dashes.length && !dashes.some(value => value > 0)))
      throw new Error("无效虚线长度");
    value.stroke = { paint: paint(style.strokeStyleContent, raw, documentBounds), width,
      opacity: unit(style.strokeStyleOpacity, "percentUnit") / 100,
      miterLimit: number(style.strokeStyleMiterLimit, "miterLimit"), cap, join, align,
      dashes, dashOffset: unit(style.strokeStyleLineDashOffset, "pointsUnit") * pointsToPixels };
  }
  return value;
}

function vectorMask(raw, path, state) {
  if (!raw.hasVectorMask || [raw.vectorMaskEnabled, raw.vectorMaskLinked, raw.vectorMaskEmpty]
    .some(value => typeof value !== "boolean") || raw.useAlignedRendering !== false)
    throw new Error("未确认的矢量蒙版状态");
  const density = number(raw.vectorMaskDensity, "vectorMaskDensity") / 255;
  const feather = number(raw.vectorMaskFeather, "vectorMaskFeather");
  if (density < 0 || density > 1 || feather < 0) throw new Error("无效的矢量蒙版参数");
  if (!state || typeof state.inverted !== "boolean" || typeof state.initialFill !== "boolean" ||
      state.enabled !== raw.vectorMaskEnabled || state.linked !== raw.vectorMaskLinked)
    throw new Error("矢量蒙版标志未确认或与描述符不一致");
  return { ...state, density, feather,
    components: raw.vectorMaskEmpty ? [] : pathComponents(path) };
}

class DescriptorReader {
  constructor(ps, doc, hostVersion, readMasks) {
    this.ps = ps; this.doc = doc; this.hostVersion = hostVersion; this.layers = new Map();
    this.readMasks = readMasks || (() => require("./mask-metadata.js").captureMaskMetadata(doc));
  }
  async maskState(id, raw) {
    if (!this.maskMetadata) this.maskMetadata = this.readMasks().catch(() => null);
    const metadata = await this.maskMetadata;
    const state = metadata && metadata.get(id);
    if (!state || state.enabled !== raw.vectorMaskEnabled || state.linked !== raw.vectorMaskLinked)
      throw new Error("矢量蒙版标志无法可靠读取");
    return state;
  }
  async get(target) {
    const [result] = await this.ps.action.batchPlay([{ _obj: "get", _target: target,
      _options: { dialogOptions: "silent" } }], {});
    if (!result || result._obj === "error") throw new Error("Photoshop 描述符不可用");
    return result;
  }
  target(id) { return [{ _ref: "layer", _id: id }, { _ref: "document", _id: this.doc.id }]; }
  async documentFeatures() {
    // RGB blend gamma is application state, absent from both layer get and document JSON.
    // Only the observed disabled value is qualified. An unavailable setting is not a default.
    try {
      if (!/^27\.4(?:\.|$)/.test(this.hostVersion || "")) throw new Error("unknown host");
      const result = await this.get([{ _property: "colorSettings" },
        { _ref: "application", _enum: "ordinal", _value: "targetEnum" }]);
      if (result.colorSettings && result.colorSettings.RGBBlendGamma === false) return [];
      return ["compositing-environment:rgb-blend-gamma"];
    } catch (_) {
      return ["compositing-environment:unknown"];
    }
  }
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
    if (raw.hasVectorMask && source.kind !== "fill") {
      try {
        source.vectorMask = vectorMask(raw, raw.vectorMaskEmpty ? null : await this.get([
          { _ref: "path", _enum: "path", _value: "vectorMask" }, ...this.target(layer.id)]),
          await this.maskState(layer.id, raw));
      } catch (error) { source.rasterFeatures.push(`vector-mask:${String(error)}`); }
      // Converting masked text to a shape can replace its original mask. Keep that operation
      // in Photoshop until its interaction with text outlines is independently qualified.
      if (source.kind === "text") source.rasterFeatures.push("text-vector-mask");
    }
    if (source.kind === "text") {
      source.text = raw.textKey && typeof raw.textKey.textKey === "string" ? raw.textKey.textKey : "";
      if (!source.features.length && !source.rasterFeatures.length) {
        // Lazy load avoids a module initialization cycle with the shared shape/color decoder.
        const result = await require("./text-outliner.js").outlineText(layer, raw, this);
        if (result.shape) source.shape = result.shape;
        else source.rasterFeatures.push(result.reason);
      }
    }
    if (source.kind === "fill") {
      const bounds = { left: 0, top: 0, right: this.doc.width, bottom: this.doc.height };
      try {
        if (raw.hasVectorMask) {
          source.kind = "shape";
          const state = await this.maskState(layer.id, raw);
          if (state.inverted || state.initialFill) throw new Error("反相或初始填充形状需要 Photoshop 渲染");
          source.shape = shape(raw, await this.get([
            { _ref: "path", _enum: "path", _value: "vectorMask" }, ...this.target(layer.id)]), bounds);
        } else {
          if (!Array.isArray(raw.adjustment) || raw.adjustment.length !== 1) throw new Error("未知填充");
          source.fill = paint(raw.adjustment[0], raw, bounds);
        }
      } catch (error) { source.rasterFeatures.push(`paint:${String(error)}`); }
    }
  }
}
module.exports = { DescriptorReader, pathComponents, shape, vectorMask, paint, color, unit };
