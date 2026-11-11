/* Copyright (C) 2026 Tencent. Licensed under the Apache License, Version 2.0. */
"use strict";
const { checkCancelled } = require("./export-service.js");
const { DescriptorReader } = require("./descriptor-reader.js");
const SRGB = "sRGB IEC61966-2.1";

function list(collection) { return Array.from(collection); }
function rect(bounds, emptyGroup) {
  const result = { left: Number(bounds.left), top: Number(bounds.top),
    width: Number(bounds.right) - Number(bounds.left),
    height: Number(bounds.bottom) - Number(bounds.top) };
  if (Object.values(result).every(Number.isFinite)) return result;
  // Photoshop returns an empty object for the bounds of empty groups (including nested ones).
  if (emptyGroup) return { left: 0, top: 0, width: 0, height: 0 };
  throw new Error("Photoshop 返回了无效的图层边界");
}
function layerKind(layer, constants) {
  const k = constants.LayerKind;
  if (layer.kind === k.GROUP) return "group";
  if (layer.kind === k.NORMAL) return "pixel";
  if (layer.kind === k.TEXT) return "text";
  if (layer.kind === k.SMARTOBJECT) return "smartObject";
  if ([k.SOLIDFILL, k.GRADIENTFILL, k.PATTERNFILL].includes(layer.kind)) return "fill";
  return "unknown";
}

class PhotoshopHost {
  constructor(photoshop, maxPixels, hostVersion) {
    this.ps = photoshop;
    this.maxPixels = maxPixels;
    this.hostVersion = hostVersion;
  }

  async withDocument(input, signal, task) {
    const { app, core } = this.ps;
    const originalId = app.documents.length ? app.activeDocument.id : undefined;
    let ownedId;
    try {
      return await core.executeAsModal(async context => {
        checkCancelled(signal, context);
        let working;
        let sourceName;
        if (input.kind === "current") {
          const source = list(app.documents).find(doc => doc.id === input.documentId);
          if (!source) throw new Error("目标文档已关闭，请重新导出");
          sourceName = source.name;
          working = await source.duplicate("PAGX export (temporary)", false);
        } else if (input.kind === "file") {
          const existingIds = new Set(list(app.documents).map(doc => doc.id));
          working = await app.open(input.file);
          if (!working) throw new Error("Photoshop 未打开所选 PSD");
          if (existingIds.has(working.id)) {
            throw new Error("所选 PSD 已在 Photoshop 中打开。请使用“导出当前文档”，或先关闭它再选择文件。");
          }
          sourceName = input.file.name;
        } else {
          throw new Error("无效的导出来源");
        }
        ownedId = working.id;
        // Photoshop owns cleanup on errors AND Esc cancellation, when further host writes are forbidden.
        await context.hostControl.registerAutoCloseDocument(ownedId);
        checkCancelled(signal, context);
        return await task({ doc: working, sourceName }, context);
      }, { commandName: "导出 PAGX", timeOut: 1 });
    } finally {
      // Auto-close runs at modal exit. Never close a document that we did not create.
      if (ownedId && list(app.documents).some(doc => doc.id === ownedId)) {
        throw new Error("Photoshop 未能关闭临时文档；请手动关闭 PAGX 临时文档（无需保存）");
      }
      const original = list(app.documents).find(doc => doc.id === originalId);
      if (original && app.activeDocument.id !== originalId) {
        await core.executeAsModal(async () => { app.activeDocument = original; },
          { commandName: "恢复原文档", timeOut: 1 });
      }
    }
  }

  async snapshot({ doc, sourceName }) {
    const { constants } = this.ps;
    const depths = constants.BitsPerChannelType;
    const bitDepth = doc.bitsPerChannel === depths.EIGHT ? 8 :
      doc.bitsPerChannel === depths.SIXTEEN ? 16 : 0;
    if (!bitDepth) throw new Error("初版支持 8/16 位文档。请先在副本中将 HDR/位图转换为 8 或 16 位。");
    if (doc.pixelAspectRatio !== 1) throw new Error("初版仅支持方形像素，请在副本中转换像素长宽比后导出。");
    if (doc.artboards.length) throw new Error("初版尚不支持画板文档，请在副本中将画板转换为普通画布后导出。");
    const width = Number(doc.width), height = Number(doc.height);
    if (!Number.isSafeInteger(width) || !Number.isSafeInteger(height) ||
        width <= 0 || height <= 0 || width * height * (bitDepth / 8) > this.maxPixels) {
      throw new Error("画布尺寸无效或超过内存限制（8 位 64M / 16 位 32M 像素）");
    }
    let count = 0;
    const reader = new DescriptorReader(this.ps, doc, this.hostVersion);
    await reader.init();
    const walk = async (layers, depth) => {
      if (depth > 128) throw new Error("图层嵌套过深");
      let base = "";
      // Photoshop DOM is top-to-bottom. The source schema and PAGX use paint order.
      const result = [];
      for (const layer of list(layers).reverse()) {
        if (++count > 100000) throw new Error("图层数量超过限制");
        const kind = layerKind(layer, constants);
        const id = `ps:${layer.id}`;
        const clippingBase = layer.isClippingMask ? base : "";
        if (!layer.isClippingMask) base = id;
        const children = kind === "group" ? await walk(layer.layers, depth + 1) : [];
        const emptyGroup = kind === "group" && children.every(child =>
          child.bounds.width === 0 || child.bounds.height === 0);
        const value = { id, sourceLayerId: layer.id, name: layer.name, kind,
          visible: layer.visible, opacity: layer.opacity / 100,
          fillOpacity: layer.fillOpacity / 100,
          blendMode: layer.blendMode === constants.BlendMode.NORMAL ? "normal" :
            layer.blendMode === constants.BlendMode.PASSTHROUGH ? "passThrough" : String(layer.blendMode),
          compositingKnown: false, clippingBase, features: [], rasterFeatures: [], bounds: rect(layer.bounds, emptyGroup),
          text: "", children };
        await reader.read(layer, value);
        result.push(value);
      }
      return result;
    };
    return { schemaVersion: 2, name: sourceName, hostVersion: this.hostVersion,
      resolution: doc.resolution, colorMode: doc.mode === constants.DocumentMode.RGB ? "RGB" : String(doc.mode),
      colorProfile: doc.colorProfileName, bitDepth,
      bounds: { left: 0, top: 0, width, height },
      features: [], layers: await walk(doc.layers, 0) };
  }

  async prepareCaptures(document, source, plan) {
    if (plan.composite || !plan.rasters.length) return;
    const layers = new Map();
    const walk = (items, ancestors) => {
      for (const layer of list(items)) {
        layers.set(`ps:${layer.id}`, { layer, ancestors, visible: layer.visible, opacity: layer.opacity });
        if (layer.kind === this.ps.constants.LayerKind.GROUP) walk(layer.layers, [...ancestors, layer]);
      }
    };
    walk(document.doc.layers, []);
    for (const request of plan.rasters) if (!layers.has(request.layerId)) throw new Error("采集计划引用了不存在的图层");
    document.captureLayers = layers;
  }

  async capture(document, source, request, consume) {
    const { doc } = document;
    if (request.layerId) {
      const state = document.captureLayers;
      const target = state && state.get(request.layerId);
      if (!target) throw new Error("无效的独立采集请求");
      for (const { layer } of state.values()) layer.visible = false;
      target.layer.visible = true;
      for (const ancestor of target.ancestors) { ancestor.visible = true; ancestor.opacity = 100; }
      target.layer.opacity = target.opacity;
    }
    const bounds = source.bounds;
    const result = await this.ps.imaging.getPixels({ documentID: doc.id,
      sourceBounds: { left: bounds.left, top: bounds.top,
        right: bounds.left + bounds.width, bottom: bounds.top + bounds.height },
      colorSpace: "RGB", colorProfile: SRGB,
      // PS 27.4 fails when requesting 8-bit pixels from a 16-bit composite. Read the source
      // precision, perform the color conversion here, and quantize once in the native bridge.
      componentSize: source.bitDepth, applyAlpha: false });
    const data = result.imageData;
    if (!data) throw new Error("Photoshop 没有返回图像数据");
    try {
      if (result.level !== 0 || data.colorSpace !== "RGB" || data.colorProfile !== SRGB ||
          data.componentSize !== source.bitDepth || ![8, 16].includes(data.componentSize) ||
          ![3, 4].includes(data.components)) {
        throw new Error("Photoshop 返回了不支持的颜色或缩放格式");
      }
      const actual = result.sourceBounds;
      const width = data.width, height = data.height;
      if (width * height * (data.componentSize / 8) > this.maxPixels) throw new Error("图像超过内存限制");
      if (width && (actual.right - actual.left !== width || actual.bottom - actual.top !== height)) {
        throw new Error("Photoshop 图像尺寸与实际采集区域不一致");
      }
      const pixels = await data.getData({ chunky: true, fullRange: true });
      consume({ requestId: request.id, left: actual.left, top: actual.top,
        width, height, components: data.components, componentSize: data.componentSize,
        colorProfile: data.colorProfile, alpha: "straight", fullRange: true }, pixels);
    } finally { data.dispose(); }
  }
}

module.exports = { PhotoshopHost, layerKind };
