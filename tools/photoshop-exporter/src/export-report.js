/* Copyright (C) 2026 Tencent. Licensed under the Apache License, Version 2.0. */
"use strict";

function makeExportReport(source, plan, forceComposite, exporter) {
  const layers = new Map();
  const walk = items => { for (const layer of items) { layers.set(layer.id, layer); walk(layer.children); } };
  walk(source.layers);
  const identify = id => ({ id, name: layers.get(id)?.name || id });
  return {
    reportVersion: 1,
    exporter: { version: exporter.version, protocolVersion: exporter.protocolVersion,
      schemaVersion: source.schemaVersion },
    source: { name: source.name, hostVersion: source.hostVersion, bounds: source.bounds,
      colorMode: source.colorMode, colorProfile: source.colorProfile,
      bitDepth: source.bitDepth, resolution: source.resolution },
    requestedMode: forceComposite ? "appearance" : "vector-first",
    outputMode: plan.composite ? "document-raster" : "mixed",
    statistics: plan.statistics,
    captures: plan.rasters.map(request => ({ id: request.id,
      scope: plan.composite ? "document" : "siblings-and-descendants",
      parent: request.parentId ? identify(request.parentId) : null,
      layers: request.layerIds.map(identify) })),
    diagnostics: plan.diagnostics.map(item => ({ ...item,
      layerName: item.layerId ? identify(item.layerId).name : source.name }))
  };
}

const messages = {
  "document-dependency": "合成信息未知或包含挖空，使用整幅外观快照",
  "document-composite": "按文档颜色、位深或导出模式生成 8 位 sRGB 外观快照",
  "backdrop-range": "与隔离边界内的背景一起转为位图，替换原范围",
  "clipping-range": "剪贴图层与基底一起转为位图，替换原范围",
  "rasterized-content": "当前外观由 Photoshop 渲染为局部位图",
  "outlined-text": "文字转为矢量轮廓，文字编辑及字体微调不保留"
};
function summarizeReport(report) {
  const s = report.statistics;
  const lines = [report.outputMode === "document-raster" ? "输出：整幅位图" : "输出：矢量优先",
    `矢量形状 ${s.vectorShapes} · 文字轮廓 ${s.outlinedTexts}`,
    `源位图层 ${s.sourceImages} · 回退范围 ${s.fallbackScopes}`];
  const diagnostics = report.diagnostics.slice(0, 12);
  for (const item of diagnostics) lines.push(`${item.layerName}：${messages[item.code] || item.message}`);
  if (report.diagnostics.length > diagnostics.length) lines.push("其余明细见同名导出报告。");
  return lines.join("\n");
}

module.exports = { makeExportReport, summarizeReport };
