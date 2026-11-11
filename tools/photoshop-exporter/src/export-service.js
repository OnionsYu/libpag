/* Copyright (C) 2026 Tencent. Licensed under the Apache License, Version 2.0. */
"use strict";

class ExportCancelled extends Error {
  constructor() { super("已取消导出"); this.name = "ExportCancelled"; }
}

function checkCancelled(signal, context) {
  if (signal.cancelled || (context && context.isCancelled)) throw new ExportCancelled();
}

// Orchestrates a transaction. Host, native bridge and storage have separate lifetimes.
class ExportService {
  constructor({ host, bridge, files }) {
    this.host = host;
    this.bridge = bridge;
    this.files = files;
    this.busy = false;
  }

  async run({ input, folder, forceComposite = false, signal = {}, progress = () => {} }) {
    if (this.busy) throw new Error("已有导出正在进行");
    this.busy = true;
    let sessionId;
    try {
      checkCancelled(signal);
      let name;
      let plan;
      await this.host.withDocument(input, signal, async (document, context) => {
        checkCancelled(signal, context);
        progress("读取 Photoshop 文档…");
        const source = await this.host.snapshot(document);
        name = source.name;
        checkCancelled(signal, context);
        plan = this.bridge.begin(source, forceComposite);
        sessionId = plan.sessionId;
        await this.host.prepareCaptures(document, source, plan);
        for (let index = 0; index < plan.rasters.length; index++) {
          checkCancelled(signal, context);
          progress(`采集图像 ${index + 1}/${plan.rasters.length}…`);
          context.reportProgress({ value: index / plan.rasters.length });
          // capture owns and disposes the PhotoshopImageData, even when the consumer throws.
          await this.host.capture(document, source, plan.rasters[index], (metadata, bytes) => {
            checkCancelled(signal, context);
            this.bridge.addRaster(sessionId, metadata, bytes);
          });
        }
        checkCancelled(signal, context);
      });
      checkCancelled(signal);
      progress("生成并校验 PAGX…");
      const xml = this.bridge.finish(sessionId);
      sessionId = undefined;
      checkCancelled(signal);
      progress("写入 PAGX…");
      const file = await this.files.publish(folder, name, xml, signal);
      return { file, composite: plan.composite, diagnostics: plan.diagnostics };
    } finally {
      try {
        if (sessionId !== undefined) this.bridge.discard(sessionId);
      } finally { this.busy = false; }
    }
  }
}

module.exports = { ExportService, ExportCancelled, checkCancelled };
