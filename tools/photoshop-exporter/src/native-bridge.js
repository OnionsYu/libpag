/* Copyright (C) 2026 Tencent. Licensed under the Apache License, Version 2.0. */
"use strict";

class NativeBridge {
  constructor(addon) {
    this.addon = addon;
    this.info = JSON.parse(addon.info());
    if (this.info.protocolVersion !== 2 || this.info.schemaVersion !== 3) {
      throw new Error("PAGX 原生模块版本不匹配，请重新安装完整插件");
    }
  }
  begin(source, forceComposite) {
    return JSON.parse(this.addon.begin(JSON.stringify(source), forceComposite));
  }
  addRaster(sessionId, metadata, bytes) {
    if (!(bytes instanceof Uint8Array || bytes instanceof Uint16Array) ||
        metadata.componentSize !== bytes.BYTES_PER_ELEMENT * 8) {
      throw new Error("图像数组类型与像素位深不匹配");
    }
    const buffer = bytes.byteOffset === 0 && bytes.byteLength === bytes.buffer.byteLength
      ? bytes.buffer : bytes.buffer.slice(bytes.byteOffset, bytes.byteOffset + bytes.byteLength);
    this.addon.addRaster(sessionId, JSON.stringify(metadata), buffer);
  }
  finish(sessionId) { return this.addon.finish(sessionId); }
  discard(sessionId) { this.addon.discard(sessionId); }
}

module.exports = { NativeBridge };
