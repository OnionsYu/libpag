/* Copyright (C) 2026 Tencent. Licensed under the Apache License, Version 2.0. */
"use strict";

// Layer/path get and action JSON both omit the PSD vector-mask inversion flag. Read only
// layer metadata from a host-saved copy. Pixel channels and the merged image are never read.
async function readMaskMetadata(read, size) {
  let bytesRead = 0;
  const bytes = async (offset, length) => {
    if (!Number.isSafeInteger(offset) || !Number.isSafeInteger(length) || offset < 0 || length < 0 ||
        offset + length > size || (bytesRead += length) > 32 * 1024 * 1024)
      throw new Error("PSD 蒙版元数据超出边界或大小限制");
    const value = await read(offset, length);
    if (!(value instanceof Uint8Array) || value.length !== length) throw new Error("PSD 元数据读取不完整");
    return value;
  };
  const u32 = async at => {
    const b = await bytes(at, 4);
    return new DataView(b.buffer, b.byteOffset, 4).getUint32(0, false);
  };
  const tag = (b, at = 0) => String.fromCharCode(...b.subarray(at, at + 4));
  const header = await bytes(0, 26);
  const view = new DataView(header.buffer, header.byteOffset, header.byteLength);
  if (tag(header) !== "8BPS" || view.getUint16(4) !== 1 || view.getUint16(22) !== 8)
    throw new Error("未验证的 PSD 蒙版元数据格式");
  let at = 26;
  for (let i = 0; i < 2; ++i) at += 4 + await u32(at); // Color mode data and image resources.
  const sectionEnd = at + 4 + await u32(at);
  at += 4;
  const layerEnd = at + 4 + await u32(at);
  at += 4;
  if (layerEnd > sectionEnd || sectionEnd > size) throw new Error("无效 PSD 图层区边界");
  if (at === layerEnd) return new Map();
  const countBytes = await bytes(at, 2);
  const count = Math.abs(new DataView(countBytes.buffer, countBytes.byteOffset, 2).getInt16(0));
  at += 2;
  const result = new Map();
  for (let i = 0; i < count; ++i) {
    const record = await bytes(at, 18);
    const channels = new DataView(record.buffer, record.byteOffset, 18).getUint16(16);
    if (channels > 56) throw new Error("无效 PSD 通道数量");
    at += 18 + channels * 6 + 12;
    const end = at + 4 + await u32(at);
    at += 4;
    if (end > layerEnd) throw new Error("无效 PSD 图层附加区");
    for (let j = 0; j < 2; ++j) at += 4 + await u32(at); // Pixel mask and blend ranges.
    const nameLength = (await bytes(at, 1))[0];
    at += Math.ceil((nameLength + 1) / 4) * 4;
    let id, mask;
    while (at < end) {
      if (at + 12 > end) throw new Error("截断的 PSD 附加信息");
      const block = await bytes(at, 12);
      const signature = tag(block), key = tag(block, 4);
      const length = new DataView(block.buffer, block.byteOffset, 12).getUint32(8);
      const data = at + 12, next = data + Math.ceil(length / 2) * 2;
      if (!["8BIM", "8B64"].includes(signature) || next > end) throw new Error("无效 PSD 信息块");
      if (key === "lyid") {
        if (id !== undefined || length !== 4) throw new Error("无效 PSD 图层标识");
        id = await u32(data);
      } else if (key === "vmsk" || key === "vsms") {
        const padding = (length - 8) % 26;
        if (mask || length < 8 || padding > 3 || await u32(data) !== 3)
          throw new Error("未验证的矢量蒙版结构");
        if (padding && (await bytes(data + length - padding, padding)).some(value => value !== 0))
          throw new Error("未知矢量蒙版尾部数据");
        const flags = await u32(data + 4);
        if (flags & ~7) throw new Error("未知矢量蒙版标志");
        mask = { inverted: !!(flags & 1), linked: !(flags & 2), enabled: !(flags & 4) };
        for (let offset = 8; offset < length - padding;) {
          const amount = Math.min(26 * 2048, length - padding - offset);
          const records = await bytes(data + offset, amount);
          const values = new DataView(records.buffer, records.byteOffset, records.byteLength);
          for (let r = 0; r < amount; r += 26) {
            if (values.getUint16(r) === 8) {
              if (mask.initialFill !== undefined || values.getUint16(r + 2) > 1)
                throw new Error("未知矢量蒙版初始填充");
              mask.initialFill = values.getUint16(r + 2) === 1;
            }
          }
          offset += amount;
        }
        if (mask.initialFill === undefined) throw new Error("缺少矢量蒙版初始填充");
      }
      at = next;
    }
    if (at !== end) throw new Error("PSD 图层附加区长度不一致");
    if (mask) {
      if (id === undefined || result.has(id)) throw new Error("蒙版缺少唯一图层标识");
      result.set(id, mask);
    }
  }
  return result;
}

let sequence = 0;
async function captureMaskMetadata(doc) {
  const { storage } = require("uxp");
  const fs = require("fs");
  const folder = await storage.localFileSystem.getTemporaryFolder();
  const file = await folder.createFile(`pagx-mask-${Date.now()}-${++sequence}.psd`, { overwrite: false });
  let fd;
  try {
    await doc.saveAs.psd(file, { embedColorProfile: true }, true);
    const metadata = await file.getMetadata();
    fd = await fs.open(`plugin-temp:/${file.name}`, "r");
    return await readMaskMetadata(async (offset, length) => {
      const buffer = new ArrayBuffer(length);
      const result = await fs.read(fd, buffer, 0, length, offset);
      if (result.bytesRead !== length) throw new Error("PSD 元数据读取不完整");
      return new Uint8Array(buffer);
    }, metadata.size);
  } finally {
    try { if (fd !== undefined) await fs.close(fd); }
    finally { await file.delete(); }
  }
}
module.exports = { readMaskMetadata, captureMaskMetadata };
