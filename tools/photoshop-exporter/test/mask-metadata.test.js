/* Copyright (C) 2026 Tencent. Licensed under the Apache License, Version 2.0. */
"use strict";
const test = require("node:test");
const assert = require("node:assert/strict");
const { readMaskMetadata } = require("../src/mask-metadata.js");
const u32 = value => { const b = Buffer.alloc(4); b.writeUInt32BE(value); return b; };
const u16 = value => { const b = Buffer.alloc(2); b.writeUInt16BE(value); return b; };
const block = (key, data) => Buffer.concat([Buffer.from(`8BIM${key}`), u32(data.length), data,
  Buffer.alloc(data.length % 2)]);
function fixture(flags = 0, initialFill = 0, duplicate = false) {
  const header = Buffer.alloc(26); header.write("8BPS"); header.writeUInt16BE(1, 4);
  header.writeUInt16BE(3, 12); header.writeUInt32BE(16, 14); header.writeUInt32BE(16, 18);
  header.writeUInt16BE(8, 22); header.writeUInt16BE(3, 24);
  // Seven 26-byte records, including initial-fill; Photoshop pads this block to 4 bytes.
  const path = Buffer.alloc(26 * 7 + 2); path.writeUInt16BE(8); path.writeUInt16BE(initialFill, 2);
  const extra = Buffer.concat([u32(0), u32(0), Buffer.from([1, 65, 0, 0]),
    block("zzzz", Buffer.alloc(100000, 0x7f)),
    block("vmsk", Buffer.concat([u32(3), u32(flags), path])), block("lyid", u32(42))]);
  const layer = Buffer.concat([Buffer.alloc(18 + 12), u32(extra.length), extra]);
  const layers = Buffer.concat([u16(duplicate ? 2 : 1), layer, ...(duplicate ? [layer] : [])]);
  const section = Buffer.concat([u32(layers.length), layers, u32(0)]);
  return Buffer.concat([header, u32(0), u32(0), u32(section.length), section, Buffer.alloc(100000, 0xff)]);
}
function read(buffer, observe = () => {}) {
  return readMaskMetadata(async (offset, size) => {
    observe(offset, size);
    return buffer.subarray(offset, offset + size);
  }, buffer.length);
}
test("PSD mask flags and initial fill are independent; reads skip image and unrelated block payloads", async () => {
  for (let flags = 0; flags < 8; ++flags) {
    const data = fixture(flags, flags % 2);
    let total = 0;
    const masks = await read(data, (offset, size) => { total += size; assert(offset + size <= data.length - 100000); });
    assert(total < 1024);
    assert.deepEqual(masks.get(42), { inverted: !!(flags & 1), linked: !(flags & 2),
      enabled: !(flags & 4), initialFill: !!(flags % 2) });
  }
});
test("unknown flags, initial fill, truncated records, duplicate ids and unsupported PSD variants fail closed", async () => {
  await assert.rejects(read(fixture(8)), /未知矢量蒙版标志/);
  await assert.rejects(read(fixture(0, 2)), /初始填充/);
  await assert.rejects(read(fixture(0, 0, true)), /唯一图层标识/);
  const data = fixture();
  await assert.rejects(read(data.subarray(0, 50)), /边界/);
  for (const [offset, value] of [[4, 2], [22, 16]]) {
    const changed = Buffer.from(data); changed.writeUInt16BE(value, offset);
    await assert.rejects(read(changed), /格式/);
  }
});
