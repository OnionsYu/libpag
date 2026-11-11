/* Copyright (C) 2026 Tencent. Licensed under the Apache License, Version 2.0. */
"use strict";
const { checkCancelled } = require("./export-service.js");

function baseName(name) {
  const safe = name.replace(/\.[^.]+$/, "").replace(/[<>:"/\\|?*\u0000-\u001f]/g, "_")
    .replace(/[. ]+$/, "").slice(0, 120);
  return safe || "Photoshop";
}

class OutputFiles {
  constructor(utf8) { this.utf8 = utf8; }
  async publish(folder, documentName, xml, signal) {
    const stem = baseName(documentName);
    const entries = await folder.getEntries();
    const names = new Set(entries.map(entry => entry.name.toLowerCase()));
    let name = `${stem}.pagx`;
    for (let n = 2; names.has(name.toLowerCase()); n++) name = `${stem} (${n}).pagx`;
    checkCancelled(signal);
    const tempName = `.pagx-export-${Date.now()}-${Math.random().toString(16).slice(2)}.tmp`;
    const temp = await folder.createFile(tempName, { overwrite: false });
    let published = false;
    try {
      await temp.write(xml, { format: this.utf8 });
      checkCancelled(signal);
      // Same-directory rename; a racing writer must not be overwritten.
      await temp.moveTo(folder, { newName: name, overwrite: false });
      published = true;
      return { name, nativePath: temp.nativePath };
    } finally {
      if (!published) {
        try { await temp.delete(); } catch (error) { console.warn("临时导出文件清理失败", error); }
      }
    }
  }
}

module.exports = { OutputFiles, baseName };
