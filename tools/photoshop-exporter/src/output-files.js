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
  async publish(folder, documentName, xml, signal, report) {
    const stem = baseName(documentName);
    const entries = await folder.getEntries();
    const names = new Set(entries.map(entry => entry.name.toLowerCase()));
    let name = `${stem}.pagx`;
    for (let n = 2; names.has(name.toLowerCase()) ||
      (report && names.has(`${name}.report.json`.toLowerCase())); n++) name = `${stem} (${n}).pagx`;
    checkCancelled(signal);
    const tempName = `.pagx-export-${Date.now()}-${Math.random().toString(16).slice(2)}.tmp`;
    const temp = await folder.createFile(tempName, { overwrite: false });
    let reportFile;
    let published = false;
    try {
      await temp.write(xml, { format: this.utf8 });
      if (report) {
        reportFile = await folder.createFile(`${tempName}.report`, { overwrite: false });
        await reportFile.write(JSON.stringify({ ...report, outputFile: name }, null, 2), { format: this.utf8 });
      }
      checkCancelled(signal);
      // Publish the report first, PAGX last. If the final rename fails, remove our report as well.
      // This is a rollback-capable pair, not a filesystem-wide atomic transaction.
      if (reportFile) await reportFile.moveTo(folder, { newName: `${name}.report.json`, overwrite: false });
      checkCancelled(signal);
      // Same-directory rename; a racing writer must not be overwritten.
      await temp.moveTo(folder, { newName: name, overwrite: false });
      published = true;
      return { name, nativePath: temp.nativePath,
        reportPath: reportFile && reportFile.nativePath };
    } finally {
      if (!published) {
        try { await temp.delete(); } catch (error) { console.warn("临时导出文件清理失败", error); }
        if (reportFile) {
          try { await reportFile.delete(); } catch (error) { console.warn("导出报告清理失败", error); }
        }
      }
    }
  }
}

module.exports = { OutputFiles, baseName };
