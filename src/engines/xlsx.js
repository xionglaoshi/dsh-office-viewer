/**
 * 引擎适配层 —— 表格系（xlsx / 旧 xls / csv）
 *
 * 用 `@file-viewer/renderer-spreadsheet`（其 Worker 内正是 SheetJS，
 * CE 版**未阉割 BIFF8**，所以旧 .xls 也能读）。
 */

import { renderFileViewerSpreadsheet } from "@file-viewer/renderer-spreadsheet";

/** 清空容器，避免切换文件时残留上一次渲染。 */
function reset(host) {
  while (host.firstChild) host.removeChild(host.firstChild);
}

/**
 * 渲染表格文件（xlsx / xls / csv）。
 *
 * @param bytes - 完整文件字节。
 * @param host - 承载渲染结果的容器。
 * @param ctx - `{ name, legacy }`：文件名；`legacy` 仅用于记录日志，引擎按签名自行分辨。
 */
export async function renderSpreadsheet(bytes, host, ctx = {}) {
  reset(host);
  const buffer = bytes.buffer.slice(bytes.byteOffset, bytes.byteOffset + bytes.byteLength);
  await renderFileViewerSpreadsheet(buffer, host, ctx.legacy ? "xls" : "xlsx", {
    filename: ctx.name,
  });
}
