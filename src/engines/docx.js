/**
 * 引擎适配层 —— Word 系（docx / 旧 doc / rtf / ofd 兜底）
 *
 * 设计：统一收敛成 `(bytes, host, ctx) => Promise<void>`，让 format.js 的分派表保持扁平。
 * 这里只做"调用 + 清理 + 错误包装"，不实现解析（解析是 @file-viewer 引擎的职责）。
 */

import { renderFileViewerWordDocx } from "@file-viewer/renderer-word";

/** 清空容器，避免切换文件时残留上一次渲染。 */
function reset(host) {
  while (host.firstChild) host.removeChild(host.firstChild);
}

/**
 * 渲染 OOXML（.docx，以及伪装成 .doc 的 ZIP 文档）。
 *
 * 引擎自身的 `resolveFileViewerWordContainer` 会按签名再判一次容器，
 * 所以即使扩展名与真实内容不符（用户改名、WPS 存成 .doc）也能正确分流。
 *
 * @param bytes - 完整文件字节。
 * @param host - 承载渲染结果的容器。
 * @param ctx - `{ name }` 文件名，仅用于报错文案。
 */
export async function renderDocx(bytes, host, ctx = {}) {
  reset(host);
  const buffer = bytes.buffer.slice(bytes.byteOffset, bytes.byteOffset + bytes.byteLength);
  await renderFileViewerWordDocx(buffer, host, "docx", { filename: ctx.name });
}
