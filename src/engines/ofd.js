/**
 * 引擎适配层 —— OFD（国产版式文档，Open Fixed-layout Document）
 *
 * 用 `@file-viewer/renderer-ofd`（Apache-2.0）。
 */

import { renderFileViewerOfd } from "@file-viewer/renderer-ofd";

/** 清空容器。 */
function reset(host) {
  while (host.firstChild) host.removeChild(host.firstChild);
}

/**
 * 渲染 OFD 文档。
 *
 * @param bytes - 完整文件字节。
 * @param host - 承载渲染结果的容器。
 * @param ctx - `{ name }` 文件名。
 */
export async function renderOfd(bytes, host, ctx = {}) {
  reset(host);
  const buffer = bytes.buffer.slice(bytes.byteOffset, bytes.byteOffset + bytes.byteLength);
  await renderFileViewerOfd(buffer, host, "ofd", { filename: ctx.name });
}
