/**
 * 引擎适配层 —— 旧版二进制 Word 系（.doc / .wps）与长尾格式（rtf / ofd）
 *
 * 旧 `.doc` 是 OLE2 复合文档（CFB），**不是 ZIP**，与 docx 结构完全不同：
 *   CFB → FIB（文件信息块）→ 文本流 / 表格流（CLX/FKP）→ PAPX 段落属性 → HTML。
 *
 * 实现走 `@file-viewer/doc`（MIT，无水印），或经 `@file-viewer/renderer-word`
 * 暴露的 `renderFileViewerWordDoc` —— 它内部会按签名自行分辨
 * OOXML / WordML / 二进制三条路，所以扩展名与实际内容不符也能正确渲染。
 */

import { renderFileViewerWordDoc } from "@file-viewer/renderer-word";

/** 清空容器。 */
function reset(host) {
  while (host.firstChild) host.removeChild(host.firstChild);
}

/** 把 Uint8Array 转成独立的 ArrayBuffer（引擎要求 ArrayBuffer 输入）。 */
function toArrayBuffer(bytes) {
  return bytes.buffer.slice(bytes.byteOffset, bytes.byteOffset + bytes.byteLength);
}

/**
 * 渲染旧版二进制 .doc / .wps。
 *
 * @param bytes - 完整文件字节。
 * @param host - 承载渲染结果的容器。
 * @param ctx - `{ name }` 文件名。
 */
export async function renderLegacyDoc(bytes, host, ctx = {}) {
  reset(host);
  await renderFileViewerWordDoc(toArrayBuffer(bytes), host, "doc", { filename: ctx.name });
}

/**
 * 渲染 RTF。
 *
 * RTF 是纯文本标记（`{\rtf1...}`），由 Word 引擎的 RTF 通道处理。
 *
 * @param bytes - 完整文件字节。
 * @param host - 承载渲染结果的容器。
 * @param ctx - `{ name }` 文件名。
 */
export async function renderRtf(bytes, host, ctx = {}) {
  reset(host);
  await renderFileViewerWordDoc(toArrayBuffer(bytes), host, "rtf", { filename: ctx.name });
}
