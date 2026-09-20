/**
 * dsh-office-viewer —— 浏览器半（构建产物入口，源码）
 *
 * 职责：用官方扩展点 `ctx.documentPreviews.register()` 接管 Office 文档的侧栏预览，
 * 声明 `loading: "bytes-complete"` —— 宿主把**整份文件字节**交给本渲染器，
 * 全程在浏览器内解析渲染，**不依赖任何外部服务**（无 office-preview / LibreOffice）。
 *
 * 覆盖格式（"办公文件预览"）：
 *   · OOXML      : docx / xlsx / pptx
 *   · 旧二进制    : doc / xls / ppt        （OLE2 复合文档）
 *   · WPS 两代    : wps / et / dps         （旧=MS 格式，新=OOXML，按签名分派）
 *   · 其他        : ofd / rtf / csv
 *
 * 明确不接管（官方已内置，重复注册是负收益）：
 *   md / markdown / html / htm / pdf / 图片 / 纯文本
 *
 * 资产：Worker / WASM / 字体由宿主半通过 `/office-viewer-assets/*` 路由发布，
 * 本文件内通过 ASSET_BASE 引用（这就是"免服务"的落点）。
 */

import { renderDocx } from "./engines/docx.js";
import { renderSpreadsheet } from "./engines/xlsx.js";
import { renderPresentation } from "./engines/pptx.js";
import { renderLegacyDoc, renderRtf } from "./engines/doc.js";
import { renderOfd } from "./engines/ofd.js";

/** 本插件在宿主侧注册的资产路由前缀（与 lib/index.js 保持一致）。 */
export const ASSET_BASE = "/office-viewer-assets";

/**
 * 本插件接管的扩展名 → 引擎标签。
 *
 * 注意 `wps`/`et`/`dps`：WPS 有两代产物 —— 旧版是 MS 二进制（OLE2），新版就是 OOXML。
 * 这里统一登记，实际由 {@link sniffFormat} 按文件头签名分派到正确引擎。
 */
export const EXTENSIONS = [
  "docx", "doc",
  "xlsx", "xls", "csv",
  "pptx", "ppt",
  "wps", "et", "dps",
  "ofd", "rtf",
];

/** OLE2 复合文档头（旧版 MS/WPS 二进制的标志）。 */
const OLE2_MAGIC = [0xd0, 0xcf, 0x11, 0xe0, 0xa1, 0xb1, 0x1a, 0xe1];
/** ZIP 头（OOXML：docx/xlsx/pptx，以及新版 WPS）。 */
const ZIP_MAGIC = [0x50, 0x4b];
/** RTF 头（`{\rtf`）。 */
const RTF_MAGIC = [0x7b, 0x5c, 0x72, 0x74, 0x66];

/** 判断字节数组是否以 `magic` 开头。 */
function startsWith(bytes, magic) {
  if (bytes.length < magic.length) return false;
  for (let i = 0; i < magic.length; i += 1) {
    if (bytes[i] !== magic[i]) return false;
  }
  return true;
}

/**
 * 按文件头签名判断真实容器类型。
 *
 * 生产背景：扩展名不可信 —— 用户常把 .doc 改名成 .docx，WPS 的 .wps 也有新旧两代。
 * 因此**始终以签名为准**，扩展名只用来决定"解析成哪个语种"。
 *
 * @param bytes - 文件前若干字节（至少 8 字节）。
 * @returns `'ole2'` | `'zip'` | `'rtf'` | `'unknown'`。
 */
export function sniffContainer(bytes) {
  if (startsWith(bytes, OLE2_MAGIC)) return "ole2";
  if (startsWith(bytes, ZIP_MAGIC)) return "zip";
  if (startsWith(bytes, RTF_MAGIC)) return "rtf";
  return "unknown";
}

/**
 * 决定用哪个引擎渲染。
 *
 * 分派规则（签名优先，扩展名兜底）：
 *   · ZIP 容器  → OOXML 引擎族；但 `et` 走表格、`dps` 走演示
 *   · OLE2 容器 → 旧二进制引擎族；`wps` 走 doc、`et` 走 xls、`ppt/dps` 走 ppt
 *   · RTF 头    → rtf 引擎
 *
 * @param ext - 小写扩展名（不含点）。
 * @param bytes - 文件前若干字节。
 * @returns 引擎标识：`'docx'|'xlsx'|'pptx'|'doc'|'xls'|'ppt'|'ofd'|'rtf'|'csv'|'unsupported'`。
 */
export function resolveEngine(ext, bytes) {
  const container = sniffContainer(bytes);

  if (container === "rtf" || ext === "rtf") return "rtf";
  if (ext === "ofd") return "ofd";
  if (ext === "csv") return "csv";

  // WPS 三件套：按签名分派到所属引擎族（旧二进制 / 新 OOXML 都能吃）
  if (ext === "wps") return container === "zip" ? "docx" : "doc";
  if (ext === "et") return container === "zip" ? "xlsx" : "xls";
  if (ext === "dps") return container === "zip" ? "pptx" : "ppt";

  if (container === "zip") {
    if (ext === "docx" || ext === "doc") return "docx";
    if (ext === "xlsx" || ext === "xls") return "xlsx";
    if (ext === "pptx" || ext === "ppt") return "pptx";
  }
  if (container === "ole2") {
    if (ext === "doc" || ext === "docx") return "doc";
    if (ext === "xls" || ext === "xlsx") return "xls";
    if (ext === "ppt" || ext === "pptx") return "ppt";
  }
  return "unsupported";
}

/**
 * 把字节交给对应引擎渲染到 `host`。
 *
 * @param engine - {@link resolveEngine} 的返回值。
 * @param bytes - 完整文件字节。
 * @param host - 承载渲染结果的 DOM 容器。
 * @param ctx - `{ name, signal }`：文件名与中止信号。
 * @returns 渲染完成时 resolve。
 */
export async function renderByEngine(engine, bytes, host, ctx = {}) {
  switch (engine) {
    case "docx":
      return renderDocx(bytes, host, ctx);
    case "xlsx":
    case "csv":
      return renderSpreadsheet(bytes, host, ctx);
    case "pptx":
      return renderPresentation(bytes, host, ctx);
    case "doc":
      return renderLegacyDoc(bytes, host, ctx);
    case "xls":
      return renderSpreadsheet(bytes, host, { ...ctx, legacy: true });
    case "ppt":
      return renderPresentation(bytes, host, { ...ctx, legacy: true });
    case "rtf":
      return renderRtf(bytes, host, ctx);
    case "ofd":
      return renderOfd(bytes, host, ctx);
    default:
      throw new Error(`暂不支持的格式：${ctx.name ?? "未知文件"}`);
  }
}
