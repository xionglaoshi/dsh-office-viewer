/**
 * 引擎适配层 —— 演示系（pptx / 旧 ppt）
 *
 * pptx：`@file-viewer/pptx` 的 `PptxViewer.open()`，还原度高（含放映、缩略图、虚拟滚动）。
 * 旧 ppt：`@file-viewer/ppt`（专有公开二进制版，**带原生水印**，见 README 的许可证声明）。
 *
 * ⚠️ 水印说明：`@file-viewer/ppt` 的 manifest 声明
 *   `watermarkEnforcement: "native-final-frame"` —— 水印由 WASM 绘制在最终帧上，
 *   **不是 DOM/CSS 层，无法通过改样式去除**。去水印需向 Flyfish 购买商业授权。
 *   本插件按用户裁决接受该水印；仅旧版 `.ppt/.dps` 受影响，pptx 不受影响。
 */

import { PptxViewer } from "@file-viewer/pptx";
import { createPptViewer } from "@file-viewer/ppt";
import { ASSET_BASE } from "../format.js";

/** 已挂载实例，切换文件时先销毁，避免 Worker 泄漏。 */
let live = null;

/** 清空容器并销毁上一个实例。 */
async function reset(host) {
  if (live) {
    try {
      await live.destroy?.();
    } catch {
      /* 销毁失败不影响后续渲染 */
    }
    live = null;
  }
  while (host.firstChild) host.removeChild(host.firstChild);
}

/**
 * 渲染演示文稿（pptx / 旧 ppt）。
 *
 * Worker 走宿主半发布的资产路由 `ASSET_BASE`（这是"免服务"的落点：
 * Worker/WASM/字体由 DSH 的 webServer 直接提供，不经任何第三方服务）。
 *
 * @param bytes - 完整文件字节。
 * @param host - 承载渲染结果的容器。
 * @param ctx - `{ name, legacy }` 文件名与是否旧格式。
 */
export async function renderPresentation(bytes, host, ctx = {}) {
  await reset(host);

  if (ctx.legacy) {
    return renderLegacyPpt(bytes, host, ctx);
  }

  const buffer = bytes.buffer.slice(bytes.byteOffset, bytes.byteOffset + bytes.byteLength);
  live = await PptxViewer.open(buffer, host, {
    workerUrl: `${ASSET_BASE}/pptx/pptx.worker.js`,
    lazySlides: true,
    lazyMedia: true,
    fitMode: "contain",
    onWarning: (w) => console.warn("[office-viewer] pptx warning:", w),
  });
}

/**
 * 旧版二进制 .ppt（PowerPoint 97-2003，OLE2 容器）。
 *
 * 该引擎体量大（WASM + 16 MB CJK 字体包），因此**按需动态 import + 按需加载运行时**
 * —— 只有真的打开旧 .ppt 时才拉取，不拖累常见格式的首屏。
 *
 * 资产全部走宿主半发布的 `ASSET_BASE` 路由（这是"免服务"的落点：
 * WASM / Worker / 字体由 DSH 自己的 webServer 提供，不经任何第三方服务）。
 *
 * @param bytes - 完整文件字节。
 * @param host - 承载渲染结果的容器。
 * @param ctx - `{ name }` 文件名。
 */
async function renderLegacyPpt(bytes, host, ctx = {}) {
  const runtime = await createPptViewer({
    wasmUrl: `${ASSET_BASE}/ppt/ppt-native.wasm`,
    fontUrl: `${ASSET_BASE}/ppt/ppt-font-cjk.otf`,
    workerUrl: `${ASSET_BASE}/ppt/worker.mjs`,
    worker: "auto",
  });

  const buffer = bytes.buffer.slice(bytes.byteOffset, bytes.byteOffset + bytes.byteLength);
  const mounted = await runtime.mount(host, buffer);
  live = {
    destroy: async () => {
      await mounted.close();
      await runtime.close();
    },
  };
}
