# 设计说明

本文件记录几个不显然的设计决策及其依据，供后续维护参考。

## 1. 为什么用官方扩展点而不是自己开预览面板

DSH 的侧栏文档预览有官方扩展点，契约（读自
`@deepseek-ai/dsh-client-ui-sidebar-documentpreview/lib/types/client/document/registry.d.ts`）：

```ts
export interface DocumentPreviewDefinition {
  readonly id: string;
  readonly extensions: readonly string[];
  readonly priority?: 'builtin' | 'extension';
  readonly title: () => string;
  readonly loading: 'text-pages' | 'bytes-complete';
  readonly wrap?: boolean;
}
```

两个关键点：

- **`loading: "bytes-complete"`** 让宿主把整份文件读成字节交给渲染器 —— 这是
  "免服务"的**全部技术前提**。官方内置的 PDF 渲染器（pdfjs）用的就是它。
- **`priority: "extension"`** 使插件优先于官方 `builtin` 渲染器。

排序规则（`matchingDocumentPreviews`）：先比 `rank`（extension=1 > builtin=0），
再比**匹配到的扩展名长度**，最后比注册顺序。

## 2. 为什么只登记部分扩展名

官方已内置：`md`、`markdown`、`html`、`htm`、`pdf`、图片，以及一个
`extensions: []` 的**纯文本兜底**（它匹配任何未识别的扩展名）。

因此本插件**不登记** md/txt/html/pdf —— 重复注册不但没收益，还可能覆盖掉
更合适的官方实现（例如 pdf 用 pdfjs 明显优于任何自研方案）。

## 3. 为什么宿主半需要一条资产路由

浏览器半是完全自包含的（约 2 MB，零外部 `require`），但它**不能**自包含
Worker 与 WASM：

- pptx 引擎要求显式提供 `workerUrl`，否则直接抛
  `PPTX Worker URL is unavailable`；
- 旧 ppt 引擎需要一个 WASM（约 1.5 MB）与一个 CJK 字体包（约 16 MB）。

这些二进制无法内联进 JS。而官方插件机制**只自动服务 `/plugins/<id>/client.js`**
（见 `@deepseek-ai/dsh-client-modules` 的 `serveBundle`），不覆盖额外资产。

所以宿主半注册一条自己的路由：

```js
ctx.webServer.register({ kind: "prefix", path: "/office-viewer-assets", handler });
```

**这不是"起一个服务"**：不监听端口、不占进程，只是往 DSH 已有的 webServer
路由表里加一条。官方 `@deepseek-ai/dsh-host-open-in-app` 等服务也用同一机制。

### 安全

`resolveAssetPath()` 做了路径穿越防护：规范化后必须仍在 `assets/` 内，
否则返回 404。已实测拦住 `../`、`%2e%2e`、绝对路径等构造。

## 4. 为什么按文件头签名分派，而不是按扩展名

扩展名不可信，有两个现实原因：

1. **用户改后缀**：把 `.doc` 改名成 `.docx` 很常见；
2. **WPS 两代产物**：旧版 `.wps/.et/.dps` 实际是 MS 二进制格式（OLE2），
   新版则是标准 OOXML（ZIP）。

因此 `resolveEngine()` 一律先看前 8 字节：

| 签名 | 容器 | 分派 |
|---|---|---|
| `D0CF11E0A1B11AE1` | OLE2 复合文档 | `doc` / `xls` / `ppt` 引擎族 |
| `504B0304` | ZIP | OOXML 引擎族 |
| `7B5C727466` (`{\rtf`) | RTF 文本 | rtf 引擎 |

额外好处：`@file-viewer/renderer-word` 的 `renderFileViewerWordDoc` 内部也会
再按签名分辨 OOXML / WordML / 二进制，所以即使分派有偏差，Word 系仍然安全。

## 5. 构建：为什么用 esbuild + 自己写 `__ModuleLoader__` 外壳

DSH 客户端插件的契约是：

```js
window.__ModuleLoader__.load({ id, factory });
// factory(require) => 模块导出；require 只能取宿主共享包（react 等）
```

用 esbuild 打成 `format: "iife"` + `globalName`，再把生成的
`var <globalName> = (() => {...})();` 改写成 `const __result = ...`，
使导出留在 factory 闭包内、不污染 `window`。

两个踩过的坑：

1. **`import.meta.url` 必须显式 define。**
   引擎内部会 `new URL(import.meta.url)` 来推断自带资产位置。IIFE 没有
   `import.meta`，esbuild 默认替换成空对象 → `new URL(undefined)` 抛
   `Invalid URL`，**整个 bundle 在 factory 阶段就崩**（表现为侧栏一片空白）。
   解法：`define: { "import.meta.url": '"http://localhost/office-viewer-assets/module.js"' }`。
   本插件始终显式传 `workerUrl/wasmUrl`，所以该值只用于让初始化不抛错。

2. **不能用动态 `import()` 引引擎。**
   `await import("@file-viewer/ppt")` 会被保留为运行时导入，在
   `__ModuleLoader__` 环境里无法解析。必须改成静态 import 让它被打进 bundle
   （代价是 bundle 变大，但保证了自包含）。

## 6. 已知天花板

- **xlsx 图表不还原**：所有纯 JS 方案都只能读出图表定义 XML，不能绘制。
- **xlsx 公式不重算**：只读文件里缓存的公式结果值（对预览通常够用）。
- **pptx/ppt 不播动画**：只做静态逐页渲染。
- **旧 ppt 带水印**：见 README「许可证」。

这些是上游引擎的限制，不是集成问题。
