# dsh-office-viewer · 办公文件预览

给 [DSH（DeepSeek Harness）](https://github.com/deepseek-ai/deepseek-harness) 的侧栏加一个**免服务**的 Office 文档预览器：
docx / xlsx / pptx、旧版 doc / xls / ppt、WPS 三件套（wps / et / dps）、以及 ofd / rtf / csv。

**核心特点：不启动任何外部服务。** 不需要 office-preview，不需要 LibreOffice，不需要 OnlyOffice。
所有解析都在浏览器里完成。

---

## 为什么可以免服务

DSH 的文档预览有官方扩展点，插件可以声明内容的投递方式：

```ts
ctx.documentPreviews.register({
  id: "dsh-office-viewer/office",
  extensions: ["docx", "xlsx", "pptx", "doc", "xls", "ppt", "wps", "et", "dps", "ofd", "rtf", "csv"],
  priority: "extension",       // 高于内置渲染器
  loading: "bytes-complete",   // ← 关键：宿主把整份文件字节交给渲染器
  title: () => "办公文件预览",
});
```

`loading: "bytes-complete"` 意味着宿主把文件完整读成字节数组交给浏览器端渲染器，
渲染器在自己的 JS/WASM 里解析。官方内置的 PDF 预览（pdfjs）走的就是同一条路。

唯一的例外是需要 Worker / WASM 的引擎（pptx、旧 ppt）：这些二进制无法内联进 JS，
本插件的**宿主半**会把它们挂到 DSH 自己的 webServer 上的一条路由
（`/office-viewer-assets/*`）—— 这**不是**另起一个端口或进程，
DSH 的 GUI 本来就跑在某个端口上，插件只是借用它的路由表分发自己目录下的文件。

---

## 支持的格式

| 格式 | 引擎 | 许可证 | 保真度 |
|---|---|---|---|
| `docx` | `@file-viewer/renderer-word` | Apache-2.0 | 高（版面、表格、图片、页眉页脚） |
| `xlsx` `csv` | `@file-viewer/renderer-spreadsheet`（内含 SheetJS） | Apache-2.0 | 高（网格、样式、多 sheet） |
| `pptx` | `@file-viewer/pptx` | Apache-2.0 | 高（含放映、缩略图、虚拟滚动） |
| `doc` | `@file-viewer/doc` | **MIT** | 高（CFB → FIB → PAPX → HTML） |
| `xls` | SheetJS | Apache-2.0 | 高（BIFF8 完整支持） |
| `ppt` | `@file-viewer/ppt` | ⚠️ **专有，带水印** | 高（WASM 原生引擎） |
| `wps` `et` `dps` | 复用上面（按签名分派） | — | 见下 |
| `ofd` | `@file-viewer/renderer-ofd` | Apache-2.0 | 中 |
| `rtf` | `@file-viewer/renderer-word` | Apache-2.0 | 中 |

### WPS 三件套怎么处理的

WPS 有两代产物，本插件按**文件头签名**自动分派，两代都能看：

- **旧版 WPS**（`D0CF11E0`，OLE2 复合文档）—— 实际就是 MS 格式：
  `.wps` 内含 `WordDocument` + `0Table` 流（= MS-DOC），`.et` 内含 `Workbook` 流（= BIFF8）。
  → 分别走 `doc` / `xls` 引擎。
- **新版 WPS**（`504B0304`，ZIP）—— 实际就是 OOXML：
  `.wps` = `word/document.xml`，`.et` = `xl/workbook.xml`，`.dps` = `ppt/presentation.xml`。
  → 零新增依赖，直接走 OOXML 引擎。

### 不接管的格式

`md` / `markdown` / `html` / `htm` / `pdf` / 图片 / 纯文本 —— **官方已内置**，
再注册一遍是重复劳动且可能更差（官方 PDF 用 pdfjs）。本插件不碰这些。

---

## 安装

```bash
# 1) 获取源码（放在 DSH 安装树之外，升级 DSH 不会丢）
git clone https://github.com/xionglaoshi/dsh-office-viewer.git ~/.dsh/packages/dsh-office-viewer
cd ~/.dsh/packages/dsh-office-viewer

# 2) 装依赖并构建（生成 lib/client.js 与 assets/）
npm install
npm run build

# 3) 链进 web profile
cd ~/.dsh/profiles/web
npm install link:/Users/$USER/.dsh/packages/dsh-office-viewer
```

然后在 `~/.dsh/cordis.patch.yml` 追加插件行：

```yaml
- insert:
    - id: office-viewer
      name: dsh-office-viewer
```

改完**先验证再重启**（DSH 的硬性要求）：

```bash
dsh --profile web --dump-config   # 退出码应为 0
```

重启宿主后，在侧栏点开任意 docx / xls / ppt 即可。

---

## 构建产物

`npm run build` 做两件事：

1. **打包 `src/` → `lib/client.js`**（约 2 MB，自包含）
   用 esbuild 打成单文件 IIFE，包进 DSH 的 `window.__ModuleLoader__.load({id, factory})` 外壳。
   只有 `react` 保持 external（由 DSH 宿主提供）。

2. **复制引擎资产 → `assets/`**（约 19 MB，其中 16 MB 是旧 ppt 的 CJK 字体）
   这些由宿主半经 `/office-viewer-assets/*` 发布。

---

## 自测

不需要启动 DSH，用真实文档直接验证渲染管线：

```bash
node scripts/selftest.mjs ~/wps          # 传入含 Office 文件的目录
```

输出示例（本机实测，真实中文文档）：

```
样本：doc, docx, ppt, pptx, xls, xlsx
✔ doc    engine=doc            13ms  text=  4194  canvas=0
✔ docx   engine=docx           75ms  text=  5472  canvas=0
✔ ppt    engine=ppt           212ms  text=     0  canvas=2
✔ pptx   engine=pptx          263ms  text=   517  canvas=0
✔ xls    engine=xls            31ms  text=    80  canvas=1
✔ xlsx   engine=xlsx           13ms  text=    60  canvas=1

全部通过 ✅
```

---

## 已知限制

- **旧版 `.ppt` 带水印。** 见下节。去水印需向 Flyfish 购买商业授权。
- **xlsx 图表不还原** —— 所有纯 JS 方案都只能读出图表定义 XML，不能绘制。这是行业普遍天花板。
- **xlsx 公式只读缓存值**，不重算（WPS/Excel 已算好并存在文件里，通常够用）。
- **pptx / ppt 不播放动画**，只做静态逐页渲染。
- **不自带中文字体**（旧 ppt 除外，它自带 CJK 字体包），中文依赖系统字体回退。
- 旧版 `.ppt` 引擎使用 `OffscreenCanvas`，Safari 的兼容性未经验证。

---

## 许可证

**本插件自身代码（`src/`、`lib/index.js`、`scripts/`）：MIT**，见 [LICENSE](./LICENSE)。

**捆绑与依赖的第三方引擎各自适用其自身许可证**，本插件的 MIT 不覆盖它们：

| 组件 | 许可证 | 说明 |
|---|---|---|
| `@file-viewer/core`、`renderer-word`、`renderer-spreadsheet`、`pptx`、`doc`、`renderer-ofd` | Apache-2.0 | 可自由分发 |
| SheetJS (`xlsx`) | Apache-2.0 | 可自由分发 |
| **`@file-viewer/ppt`** | **Flyfish Public Watermarked Runtime License v2（专有）** | ⚠️ 见下 |

### ⚠️ 关于 `@file-viewer/ppt`（旧版 .ppt 引擎）

这是本仓库里**唯一非开源**的组件，其 `manifest.json` 明确声明：

```json
{
  "edition": "public-watermarked",
  "distribution": "proprietary-public-binary",
  "apacheLicenseApplies": false,
  "sourceCodeRights": false,
  "watermarkRequired": true,
  "watermarkText": "Flyfish Viewer",
  "watermarkEnforcement": "native-final-frame",
  "watermarkRemovalRequiresCommercialAuthorization": true,
  "commercialUse": true,
  "redistributionAllowed": true,
  "redistributionPolicy": "unmodified-integrated-bundling-only"
}
```

需要知道的几点：

1. **水印去不掉**：`watermarkEnforcement: "native-final-frame"` 表示水印由 WASM
   绘制在**最终帧**上，不在 DOM/CSS 层 —— 改样式或删节点都无法去除。
2. **商业使用是允许的**（`commercialUse: true`），前提是保留水印。
3. **可随本项目一起分发**（`redistributionAllowed: true`），但必须是
   "unmodified-integrated-bundling-only"：不能单独再分发、不能修改。
4. **只影响旧版 `.ppt` / `.dps`**，docx / xlsx / pptx / doc / xls 等全部链路都不受影响
   （`@file-viewer/doc` 是干净的 MIT，实测其包内无水印代码）。
5. 若你不接受该水印，可删除 `package.json` 里的 `@file-viewer/ppt` 依赖后重新构建，
   旧 `.ppt` 将显示"暂不支持"，其余格式不受影响。

---

## 致谢

渲染能力全部来自这些项目，本插件只做集成与 DSH 适配：

- [file-viewer](https://github.com/flyfish-dev/file-viewer)（flyfish-dev）
- [SheetJS](https://sheetjs.com/)
- [docxjs / docx-preview](https://github.com/volodymyrbaydalka/docxjs)

## 许可

MIT © 2026 熊翔
