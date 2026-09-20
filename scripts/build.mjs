/**
 * 构建脚本：把 src/ 打包成 DSH 客户端可加载的 lib/client.js
 *
 * DSH 客户端 bundle 的契约（读自 @deepseek-ai/dsh-client-modules）：
 *   window.__ModuleLoader__.load({ id, factory })
 *   其中 factory(require) 返回模块导出；require 只能取宿主提供的共享包
 *   （react / react/jsx-runtime / @deepseek-ai/* 等）。
 *
 * 因此这里把 @file-viewer/* 全部**内联进同一个文件**（不能留外部 import），
 * 只把 `react` 保持为 external。
 */

import { build } from "esbuild";
import { mkdirSync, writeFileSync, readFileSync, cpSync, existsSync, rmSync } from "node:fs";
import { dirname, join } from "node:path";
import { fileURLToPath } from "node:url";

const ROOT = dirname(fileURLToPath(import.meta.url));
const PKG = join(ROOT, "..");
const OUT = join(PKG, "lib", "client.js");
const ASSETS = join(PKG, "assets");

/** 插件在 DSH 里的模块 id，必须与 package.json 的 name 一致。 */
const MODULE_ID = "dsh-office-viewer";

/** 把打包产物包进 __ModuleLoader__ 的外壳。 */
function wrap(bundleText) {
  return `/**
 * ${MODULE_ID} —— DSH 客户端 bundle（由 scripts/build.mjs 生成，勿手改）
 */
window.__ModuleLoader__.load({
  id: ${JSON.stringify(MODULE_ID)},
  factory: (require) => {
    var module = { exports: {} };
    var exports = module.exports;
${bundleText
  .split("\n")
  .map((line) => (line.length > 0 ? `    ${line}` : line))
  .join("\n")
  .replace(/^(\s*)var\s+__dshOfficeViewer\s*=\s*/m, "$1const __result = ")}
    if (typeof __result !== "undefined") module.exports = __result;
    return module.exports;
  },
});
`;
}

/** 把打包产物打成单文件。 */
async function bundle() {
  const result = await build({
    entryPoints: [join(PKG, "src", "index.js")],
    bundle: true,
    format: "iife",
    globalName: "__dshOfficeViewer",
    platform: "browser",
    target: ["chrome110"],
    minify: true,
    legalComments: "none",
    write: false,
    external: ["react", "react/jsx-runtime", "react-dom"],
    loader: { ".css": "text" },
    // 引擎内部会 `new URL(import.meta.url)` 来推断自带资产位置。
    // IIFE 产物没有 import.meta，esbuild 会替换成空对象 → `new URL(undefined)` 抛
    // "Invalid URL"，整个 bundle 在 factory 阶段就崩。这里给一个恒合法的绝对 URL：
    // 引擎仅在"未显式传 workerUrl/wasmUrl"时才用它，而本插件**始终显式传资产地址**
    // （见 src/engines/pptx.js），所以该值只用于让初始化不抛错。
    define: {
      "process.env.NODE_ENV": '"production"',
      "import.meta.url": '"http://localhost/office-viewer-assets/module.js"',
    },
    logLevel: "warning",
  });

  const code = result.outputFiles[0].text;

  mkdirSync(dirname(OUT), { recursive: true });
  // IIFE 产物形如 `var __dshOfficeViewer = (() => {...})();`
  // → 取其返回值直接交给 module.exports，避免依赖 esbuild 内部导出辅助函数。
  writeFileSync(OUT, wrap(code), "utf8");
  return OUT;
}

/** 复制引擎运行时资产（Worker / WASM / 字体）到 assets/，供宿主半发布。 */
function copyAssets() {
  rmSync(ASSETS, { recursive: true, force: true });
  mkdirSync(ASSETS, { recursive: true });

  const nm = join(PKG, "node_modules", "@file-viewer");
  const jobs = [
    // pptx Worker：引擎要求显式 workerUrl，缺它直接抛错
    [join(nm, "pptx", "dist", "worker", "pptx.worker.js"), "pptx/pptx.worker.js"],
  ];

  const copied = [];
  for (const [from, rel] of jobs) {
    if (!existsSync(from)) {
      console.warn(`  ⚠ 缺少资产（跳过）：${from}`);
      continue;
    }
    const to = join(ASSETS, rel);
    mkdirSync(dirname(to), { recursive: true });
    cpSync(from, to);
    copied.push(rel);
  }

  // 旧 ppt 引擎资产（专有、带水印）：装了才复制。
  // 注意 16 MB 的 CJK 字体包 —— 它决定了旧 .ppt 的首屏成本，故只在装了该引擎时携带。
  const pptDir = join(nm, "ppt");
  if (existsSync(pptDir)) {
    for (const f of [
      "ppt-native.wasm",
      "worker.mjs",
      "ppt-font-cjk.otf",
      "index.mjs",
      "frame-cache.mjs",
      "manifest.json",
      "LICENSE",
      "NOTICE",
    ]) {
      const from = join(pptDir, f);
      if (existsSync(from)) {
        cpSync(from, join(ASSETS, "ppt", f));
        copied.push(`ppt/${f}`);
      }
    }
  } else {
    console.warn("  ⚠ 未安装 @file-viewer/ppt —— 旧版 .ppt/.dps 将无法预览");
  }
  return copied;
}

const out = await bundle();
const assets = copyAssets();
const kb = (readFileSync(out).length / 1024).toFixed(0);
console.log(`✔ 已生成 ${out}（${kb} KB）`);
console.log(`✔ 资产 ${assets.length} 项：${assets.join(", ") || "（无）"}`);
