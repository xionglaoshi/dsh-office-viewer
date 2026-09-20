/**
 * dsh-office-viewer（办公文件预览）—— 宿主半（Node）
 *
 * 职责：**只做一件事 —— 发布本插件自带的渲染器静态资产**（Worker / WASM / CJK 字体）。
 *
 * 🔴 这不是"启动一个服务"：
 *   · 它不监听任何端口，只是往 DSH 自己的 webServer 上挂一条路由
 *     （`ctx.webServer.register`，与官方 open-in-app / client-connection 同一机制）；
 *   · DSH 的 Web GUI 本来就跑在 3080 上，本插件只是借它的路由表分发自己目录下的文件；
 *   · **不需要 office-preview（8765）、不需要 LibreOffice、不需要任何第三方进程**。
 *
 * 为什么需要它：浏览器半的渲染引擎（pptx / 旧 ppt）要加载 Worker 与 WASM，
 * 这些二进制无法内联进 JS bundle，必须由 HTTP 提供。官方插件机制只自动服务
 * `/plugins/<id>/client.js`，不覆盖额外资产，因此本路由是必需的。
 */

import { createReadStream, existsSync, statSync } from "node:fs";
import { createHash } from "node:crypto";
import { extname, join, normalize, sep } from "node:path";
import { fileURLToPath } from "node:url";

/** 本插件在 DSH 里注册的资产路由前缀（须与浏览器半 lib/client.js 的 ASSET_BASE 一致）。 */
export const ASSET_ROUTE = "/office-viewer-assets";

/** 资产目录：<包根>/assets。 */
const PACKAGE_ROOT = fileURLToPath(new URL("..", import.meta.url));
const ASSET_ROOT = join(PACKAGE_ROOT, "assets");

/** 扩展名 → Content-Type。 */
const MIME = {
  ".js": "text/javascript; charset=utf-8",
  ".mjs": "text/javascript; charset=utf-8",
  ".wasm": "application/wasm",
  ".otf": "font/otf",
  ".ttf": "font/ttf",
  ".woff2": "font/woff2",
  ".json": "application/json; charset=utf-8",
  ".css": "text/css; charset=utf-8",
  ".map": "application/json; charset=utf-8",
};

/**
 * 把 URL 路径解析为资产目录内的真实文件路径。
 *
 * 安全闸：拒绝任何越出 `ASSET_ROOT` 的路径（`..`、绝对路径、编码绕过），
 * 防止本路由变成任意文件读取口子。
 *
 * @param pathname - 请求的 URL pathname。
 * @returns 绝对文件路径；不合法或越界时返回 undefined。
 */
export function resolveAssetPath(pathname) {
  let rel;
  try {
    rel = decodeURIComponent(String(pathname).slice(ASSET_ROUTE.length));
  } catch {
    return undefined; // 非法百分号编码
  }
  if (rel === "" || rel === "/") return undefined;

  // 先规范化再比对前缀，挡住 ../ 穿越
  const candidate = normalize(join(ASSET_ROOT, rel));
  const rootWithSep = ASSET_ROOT.endsWith(sep) ? ASSET_ROOT : ASSET_ROOT + sep;
  if (candidate !== ASSET_ROOT && !candidate.startsWith(rootWithSep)) return undefined;

  try {
    if (!existsSync(candidate) || !statSync(candidate).isFile()) return undefined;
  } catch {
    return undefined;
  }
  return candidate;
}

/**
 * 处理一次资产请求。
 *
 * 资产内容随 npm 包固定不变，因此给长缓存 + 强 ETag；
 * DSH 的 GUI 是本机回环访问，无需额外鉴权。
 *
 * @param req - Node 请求对象。
 * @param res - Node 响应对象。
 */
function serveAsset(req, res) {
  if (req.method !== "GET" && req.method !== "HEAD") {
    res.statusCode = 405;
    res.setHeader("allow", "GET, HEAD");
    res.end();
    return;
  }

  const pathname = new URL(String(req.url), "http://localhost").pathname;
  const file = resolveAssetPath(pathname);
  if (file === undefined) {
    res.statusCode = 404;
    res.setHeader("content-type", "text/plain; charset=utf-8");
    res.end("not found");
    return;
  }

  const { size, mtimeMs } = statSync(file);
  const etag = `"${createHash("sha1").update(`${file}:${size}:${mtimeMs}`).digest("hex").slice(0, 16)}"`;

  res.setHeader("content-type", MIME[extname(file).toLowerCase()] ?? "application/octet-stream");
  res.setHeader("cache-control", "public, max-age=31536000, immutable");
  res.setHeader("etag", etag);
  res.setHeader("content-length", String(size));
  // Worker + WASM 需要同源；本路由与 GUI 同源，天然满足。

  if (req.headers["if-none-match"] === etag) {
    res.statusCode = 304;
    res.end();
    return;
  }
  if (req.method === "HEAD") {
    res.statusCode = 200;
    res.end();
    return;
  }

  res.statusCode = 200;
  const stream = createReadStream(file);
  stream.on("error", () => {
    if (!res.headersSent) res.statusCode = 500;
    res.end();
  });
  stream.pipe(res);
}

/** 宿主半插件体。 */
export function apply(ctx) {
  ctx.effect(
    () =>
      ctx.webServer.register({
        kind: "prefix",
        path: ASSET_ROUTE,
        handler: serveAsset,
      }),
    `office-viewer: GET ${ASSET_ROUTE}/*`
  );
}

/** 依赖 DSH 的 webServer 服务（Web 载体）。 */
export const inject = ["webServer"];
