/**
 * dsh-office-viewer（办公文件预览）—— 浏览器半入口
 *
 * 用官方扩展点 `ctx.documentPreviews.register()` + 会话文档槽位接管 Office 文档预览：
 *   · `loading: "bytes-complete"` → 宿主把**整份文件字节**交给本渲染器
 *   · 全程浏览器内解析渲染，**不启动任何外部服务**（无 office-preview、无 LibreOffice）
 *
 * 注册优先级 `extension` 高于内置 `builtin`，因此本插件优先于官方渲染器生效；
 * 但我们**只登记官方未覆盖的扩展名**（md/txt/html/pdf/图片留给官方内置）。
 */

import { createElement, useEffect, useRef, useState } from "react";
import { EXTENSIONS, resolveEngine, renderByEngine, sniffContainer } from "./format.js";

/** 本渲染器的实现身份键（与槽位 key 共用）。 */
export const BODY_ID = "dsh-office-viewer/office";

/** 宿主提供的共享依赖：注册表、远程文件 API、槽位系统。 */
export const inject = ["documentPreviews", "remote", "remote.workspaceFiles", "slots"];

/** `dsh-resource://file/session/<sessionId>/<path>` 前缀（与官方同规则）。 */
const FILE_ADDRESS_PREFIX = "dsh-resource://file/";

/**
 * 供离线自测使用的内部导出。
 *
 * DSH 侧栏只依赖 `apply` / `inject` / `BODY_ID`；这里额外暴露分派与渲染函数，
 * 便于在**不启动 DSH** 的情况下用真实文档做回归（scripts/selftest）。
 */
export { EXTENSIONS, resolveEngine, renderByEngine, sniffContainer };

/**
 * 解析标签页的资源地址。
 *
 * @param address - `props.resourceAddress`。
 * @returns `{ sessionId, path }`；非 session 作用域返回 undefined。
 */
function parseSessionFileAddress(address) {
  try {
    const text = String(address ?? "");
    if (!text.startsWith(FILE_ADDRESS_PREFIX)) return undefined;
    const end = text.search(/[?#]/);
    const [scope, ...rest] = text
      .slice(FILE_ADDRESS_PREFIX.length, end === -1 ? undefined : end)
      .split("/");
    if (scope !== "session") return undefined;
    const [id, ...segments] = rest;
    if (id === undefined || id === "" || segments.length === 0) return undefined;
    return { sessionId: decodeURIComponent(id), path: segments.map(decodeURIComponent).join("/") };
  } catch {
    return undefined;
  }
}

/** 从路径取小写扩展名。 */
function extensionOf(path) {
  const match = /\.([^.\\/]+)$/.exec(String(path ?? ""));
  return match ? match[1].toLowerCase() : "";
}

/**
 * Office 正文组件。
 *
 * 渲染流程：解析地址 → 取整份字节 → 按签名分派引擎 → 渲染到容器。
 * 状态机：loading → ready / failed。
 */
function OfficeBody(props) {
  const { resourceAddress, content, statFile, activeFile } = props;
  const [state, setState] = useState({ kind: "loading" });
  const hostRef = useRef(null);

  useEffect(() => {
    let alive = true;
    const host = hostRef.current;

    const run = async () => {
      const parsed = parseSessionFileAddress(resourceAddress);
      const name = activeFile?.name ?? parsed?.path ?? "文件";
      if (parsed === undefined) {
        setState({ kind: "failed", message: `无法解析文件地址：${String(resourceAddress)}` });
        return;
      }
      if (content === undefined || content === null || content.kind !== "bytes") {
        setState({ kind: "failed", message: "宿主未提供文件字节（本渲染器需要 bytes-complete 模式）" });
        return;
      }

      const bytes = content.data;
      const ext = extensionOf(parsed.path);
      const engine = resolveEngine(ext, bytes);
      if (engine === "unsupported") {
        const container = sniffContainer(bytes);
        setState({
          kind: "failed",
          message: `暂不支持该格式（扩展名 .${ext || "?"}，容器 ${container}）`,
        });
        return;
      }

      try {
        setState({ kind: "loading", message: `正在解析 ${name}…` });
        await renderByEngine(engine, bytes, host, { name });
        if (alive) setState({ kind: "ready", engine });
      } catch (error) {
        if (alive) {
          setState({
            kind: "failed",
            message: `渲染失败（${engine}）：${String(error?.message ?? error)}`,
          });
        }
      }
    };

    void run();
    return () => {
      alive = false;
      if (host) while (host.firstChild) host.removeChild(host.firstChild);
    };
  }, [resourceAddress, content, activeFile, statFile]);

  if (state.kind === "failed") {
    return createElement(
      "div",
      { style: { padding: "24px", color: "#b3261e", fontSize: "13px", lineHeight: 1.7 } },
      createElement("strong", null, "无法预览此文档"),
      createElement("div", { style: { marginTop: "8px", color: "#5f6368" } }, state.message)
    );
  }

  return createElement(
    "div",
    { style: { position: "relative", width: "100%", height: "100%", overflow: "hidden" } },
    state.kind === "loading"
      ? createElement(
          "div",
          {
            style: {
              position: "absolute",
              inset: 0,
              display: "flex",
              alignItems: "center",
              justifyContent: "center",
              color: "#5f6368",
              fontSize: "13px",
            },
          },
          state.message
        )
      : null,
    createElement("div", {
      ref: hostRef,
      style: { width: "100%", height: "100%", overflow: "auto" },
    })
  );
}

/** 浏览器半插件体。 */
export function apply(ctx) {
  ctx.effect(() =>
    ctx.documentPreviews.register({
      id: BODY_ID,
      extensions: EXTENSIONS,
      priority: "extension",
      title: () => "办公文件预览",
      loading: "bytes-complete",
      wrap: false,
    })
  );

  ctx.effect(() =>
    ctx.slots.inject("sidebar.right.tab.document", () =>
      ctx.slots.register(
        {
          name: "sidebar.right.tab.document",
          key: BODY_ID,
          inject: () => ({
            statFile: (sessionId, path, signal) =>
              ctx.remote.workspaceFiles.stat(sessionId, path, signal),
          }),
        },
        OfficeBody
      )
    )
  );
}
