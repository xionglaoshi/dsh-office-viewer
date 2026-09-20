/**
 * 离线自测：在**不启动 DSH** 的前提下，用真实文档验证渲染管线。
 *
 * 作用：CI / 新环境排查"到底是插件坏了还是 DSH 集成坏了"。
 * 做法：起一个极简 HTTP 服务（模拟 DSH 的 `/office-viewer-assets` 路由），
 *       用 Playwright 打开测试页，逐格式断言渲染结果。
 *
 * 用法：
 *   node scripts/selftest.mjs                      # 用内置样本
 *   node scripts/selftest.mjs ~/某目录             # 用指定目录里的真实文档
 */

import { createServer } from "node:http";
import { readFileSync, existsSync, statSync, readdirSync } from "node:fs";
import { join, extname } from "node:path";
import { fileURLToPath } from "node:url";

const PKG = fileURLToPath(new URL("..", import.meta.url));
const PORT = 8912;

const MIME = {
  ".html": "text/html; charset=utf-8",
  ".js": "text/javascript; charset=utf-8",
  ".mjs": "text/javascript; charset=utf-8",
  ".wasm": "application/wasm",
  ".otf": "font/otf",
  ".json": "application/json",
  ".png": "image/png",
};

/** 把待测文件读成 { 扩展名: Buffer }。 */
function collectSamples(dir) {
  if (!dir || !existsSync(dir) || !statSync(dir).isDirectory()) return {};
  const wanted = ["docx", "xlsx", "pptx", "doc", "xls", "ppt", "csv", "rtf", "ofd", "wps", "et", "dps"];
  const out = {};
  for (const name of readdirSync(dir)) {
    const ext = extname(name).slice(1).toLowerCase();
    if (wanted.includes(ext) && !(ext in out)) out[ext] = readFileSync(join(dir, name));
  }
  return out;
}

const samples = collectSamples(process.argv[2]);
const available = Object.keys(samples);
if (available.length === 0) {
  console.error("没有找到可测样本。请传入含 Office 文件的目录，例如：node scripts/selftest.mjs ~/wps");
  process.exit(2);
}
console.log(`样本：${available.join(", ")}`);

const client = readFileSync(join(PKG, "lib", "client.js"));
const assetsDir = join(PKG, "assets");

const page = `<!DOCTYPE html><html><head><meta charset="utf-8"><title>selftest</title>
<style>#host{width:1200px;height:800px;overflow:auto}</style></head><body>
<div id="host"></div>
<script src="/vendor/react.js"></script>
<script>
const factories=new Map();
window.__ModuleLoader__={load({id,factory}){factories.set(id,factory);}};
</script>
<script src="/client.js"></script>
<script>
window.__READY__=false;
try{
  const f=factories.get('dsh-office-viewer');
  window.__MOD__=f((n)=>{ if(n==='react') return window.React; throw new Error('unexpected require: '+n); });
  window.__READY__=true;
}catch(e){ window.__INIT_ERR__=String(e&&e.stack||e); }
window.__run=async(ext)=>{
  const host=document.getElementById('host');
  host.innerHTML='';
  try{
    const buf=new Uint8Array(await (await fetch('/sample/'+ext)).arrayBuffer());
    const m=window.__MOD__;
    const engine=m.resolveEngine(ext,buf);
    if(engine==='unsupported') return {ok:false,engine,err:'unsupported'};
    const t0=performance.now();
    await m.renderByEngine(engine,buf,host,{name:'sample.'+ext});
    // pptx 引擎的幻灯片是**异步**绘制到 Worker 后再回填 DOM 的：
    // open() resolve 时容器可能还是空的，因此这里轮询等待内容出现。
    for(let i=0;i<40;i++){
      const n=host.querySelectorAll('canvas,img,svg').length;
      const t=(host.innerText||'').trim().length;
      if(n>0||t>0) break;
      await new Promise(r=>setTimeout(r,250));
    }
    return {ok:true,engine,ms:Math.round(performance.now()-t0),
      textLen:(host.innerText||'').trim().length,
      canvas:host.querySelectorAll('canvas').length,
      img:host.querySelectorAll('img').length};
  }catch(e){ return {ok:false,err:String(e&&e.message||e)}; }
};
</script></body></html>`;

const server = createServer((req, res) => {
  const url = new URL(req.url, "http://localhost");
  const p = url.pathname;
  const send = (body, type) => {
    res.writeHead(200, { "content-type": type });
    res.end(body);
  };
  if (p === "/") return send(page, MIME[".html"]);
  if (p === "/client.js") return send(client, MIME[".js"]);
  if (p === "/vendor/react.js") {
    // react 由调用方保证存在于 node_modules（本插件不打包 react）
    const react = join(PKG, "node_modules", "react", "umd", "react.production.min.js");
    if (!existsSync(react)) {
      res.writeHead(500);
      return res.end("缺少 react，请先 npm install");
    }
    return send(readFileSync(react), MIME[".js"]);
  }
  if (p.startsWith("/sample/")) {
    const ext = p.slice(8);
    if (!(ext in samples)) {
      res.writeHead(404);
      return res.end();
    }
    return send(samples[ext], "application/octet-stream");
  }
  if (p.startsWith("/office-viewer-assets/")) {
    const rel = decodeURIComponent(p.slice("/office-viewer-assets/".length));
    const file = join(assetsDir, rel);
    if (!file.startsWith(assetsDir) || !existsSync(file)) {
      res.writeHead(404);
      return res.end();
    }
    return send(readFileSync(file), MIME[extname(file)] ?? "application/octet-stream");
  }
  res.writeHead(404);
  res.end();
});

server.listen(PORT, "127.0.0.1", async () => {
  let chromium;
  try {
    ({ chromium } = await import("playwright"));
  } catch {
    console.log(`\n未安装 playwright —— 跳过浏览器断言。\n可手动打开 http://127.0.0.1:${PORT}/ 检查。`);
    return;
  }
  const browser = await chromium.launch({
    // 允许复用共享 Chromium（避免重复下载）；未设置时用 playwright 自带的
    executablePath: process.env.CHROMIUM_PATH || undefined,
  });
  const pg = await browser.newPage();
  const errs = [];
  pg.on("pageerror", (e) => errs.push(String(e).slice(0, 200)));
  await pg.goto(`http://127.0.0.1:${PORT}/`);

  const initErr = await pg.evaluate("window.__INIT_ERR__");
  if (initErr) {
    console.error("✘ bundle 初始化失败：", initErr.slice(0, 400));
    await browser.close();
    server.close();
    process.exit(1);
  }

  let fail = 0;
  for (const ext of available) {
    const r = await pg.evaluate(`window.__run(${JSON.stringify(ext)})`);
    const good = r.ok && (r.textLen > 0 || r.canvas > 0 || r.img > 0);
    if (!good) fail += 1;
    console.log(
      `${good ? "✔" : "✘"} ${ext.padEnd(6)} engine=${String(r.engine).padEnd(10)} ` +
        `${String(r.ms ?? "-").padStart(6)}ms  text=${String(r.textLen ?? "-").padStart(6)}  ` +
        `canvas=${r.canvas ?? "-"} ${r.err ? "ERR:" + String(r.err).slice(0, 120) : ""}`
    );
  }
  await browser.close();
  server.close();
  if (errs.length) console.log("页面错误：", errs.slice(0, 5));
  console.log(fail === 0 ? "\n全部通过 ✅" : `\n${fail} 项失败 ❌`);
  process.exit(fail === 0 ? 0 : 1);
});
