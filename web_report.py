# -*- coding: utf-8 -*-
"""
执法处周报材料生成器 —— 网页版（本地 Flask 应用）
================================================

与 weekly_report.py（桌面GUI版）共用同一套计算逻辑，输入仅为：
检查原始数据导出 + 处罚原始数据导出（+ 三个固定文件）。

运行：双击 start_web.bat，或用 Python312 运行本脚本后访问 http://127.0.0.1:8765/
"""

import os
import traceback
import uuid

from flask import Flask, jsonify, request, send_file

import weekly_report as wr

app = Flask(__name__)
GENERATED = {}   # token -> 生成的 docx 路径（仅本次运行内有效）

PAGE = r"""<!doctype html>
<html lang="zh-CN">
<head>
<meta charset="utf-8">
<title>执法处周报材料生成器</title>
<meta name="viewport" content="width=device-width, initial-scale=1">
<style>
  * { box-sizing: border-box; }
  body { margin: 0; padding: 24px; background: #f2f4f8; color: #222;
         font-family: "Microsoft YaHei", "微软雅黑", sans-serif; }
  .wrap { max-width: 980px; margin: 0 auto; }
  h1 { font-size: 22px; margin: 0 0 16px; }
  .card { background: #fff; border-radius: 8px; padding: 16px 20px;
          margin-bottom: 16px; box-shadow: 0 1px 3px rgba(0,0,0,.08); }
  .card h2 { font-size: 15px; margin: 0 0 12px; color: #555;
             border-left: 4px solid #2f6fdb; padding-left: 8px; }
  .row { display: flex; align-items: center; margin-bottom: 8px; }
  .row label { width: 210px; text-align: right; padding-right: 12px;
               color: #333; font-size: 14px; flex-shrink: 0; }
  .row input { flex: 1; padding: 7px 10px; border: 1px solid #ccc;
               border-radius: 4px; font-size: 14px; min-width: 0; }
  .row input:focus { outline: none; border-color: #2f6fdb; }
  .btns { margin: 4px 0 0 222px; }
  button { padding: 8px 22px; font-size: 14px; border: none; border-radius: 4px;
           cursor: pointer; margin-right: 10px; }
  #btn-auto { background: #eef2fb; color: #2f6fdb; border: 1px solid #2f6fdb; }
  #btn-gen  { background: #2f6fdb; color: #fff; }
  button:disabled { opacity: .55; cursor: wait; }
  pre { background: #1e1e1e; color: #d6e2c8; padding: 12px; border-radius: 6px;
        font-size: 13px; line-height: 1.55; white-space: pre-wrap;
        word-break: break-all; max-height: 380px; overflow: auto; margin: 0; }
  .para { background: #fbfbf6; border: 1px solid #e5e1cf; border-radius: 6px;
          padding: 12px 14px; font-size: 15px; line-height: 1.9;
          margin-bottom: 10px; text-indent: 2em; }
  #dl { display: none; margin: 4px 0 0 222px; }
  #dl a { display: inline-block; background: #2e9e5b; color: #fff;
          padding: 9px 26px; border-radius: 4px; text-decoration: none;
          font-size: 15px; }
  .hint { font-size: 12px; color: #999; margin: 6px 0 0 222px; }
</style>
</head>
<body>
<div class="wrap">
  <h1>执法处周报材料生成器 <span style="font-size:13px;color:#999">网页版</span></h1>

  <div class="card">
    <h2>每周数据（平台原始导出）</h2>
    <div class="row"><label>检查原始数据导出：</label><input id="check_raw" spellcheck="false"></div>
    <div class="row"><label>处罚原始数据导出：</label><input id="punish_raw" spellcheck="false"></div>
  </div>

  <div class="card">
    <h2>参数</h2>
    <div class="row"><label>本周开始日期：</label><input id="d0" type="date"></div>
    <div class="row"><label>本周结束日期：</label><input id="d1" type="date"></div>
    <div class="row"><label>上周处罚件数（可留空）：</label><input id="lw_punish" type="number" min="0"></div>
    <div class="row"><label>输出 docx：</label><input id="out" spellcheck="false"></div>
    <div class="btns">
      <button id="btn-auto">自动探测文件</button>
      <button id="btn-gen">生成周报</button>
    </div>
    <div class="hint">生成约需半分钟（首次读取大文件较慢）；上周数值优先采用历史记录，与上周报告保持一致。</div>
    <div id="dl"><a id="dl-a" href="#">下载生成的 docx</a></div>
  </div>

  <div class="card" id="result-card" style="display:none">
    <h2>生成结果</h2>
    <div class="para" id="para1"></div>
    <div class="para" id="para2"></div>
  </div>

  <div class="card">
    <h2>运行日志</h2>
    <pre id="log">就绪。</pre>
  </div>
</div>

<script>
const FIELDS = ["check_raw","punish_raw","d0","d1","lw_punish","out"];
const $ = id => document.getElementById(id);
const logEl = $("log");
function setLog(t){ logEl.textContent = t; logEl.scrollTop = logEl.scrollHeight; }

async function post(url, body){
  const r = await fetch(url, {method:"POST",
    headers:{"Content-Type":"application/json"}, body: JSON.stringify(body||{})});
  return await r.json();
}

$("btn-auto").onclick = async () => {
  $("btn-auto").disabled = true;
  setLog("正在自动探测文件与日期……");
  try {
    const res = await post("/api/autodetect");
    let n = 0;
    for (const k of FIELDS) if (res[k]) { $(k).value = res[k]; n++; }
    setLog("自动探测完成，已填充 " + n + " 项。请核对后点击“生成周报”。");
  } catch(e) { setLog("自动探测失败：" + e); }
  $("btn-auto").disabled = false;
};

$("btn-gen").onclick = async () => {
  $("btn-gen").disabled = true;
  $("btn-gen").textContent = "生成中，请稍候……";
  $("dl").style.display = "none";
  $("result-card").style.display = "none";
  setLog("正在生成，请稍候……\n");
  const params = {};
  for (const k of FIELDS) params[k] = $(k).value;
  try {
    const res = await post("/api/generate", params);
    setLog(res.log + (res.ok ? "\n已生成: " + res.out
                              : "\n生成失败：\n" + (res.error||"")));
    if (res.ok) {
      $("para1").textContent = res.para1;
      $("para2").textContent = res.para2;
      $("result-card").style.display = "";
      $("dl-a").href = "/api/download/" + res.token;
      $("dl").style.display = "";
    }
  } catch(e) { setLog("请求失败：" + e); }
  $("btn-gen").disabled = false;
  $("btn-gen").textContent = "生成周报";
};

window.addEventListener("load", () => $("btn-auto").click());
</script>
</body>
</html>
"""


@app.route("/")
def index():
    return PAGE


@app.route("/api/autodetect", methods=["POST"])
def api_autodetect():
    return jsonify(wr.autodetect_params())


@app.route("/api/generate", methods=["POST"])
def api_generate():
    params = request.get_json(force=True, silent=True) or {}
    logs = []
    try:
        para1, para2, out_path = wr.run_report(params, logs.append)
    except Exception:
        return jsonify(ok=False, log="\n".join(logs),
                       error=traceback.format_exc())
    token = uuid.uuid4().hex
    GENERATED[token] = out_path
    if len(GENERATED) > 32:
        GENERATED.pop(next(iter(GENERATED)))
    return jsonify(ok=True, log="\n".join(logs), para1=para1, para2=para2,
                   out=out_path, token=token)


@app.route("/api/download/<token>")
def api_download(token):
    path = GENERATED.get(token)
    if not path or not os.path.exists(path):
        return "文件不存在或链接已过期，请重新生成。", 404
    return send_file(path, as_attachment=True)


if __name__ == "__main__":
    app.run(host="127.0.0.1", port=8765, debug=False, threaded=True)
