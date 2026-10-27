# -*- coding: utf-8 -*-
"""诊断：同一页面在“半加载 / 全加载”时各信号长什么样，用于确定可靠判定口径。
用法: python diag.py B0CPSJT5SN,B0HJZS4GY3 --out diag.json
每页读 4 次（间隔 3s），记录 tfLen / 状态script / dimensionValuesDisplayData / 可见文案。
"""
import base64
import json
import os
import subprocess
import sys
import tempfile
import time
from pathlib import Path

CLI = os.environ.get("AGENT_BROWSER_CLI") or str(
    Path(os.environ.get("LOCALAPPDATA", str(Path.home() / "AppData/Local"))) /
    "hermes/hermes-agent/node_modules/agent-browser/bin/agent-browser-win32-x64.exe")
ENV = dict(os.environ)
ENV.setdefault("AGENT_BROWSER_IDLE_TIMEOUT_MS", "3600000")

PROBE = r"""
(() => {
  const html = document.documentElement.innerHTML;
  const tf = document.querySelector('#twister_feature_div');
  const twister = document.querySelector('#twister');
  const body = document.body ? document.body.innerText : '';
  const keys = [];
  document.querySelectorAll('script[data-a-state]').forEach(s => {
    let k = ''; try { k = JSON.parse(s.getAttribute('data-a-state')).key; } catch(e) { k = s.getAttribute('data-a-state'); }
    if (/twister|dimension/i.test(String(k))) keys.push(String(k));
  });
  // dimensionValuesDisplayData 常在页面内联 JSON 里（{值: ASIN}），是服务端渲染的强信号
  const dvdd = /"dimensionValuesDisplayData"\s*:\s*(\{.{0,4000}?\})/.exec(html);
  let dvddN = 0, dvddRaw = '';
  if (dvdd) {
    dvddRaw = dvdd[1].slice(0, 300);
    dvddN = (dvdd[1].match(/[A-Z0-9]{10}/g) || []).length;
  }
  const asinVar = /asinVariationValues/.test(html);
  const swatches = twister ? twister.querySelectorAll('[id^="color_name_"], [id^="variation_"], li').length : 0;
  return {
    ready: document.readyState,
    tfLen: tf ? tf.innerHTML.length : 0,
    twisterKeys: keys,
    hasDvdd: !!dvdd, dvddN: dvddN, dvddRaw: dvddRaw,
    hasAsinVariation: asinVar,
    swatches: swatches,
    colorText: (/Color\s*:\s*[^\n]{0,60}/.exec(body) || [''])[0],
    dimWord: /(Style|Size|Color|Pattern|Set|Pack)\s*:/.test(body) ? 'yes' : 'no',
    htmlLen: html.length
  };
})()
"""


def run(args, timeout=240):
    with tempfile.TemporaryFile("w+", encoding="utf-8", errors="replace") as fo, \
            tempfile.TemporaryFile("w+", encoding="utf-8", errors="replace") as fe:
        try:
            p = subprocess.run(args, stdout=fo, stderr=fe, timeout=timeout, env=ENV,
                               encoding="utf-8", errors="replace")
            rc = p.returncode
        except subprocess.TimeoutExpired:
            rc = "TIMEOUT"
        fo.seek(0); fe.seek(0); return rc, fo.read(), fe.read()


def jload(out):
    i = out.find("{")
    while i != -1:
        try:
            return json.loads(out[i:])
        except json.JSONDecodeError:
            for j in range(len(out), i, -1):
                try:
                    return json.loads(out[i:j])
                except json.JSONDecodeError:
                    continue
        i = out.find("{", i + 1)
    return None


def ev(code, timeout=150):
    b64 = base64.b64encode(code.encode("utf-8")).decode("ascii")
    rc, out, err = run([CLI, "--json", "eval", "--base64", b64], timeout=timeout)
    obj = jload(out)
    if obj is None:
        return None, "rc=%s %s" % (rc, (out + err)[:120])
    cur = obj
    for _ in range(6):
        if isinstance(cur, dict) and ("result" in cur or "data" in cur):
            cur = cur.get("result") if "result" in cur else cur.get("data")
        else:
            break
    return cur, None


asins = [x.strip() for x in sys.argv[1].replace("\n", ",").split(",") if x.strip()]
outp = Path(sys.argv[2]) if len(sys.argv) > 2 else Path("diag.json")
res = {}
subprocess.run([CLI, "--json", "open", "about:blank"], stdout=subprocess.DEVNULL,
               stderr=subprocess.DEVNULL, timeout=240, env=ENV)
time.sleep(2)
for asin in asins:
    run([CLI, "--json", "open", "https://www.amazon.com/dp/%s?th=1&psc=1" % asin], timeout=240)
    time.sleep(4)
    seq = []
    for k in range(4):
        d, e = ev(PROBE)
        seq.append(d if d else {"error": e})
        time.sleep(3)
    res[asin] = seq
    print(asin, "collected", flush=True)
outp.write_text(json.dumps(res, ensure_ascii=False, indent=1), encoding="utf-8")
