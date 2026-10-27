# -*- coding: utf-8 -*-
"""对“矛盾”ASIN 做第三次判定（带 cache-buster，独立会话），看哪个状态可复现。
用法: python tri.py B0CPSJT5SN,B0FH53CKPW --out tri.json [--rounds 2]
"""
import base64
import json
import os
import random
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

JS = r"""
(() => {
  const html = document.documentElement.innerHTML;
  const tf = document.querySelector('#twister_feature_div');
  const dv = /"dimensionValuesDisplayData"\s*:\s*(\{[\s\S]{0,60000}?\})\s*[,}]/.exec(html);
  let n = 0, keys = [];
  if (dv) { try { const o = JSON.parse(dv[1]); keys = Object.keys(o); n = keys.length; } catch(e) {} }
  const dims = {};
  document.querySelectorAll('script[data-a-state]').forEach(s => {
    if (!/twister/i.test(s.getAttribute('data-a-state')||'')) return;
    let d=null; try{ d=JSON.parse(s.textContent); }catch(e){ return; }
    const all=(d||{}).sortedDimValuesForAllDims; if(!all) return;
    Object.keys(all).forEach(k => (all[k]||[]).forEach(v => {
      (dims[k]=dims[k]||new Set()).add((v.defaultAsin||'').trim()); }));
  });
  const st = new Set(); Object.values(dims).forEach(s => s.forEach(a => { if (a) st.add(a); }));
  const others = n ? keys.filter(k => /^[A-Z0-9]{10}$/.test(k)) : [];
  return { ready: document.readyState, tfLen: tf?tf.innerHTML.length:0,
    dvddKeys: n, dvddAsins: others.slice(0,80), nState: st.size,
    stateAsins: [...st].slice(0,80),
    url: location.href };
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


def jload(o):
    i = o.find("{")
    while i != -1:
        try:
            return json.loads(o[i:])
        except json.JSONDecodeError:
            for j in range(len(o), i, -1):
                try:
                    return json.loads(o[i:j])
                except json.JSONDecodeError:
                    continue
        i = o.find("{", i + 1)
    return None


def ev(code):
    b64 = base64.b64encode(code.encode()).decode()
    rc, out, err = run([CLI, "--json", "eval", "--base64", b64], timeout=150)
    o = jload(out)
    if o is None:
        return None
    cur = o
    for _ in range(6):
        if isinstance(cur, dict) and ("result" in cur or "data" in cur):
            cur = cur.get("result") if "result" in cur else cur.get("data")
        else:
            break
    return cur


asins = [x.strip() for x in sys.argv[1].replace("\n", ",").split(",") if x.strip()]
outp = Path(sys.argv[2]) if len(sys.argv) > 2 else Path("tri.json")
rounds = int(sys.argv[4]) if len(sys.argv) > 4 else 2
res = {}
for rd in range(rounds):
    subprocess.run([CLI, "--json", "close", "--all"], stdout=subprocess.DEVNULL,
                   stderr=subprocess.DEVNULL, timeout=90, env=ENV)
    time.sleep(2)
    subprocess.run([CLI, "--json", "open", "about:blank"], stdout=subprocess.DEVNULL,
                   stderr=subprocess.DEVNULL, timeout=240, env=ENV)
    time.sleep(2)
    for asin in asins:
        url = "https://www.amazon.com/dp/%s?th=1&psc=1&_=%d" % (asin, random.randint(10**8, 10**9))
        run([CLI, "--json", "open", url], timeout=240)
        time.sleep(5)
        d = None
        for _ in range(6):
            d = ev(JS)
            if d and d.get("ready") == "complete" and (d.get("dvddKeys") or d.get("nState")):
                break
            time.sleep(3)
        res.setdefault(asin, []).append(d)
        print("r%d %s dvdd=%s state=%s" % (rd + 1, asin, (d or {}).get("dvddKeys"), (d or {}).get("nState")), flush=True)
outp.write_text(json.dumps(res, ensure_ascii=False, indent=1), encoding="utf-8")
