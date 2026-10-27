# -*- coding: utf-8 -*-
"""精细探测单个 ASIN 的变体选择器真实选项（多信号，落原始片段人工核对）。
用法: python probe.py B0HJZSBX7X,B0HJZYH5QG [--out probe.json]
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

STATE_JS = """(() => { const b=document.body?document.body.innerText:'';
 return {len:b.length, continuePage:/continue shopping/i.test(b)&&b.length<1500,
 hasTitle:!!document.querySelector('#productTitle'),
 twisterLen: (document.querySelector('#twister_feature_div')||{}).innerHTML ?
   document.querySelector('#twister_feature_div').innerHTML.length : 0,
 captcha:/not a robot|Enter the characters/i.test(b)};})()"""

CONTINUE_JS = r"""(() => { const txt=e=>((e.innerText||e.value||e.getAttribute('aria-label')||'')+'').trim();
 const c=[...document.querySelectorAll('button,input[type=submit],input[type=button],a,span.a-button-text')];
 const t=c.find(e=>/continue shopping/i.test(txt(e))); if(t){t.click();return {clicked:1};}
 const f=document.querySelector('form'); if(f){const s=f.querySelector('input[type=submit],button');
 if(s){s.click();return {clicked:2};} f.submit(); return {clicked:3};} return {clicked:0};})()"""

PROBE_JS = r"""
(() => {
  const clean = s => (s||'').replace(/\s+/g,' ').trim();
  const out = {asin:((document.querySelector('#ASIN')||{}).value)||'',
    landed:(location.href.match(/\/dp\/([A-Z0-9]{10})/)||[])[1]||'', scripts:[]};
  // 1) 所有 twister 相关 script 的原始状态（只留结构摘要）
  document.querySelectorAll('script[data-a-state]').forEach(s => {
    const key = (()=>{ try{ return JSON.parse(s.getAttribute('data-a-state')).key; }catch(e){ return s.getAttribute('data-a-state'); } })();
    if (!/twister|dimension/i.test(String(key))) return;
    let d=null; try{ d=JSON.parse(s.textContent); }catch(e){ return; }
    const rec={key:key, topKeys:Object.keys(d||{}).slice(0,20)};
    const all=(d||{}).sortedDimValuesForAllDims||{};
    rec.dims=Object.keys(all).map(k=>({dim:k, n:(all[k]||[]).length,
      items:(all[k]||[]).slice(0,40).map(v=>({v:v.dimensionValueDisplayText, a:v.defaultAsin, s:v.dimensionValueState}))}));
    if (d.asinsInCollapsedView) rec.collapsed=(d.asinsInCollapsedView||[]).slice(0,40);
    if (d.dimensionValuesDisplayData) rec.dvdd=Object.keys(d.dimensionValuesDisplayData).slice(0,20);
    out.scripts.push(rec);
  });
  // 2) DOM: 选择器里的可选 ASIN
  const tw = document.querySelector('#twister, #twister_feature_div');
  const domAsins = new Set(); const selInfo=[];
  if (tw) {
    tw.querySelectorAll('[data-defaultasin],[data-asin],option').forEach(e=>{
      const a=(e.getAttribute('data-defaultasin')||e.getAttribute('data-asin')||e.value||'').trim();
      if (/^[A-Z0-9]{10}$/.test(a)) domAsins.add(a);
    });
    tw.querySelectorAll('select').forEach(s=>selInfo.push({id:s.id, n:s.options.length,
      vals:[...s.options].slice(0,20).map(o=>o.value+'='+clean(o.text))}));
    out.twText = clean(tw.innerText).slice(0,300);
    out.twHtmlLen = tw.innerHTML.length;
    out.swatchCount = tw.querySelectorAll('li, .swatchUnavailableDim, [id^="color_name_"], [id^="variation_"]').length;
    out.dimLabels = [...new Set([...tw.querySelectorAll('label,.a-form-label,#variation_color_name, #inline-twister-row-color_name, #inline-twister-expander-content-color_name')].map(e=>clean(e.innerText)).filter(Boolean))].slice(0,10);
  }
  out.domAsins=[...domAsins].slice(0,60);
  out.selects=selInfo;
  // 3) window 上的老口径数据
  try { const jq = (typeof jQuery!=='undefined') ? jQuery('#twister').data() : null;
    if (jq) out.jqKeys = Object.keys(jq).slice(0,20); } catch(e){}
  return out;
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
        return None, "rc=%s %s" % (rc, (out + err)[:150])
    cur = obj
    for _ in range(6):
        if isinstance(cur, dict) and ("result" in cur or "data" in cur):
            cur = cur.get("result") if "result" in cur else cur.get("data")
        else:
            break
    return cur, None


def main():
    asins = [x.strip() for x in sys.argv[1].replace("\n", ",").split(",") if x.strip()]
    outp = Path(sys.argv[3]) if len(sys.argv) > 3 else Path("probe.json")
    res = {}
    subprocess.run([CLI, "--json", "open", "about:blank"], stdout=subprocess.DEVNULL,
                   stderr=subprocess.DEVNULL, timeout=240, env=ENV)
    time.sleep(2)
    for asin in asins:
        run([CLI, "--json", "open", "https://www.amazon.com/dp/%s?th=1&psc=1" % asin], timeout=240)
        for _ in range(4):
            st, _e = ev(STATE_JS) or (None, None)
            if st and st.get("continuePage"):
                ev(CONTINUE_JS); time.sleep(4)
            if st and st.get("hasTitle") and st.get("twisterLen", 0) > 200:
                break
            time.sleep(3)
        time.sleep(2)
        d, err = ev(PROBE_JS)
        res[asin] = d or {"error": err}
        print(asin, "ok" if d else "FAIL %s" % err, flush=True)
        outp.write_text(json.dumps(res, ensure_ascii=False, indent=1), encoding="utf-8")


if __name__ == "__main__":
    main()
