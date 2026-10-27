# -*- coding: utf-8 -*-
"""对首轮 nVar<2 的 ASIN 做稳健复抓：等 twister 数据稳定，再用 DOM 选项交叉校验。
结论口径：
  捆绑 = 前台变体选择器里可切换的不同 ASIN >= 2
用法: python final_check.py --front full_pd.json --asins-file asins.txt --out final_pd.json
"""
import argparse
import base64
import json
import os
import random
import subprocess
import tempfile
import time
from pathlib import Path

CLI = os.environ.get("AGENT_BROWSER_CLI") or str(
    Path(os.environ.get("LOCALAPPDATA", str(Path.home() / "AppData/Local"))) /
    "hermes/hermes-agent/node_modules/agent-browser/bin/agent-browser-win32-x64.exe")
ENV = dict(os.environ)
ENV.setdefault("AGENT_BROWSER_IDLE_TIMEOUT_MS", "3600000")

READY_JS = r"""(() => { const b=document.body?document.body.innerText:'';
 const tf=document.querySelector('#twister_feature_div');
 return {len:b.length, continuePage:/continue shopping/i.test(b)&&b.length<1500,
  hasTitle:!!document.querySelector('#productTitle'),
  twLen: tf?tf.innerHTML.length:0,
  hasState: /desktop-twister-sort-filter-data/.test(document.documentElement.innerHTML),
  captcha:/not a robot|Enter the characters/i.test(b)};})()"""

CONTINUE_JS = r"""(() => { const txt=e=>((e.innerText||e.value||e.getAttribute('aria-label')||'')+'').trim();
 const c=[...document.querySelectorAll('button,input[type=submit],input[type=button],a,span.a-button-text')];
 const t=c.find(e=>/continue shopping/i.test(txt(e))); if(t){t.click();return {clicked:1};}
 const f=document.querySelector('form'); if(f){const s=f.querySelector('input[type=submit],button');
 if(s){s.click();return {clicked:2};} f.submit(); return {clicked:3};} return {clicked:0};})()"""

EX_JS = r"""
(() => {
  const clean = s => (s||'').replace(/\s+/g,' ').trim();
  const A = ((document.querySelector('#ASIN')||{}).value)||'';
  const landed = (location.href.match(/\/dp\/([A-Z0-9]{10})/)||[])[1]||'';
  const dims = {}; let hasState=false;
  document.querySelectorAll('script[data-a-state]').forEach(s => {
    const raw = s.getAttribute('data-a-state')||'';
    if (!/twister/i.test(raw)) return;
    let d=null; try{ d=JSON.parse(s.textContent); }catch(e){ return; }
    const all=(d||{}).sortedDimValuesForAllDims; if (!all) return; hasState=true;
    Object.keys(all).forEach(dim => (all[dim]||[]).forEach(v => {
      (dims[dim]=dims[dim]||[]).push({asin:(v.defaultAsin||'').trim(),
        text:v.dimensionValueDisplayText||'', state:v.dimensionValueState||''});
    }));
  });
  const set = new Set(); const vals=[];
  Object.values(dims).forEach(arr=>arr.forEach(v=>{ if(v.asin) set.add(v.asin);
    if(v.text && !vals.includes(v.text)) vals.push(v.text); }));
  const tw = document.querySelector('#twister, #twister_feature_div');
  const domAsins = new Set();
  if (tw) tw.querySelectorAll('[data-defaultasin],[data-asin],option').forEach(e=>{
    const a=(e.getAttribute('data-defaultasin')||e.getAttribute('data-asin')||e.value||'').trim();
    if (/^[A-Z0-9]{10}$/.test(a)) domAsins.add(a); });
  const rating = clean(((document.querySelector('#acrPopover')||{}).getAttribute
      ? (document.querySelector('#acrPopover')||{}).getAttribute('title')||'' : ''));
  const ratingCount = clean(((document.querySelector('#acrCustomerReviewText')||{}).innerText)||'');
  const title = clean(((document.querySelector('#productTitle')||{}).innerText)||'');
  const brand = clean(((document.querySelector('#bylineInfo')||{}).innerText)||'');
  return { asin:A, landedAsin:landed, hasState:hasState,
    nStateAsins:set.size, stateAsins:[...set].slice(0,60),
    domAsins:[...domAsins].slice(0,60), nDomAsins:domAsins.size,
    dimNames:Object.keys(dims), dimN:Object.keys(dims).map(k=>k+':'+dims[k].length),
    vals:vals.slice(0,20), rating:rating, ratingCount:ratingCount,
    title:title.slice(0,130), brand:brand.slice(0,60),
    twText: tw?clean(tw.innerText).slice(0,160):'' };
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


def main():
    ap = argparse.ArgumentParser()
    ap.add_argument("--front", required=True)
    ap.add_argument("--asins-file", required=True)
    ap.add_argument("--out", required=True)
    a = ap.parse_args()
    base = json.loads(Path(a.front).read_text(encoding="utf-8"))
    asins = [x.strip() for x in Path(a.asins_file).read_text(encoding="utf-8").replace("\n", ",").split(",") if x.strip()]
    todo = [x for x in asins if x in base["rows"] and int(base["rows"][x].get("nVariantAsins") or 0) < 2]
    print("复核 nVar<2:", len(todo), flush=True)
    out = {"rows": {}, "failed": {}}
    if Path(a.out).exists():
        try:
            out = json.loads(Path(a.out).read_text(encoding="utf-8"))
        except Exception:
            pass
    subprocess.run([CLI, "--json", "open", "about:blank"], stdout=subprocess.DEVNULL,
                   stderr=subprocess.DEVNULL, timeout=240, env=ENV)
    time.sleep(2)
    for i, asin in enumerate(todo, 1):
        if asin in out["rows"]:
            print("[%d/%d] %s skip" % (i, len(todo), asin), flush=True); continue
        run([CLI, "--json", "open", "https://www.amazon.com/dp/%s?th=1&psc=1" % asin], timeout=240)
        time.sleep(3 + random.random() * 2)
        d, err, ok = None, None, False
        for attempt in range(6):
            st, se = ev(READY_JS)
            if st is None:
                time.sleep(3); continue
            if st.get("captcha"):
                err = "captcha"; break
            if st.get("continuePage"):
                ev(CONTINUE_JS); time.sleep(4 + random.random() * 3); continue
            if st.get("hasTitle") and st.get("hasState"):
                ok = True; break
            if st.get("hasTitle") and attempt >= 3 and (st.get("twLen") or 0) < 200:
                ok = True; break   # 无 twister 容器 → 真·无变体选择器
            time.sleep(3)
        if ok:
            d, err = ev(EX_JS)
            if d:
                # 稳定性：再读一次，取较大的 ASIN 集合
                time.sleep(2.5)
                d2, _ = ev(EX_JS)
                if d2 and (d2.get("nStateAsins") or 0) > (d.get("nStateAsins") or 0):
                    d = d2
        if d and (d.get("asin") or d.get("landedAsin") or (d.get("nStateAsins") or 0) > 0):
            d.setdefault("asin", asin)
            out["rows"][asin] = d
            print("[%d/%d] %s state=%s dom=%s dims=%s" % (i, len(todo), asin, d.get("nStateAsins"),
                                                          d.get("nDomAsins"), d.get("dimN")), flush=True)
        else:
            out["failed"][asin] = err or "no title/twister state"
            print("[%d/%d] %s FAIL %s" % (i, len(todo), asin, err), flush=True)
            time.sleep(15 + random.random() * 10)
        Path(a.out).write_text(json.dumps(out, ensure_ascii=False, indent=1), encoding="utf-8")
    print("done %d / failed %d" % (len(out["rows"]), len(out["failed"])), flush=True)


if __name__ == "__main__":
    main()
