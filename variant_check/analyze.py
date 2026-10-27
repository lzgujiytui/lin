# -*- coding: utf-8 -*-
"""车灯总成 变体结构分析：家族×车型拆解、评分/评论分布、单ASIN对照。"""
import json
import re
from collections import Counter, defaultdict
from pathlib import Path

import pandas as pd

W = Path(r"C:\Users\HP\Desktop\周环比+周报\发布\variant_check")
XLSX = Path(r"C:\Users\HP\Desktop\ASIN前台变体状态_重抓_20261022.xlsx")
rp = json.loads((W / "repull_all.json").read_text(encoding="utf-8"))["rows"]

xl = pd.ExcelFile(XLSX)
inb = pd.read_excel(xl, "在变体组内")
out = pd.read_excel(xl, "未捆绑")
fam = pd.read_excel(xl, "变体家族")

# 车型抽取：变体值/标题里 "for XXX" 片段
YR = re.compile(r"\b(19|20)\d{2}\b")
def models_of(texts):
    got = set()
    for t in texts:
        for m in re.finditer(r"(?:for|with|fits?|compatible with|replacement for)\s+([^,;/]{3,45})", str(t), re.I):
            s = " " + YR.sub(" ", m.group(1)).lower()
            s = re.sub(r"[^a-z0-9 &-]", " ", s)
            s = " ".join(w for w in s.split() if len(w) > 1)[:38].strip()
            if s:
                got.add(s)
    return got

rows = []
for _, fr in fam.iterrows():
    ours = str(fr["清单内我方ASIN"]).split("/")
    r = rp.get(ours[0], {})
    vals = (r.get("vals") or [])
    texts = vals + [rp.get(a, {}).get("title", "") for a in ours]
    ms = models_of(texts)
    rat, cnts = [], []
    for a in ours:
        d = rp.get(a, {})
        m = re.search(r"([\d.]+)", str(d.get("rating") or ""))
        c = re.search(r"([\d,]+)", str(d.get("ratingCount") or ""))
        if m:
            rat.append(float(m.group(1)))
        if c:
            cnts.append(int(c.group(1).replace(",", "")))
    rows.append({"家族": int(fr["家族编号"]), "选项数": int(fr["前台家族ASIN数"]),
                 "我方条数": len(ours), "识别车型数": len(ms),
                 "车型示例": " / ".join(sorted(ms)[:4]), "变体值数": len(vals),
                 "星级中位": round(pd.Series(rat).median(), 2) if rat else None,
                 "评论合计": sum(cnts) if cnts else None,
                 "有评分条数": len(rat)})
F = pd.DataFrame(rows)
print("=== 家族结构（按选项数降序）")
print(F.to_string(index=False))

def stat(df, name):
    rat = pd.to_numeric(df["星级中位"], errors="coerce")
    print("\n%s: n=%d, 有评分 %d, 星级中位 %.2f, 评论合计 %s" % (
        name, len(df), df["有评分条数"].sum() if "有评分条数" in df else 0,
        rat.median() if rat.notna().any() else 0, df["评论合计"].sum() if "评论合计" in df else ""))


# 单ASIN对照
srows = []
for _, r in out.iterrows():
    d = rp.get(r["ASIN"], {})
    m = re.search(r"([\d.]+)", str(d.get("rating") or ""))
    c = re.search(r"([\d,]+)", str(d.get("ratingCount") or ""))
    srows.append({"星级": float(m.group(1)) if m else None,
                  "评论": int(c.group(1).replace(",", "")) if c else None})
S = pd.DataFrame(srows)
print("\n=== 未捆绑 33 条")
print("有评分 %d 条 | 星级中位 %.2f | 评论中位 %s | 评论合计 %d" % (
    S["星级"].notna().sum(), S["星级"].median(), S["评论"].median(), S["评论"].fillna(0).sum()))

print("\n=== 大组 vs 小组 vs 单ASIN")
big = F[F["选项数"] >= 12]
mid = F[(F["选项数"] >= 2) & (F["选项数"] < 12)]
print("大组(%d个, 我方%d条): 星级中位 %s, 评论合计 %s" % (
    len(big), big["我方条数"].sum(), big["星级中位"].median(), big["评论合计"].sum()))
print("小组(%d个, 我方%d条): 星级中位 %s, 评论合计 %s" % (
    len(mid), mid["我方条数"].sum(), mid["星级中位"].median(), mid["评论合计"].sum()))

# 同车型内是否真有颜色变化（判断合法变体）
print("\n=== 抽查：家族内变体值是否同一车型多颜色")
for fid in [3, 4, 5, 6, 7, 8, 9]:
    r = fam[fam["家族编号"] == fid]
    if not len(r):
        continue
    ours = str(r.iloc[0]["清单内我方ASIN"]).split("/")
    vals = (rp.get(ours[0], {}).get("vals") or [])[:8]
    print("家族%d (%d选项/%d条): %s" % (fid, int(r.iloc[0]["前台家族ASIN数"]), len(ours), " | ".join(vals)))
F.to_csv(W / "family_analysis.csv", index=False, encoding="utf-8-sig")
print("\nsaved family_analysis.csv")
