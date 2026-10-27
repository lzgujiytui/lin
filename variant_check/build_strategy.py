# -*- coding: utf-8 -*-
"""车灯总成 变体结构诊断 + 建组方案交付表。"""
import json
import re
from collections import defaultdict
from datetime import date
from pathlib import Path

import pandas as pd
from openpyxl.styles import Alignment, Border, Font, PatternFill, Side
from openpyxl.utils import get_column_letter

W = Path(r"C:\Users\HP\Desktop\周环比+周报\发布\variant_check")
XLSX = Path(r"C:\Users\HP\Desktop\ASIN前台变体状态_重抓_20261022.xlsx")
OUT = Path(r"C:\Users\HP\Desktop\车灯总成_变体结构诊断与建组方案_20261022.xlsx")
TODAY = str(date.today())

rp = json.loads((W / "repull_all.json").read_text(encoding="utf-8"))["rows"]
xl = pd.ExcelFile(XLSX)
inb = pd.read_excel(xl, "在变体组内")
out = pd.read_excel(xl, "未捆绑")
fam = pd.read_excel(xl, "变体家族")

YR = re.compile(r"\b(19|20)\d{2}\b")


def norm_model(s):
    s = " " + YR.sub(" ", str(s).lower())
    s = re.sub(r"[^a-z0-9 &-]", " ", s)
    s = " ".join(w for w in s.split() if len(w) > 1)
    for junk in (" pair fr", " pair", " base", " ltz", " ls", " lt", " se", " gls", " limited"):
        s = s.replace(junk, " ")
    return " ".join(s.split())[:38]


def models_of(texts):
    got = set()
    for t in texts:
        for m in re.finditer(r"(?:for|with|fits?|compatible with|replacement for)\s+([^,;/]{3,45})", str(t), re.I):
            k = norm_model(m.group(1))
            if k:
                got.add(k)
    # 品牌前缀/后缀归一：camry ⊂ toyota camry，wrangler ⊂ jeep wrangler
    keep = []
    for k in sorted(got, key=len, reverse=True):
        if not any(k in x or x in k for x in keep):
            keep.append(k)
    return set(keep)


def rating(a):
    d = rp.get(a, {})
    m = re.search(r"([\d.]+)", str(d.get("rating") or ""))
    c = re.search(r"([\d,]+)", str(d.get("ratingCount") or ""))
    return (float(m.group(1)) if m else None,
            int(c.group(1).replace(",", "")) if c else None)


rows = []
for _, fr in fam.iterrows():
    ours = str(fr["清单内我方ASIN"]).split("/")
    vals = (rp.get(ours[0], {}).get("vals") or [])
    texts = vals + [rp.get(a, {}).get("title", "") for a in ours]
    ms = models_of(texts)
    rat, cnts = [], []
    for a in ours:
        r, c = rating(a)
        if r:
            rat.append(r)
        if c:
            cnts.append(c)
    opts = int(fr["前台家族ASIN数"])
    rows.append({"家族": int(fr["家族编号"]), "前台选项数": opts, "我方条数": len(ours),
                 "识别车型数": len(ms), "每车型平均选项数": round(opts / max(len(ms), 1), 2),
                 "星级中位": round(pd.Series(rat).median(), 2) if rat else None,
                 "评论合计": sum(cnts) if cnts else None,
                 "变体维度": fr["变体维度"],
                 "车型示例": " / ".join(sorted(ms)[:5]) or "（未识别）"})
F = pd.DataFrame(rows)
F["结构判定"] = F["识别车型数"].apply(lambda n: "跨车型乱挂" if n >= 2 else "同车型内变体（合法形状）")
F["优先级"] = F["前台选项数"].apply(lambda n: "高" if n >= 12 else ("中" if n >= 6 else "低"))
F["建议父体数(按车型)"] = F["识别车型数"]
F["动作"] = F.apply(lambda r: (
    "拆成 %d 个父体：同车型同款式的颜色/侧别留作子体，跨车型全部脱钩" % r["识别车型数"]
    if r["识别车型数"] >= 2 else "保留，补齐同款颜色子体"), axis=1)
F = F.sort_values(["优先级", "前台选项数"], ascending=[True, False]).reset_index(drop=True)
F["优先级"] = pd.Categorical(F["优先级"], ["高", "中", "低"], ordered=True)
F = F.sort_values(["优先级", "前台选项数"], ascending=[True, False]).reset_index(drop=True)

# 单 ASIN 归位建议
srows = []
for _, r in out.iterrows():
    d = rp.get(r["ASIN"], {})
    note = str(r["说明"])
    if "残留在他人变体组" in note:
        act = "拆干净：删掉指向他人的死维度，按自身车型独立建父体"
        pr = "高"
    elif "仅1个选项" in note:
        act = "确认是否与已有同车型父体同款：同款→并入作子体；不同款→删死维度独立在售"
        pr = "中"
    else:
        act = "无变体结构问题；若同车型同款已有父体，作为子体并入拿评论聚合，否则保持独立"
        pr = "低"
    rr, cc = rating(r["ASIN"])
    srows.append({"ASIN": r["ASIN"], "说明": note, "星级": rr, "评分人数": cc,
                  "优先级": pr, "建议动作": act, "标题": str(r["标题"])[:90]})
S = pd.DataFrame(srows)
S["优先级"] = pd.Categorical(S["优先级"], ["高", "中", "低"], ordered=True)
S = S.sort_values(["优先级", "ASIN"]).reset_index(drop=True)

# 规则表
rules = pd.DataFrame([
    ["合法变体维度", "颜色/外观：Chrome、Black Housing、Smoke Lens、Amber Reflector", "允许", "同车型同款式内"],
    ["合法变体维度", "左右侧：Driver Side / Passenger Side", "允许", "车灯天然属性，可做子体"],
    ["合法变体维度", "数量：1只 / 1对(Pair)", "允许(慎用)", "Pair 建议独立父体，避免“只收到一只”类差评"],
    ["合法变体维度", "透镜材质/是否带灯泡", "允许", "同款内，文案须写清"],
    ["禁止", "不同车型（Silverado vs Ram vs Camry）", "禁止", "不同产品，必须各自父体"],
    ["禁止", "不同年款区间且模具不同（1999-2002 vs 2007-2013）", "禁止", "同车型不同代模具=不同产品"],
    ["禁止", "不同款式（OE Style / Projector / LED / Sequential）", "禁止", "模具、价格带、搜索词都不同"],
    ["禁止", "前大灯 / 雾灯 / 尾灯 混挂", "禁止", "不同品类，必被拆"],
    ["禁止", "同父体塞 10+ 个车型", "高风险", "变体滥用，可能被抑制/强制拆分"],
], columns=["类型", "情形", "是否可捆", "理由/边界"])

three_step = pd.DataFrame([
    ["第1步", "是否同一车型（年款区间+模具相同）？", "否 → 独立父体（停下）", ""],
    ["第2步", "是否同一款式（OE/Projector/LED/Sequential）？", "否 → 独立父体（停下）", ""],
    ["第3步", "差异是否只在 颜色/左右侧/数量/带灯？", "是 → 同父体子体；否 → 独立父体", ""],
], columns=["步骤", "判定问题", "处理", ""])

# 汇总
big = F[F["前台选项数"] >= 12]
small = F[(F["前台选项数"] >= 2) & (F["前台选项数"] < 12)]
summ = pd.DataFrame([
    ["在变体组内 ASIN", int(len(inb)), "占 235 的 %.0f%%" % (len(inb) / 235 * 100)],
    ["未捆绑 ASIN", int(len(out)), ""],
    ["变体家族数", int(len(F)), ""],
    ["其中跨车型乱挂家族", int((F["结构判定"] == "跨车型乱挂").sum()),
     "识别到 ≥2 种车型，占 %.0f%%" % ((F["结构判定"] == "跨车型乱挂").mean() * 100)],
    ["大组（≥12 选项）", "%d 个 / 我方 %d 条" % (len(big), big["我方条数"].sum()),
     "星级中位 %s，评论合计 %s" % (big["星级中位"].median(), int(big["评论合计"].sum()))],
    ["小组（2-11 选项）", "%d 个 / 我方 %d 条" % (len(small), small["我方条数"].sum()),
     "星级中位 %s，评论合计 %s" % (small["星级中位"].median(), int(small["评论合计"].sum()))],
    ["单ASIN 33 条", "星级中位 %s" % S["星级"].median(),
     "评论中位 %s，评论合计 %s" % (S["评分人数"].median(), int(S["评分人数"].fillna(0).sum()))],
    ["诊断结论", "不是“变体”，是“把独立产品塞进一个父体”",
     "平均每车型仅 %.2f 个选项，说明 55/32 个选项里几乎没有真正的颜色变体" % (
         F["每车型平均选项数"].median())],
    ["抓取日期", TODAY, "amazon.com 前台实测（235/235 成功）"],
], columns=["指标", "数值", "说明"])

with pd.ExcelWriter(OUT, engine="openpyxl") as xw:
    summ.to_excel(xw, sheet_name="汇总诊断", index=False)
    F[["家族", "前台选项数", "我方条数", "识别车型数", "每车型平均选项数", "结构判定",
       "优先级", "建议父体数(按车型)", "星级中位", "评论合计", "变体维度",
       "车型示例", "动作"]].to_excel(xw, sheet_name="家族诊断与拆法", index=False)
    three_step.to_excel(xw, sheet_name="建组三步判定", index=False)
    rules.to_excel(xw, sheet_name="能捆不能捆", index=False)
    S[["ASIN", "说明", "星级", "评分人数", "优先级", "建议动作", "标题"]].to_excel(
        xw, sheet_name="33条单ASIN归位", index=False)
    # 样式
    hf, hfill = Font(bold=True, color="FFFFFF"), PatternFill("solid", fgColor="2F5597")
    thin = Side(style="thin", color="BFBFBF")
    bd = Border(left=thin, right=thin, top=thin, bottom=thin)
    lefty = ("车型示例", "动作", "理由/边界", "判定问题", "处理", "说明", "标题", "建议动作", "家族")
    red = PatternFill("solid", fgColor="FCE4E4")
    amber = PatternFill("solid", fgColor="FFF2CC")
    for ws in xw.book.worksheets:
        ws.freeze_panes = "A2"
        for c in ws[1]:
            c.font, c.fill, c.border = hf, hfill, bd
            c.alignment = Alignment(horizontal="center", vertical="center", wrap_text=True)
        ws.row_dimensions[1].height = 30
        for j, col in enumerate(ws.iter_cols(min_row=1, max_row=1), start=1):
            nm = str(col[0].value or "")
            if nm in ("车型示例", "动作", "建议动作", "标题"):
                w = 46
            elif nm in ("理由/边界", "判定问题", "处理", "说明"):
                w = 40
            elif nm == "家族":
                w = 8
            else:
                w = 13
            ws.column_dimensions[get_column_letter(j)].width = w
        headers = [str(c.value or "") for c in ws[1]]
        for row in ws.iter_rows(min_row=2):
            for i, c in enumerate(row):
                c.border = bd
                hdr = headers[i]
                c.alignment = Alignment(horizontal="left" if any(k in hdr for k in lefty) else "center",
                                        vertical="center", wrap_text=True)
                if hdr == "结构判定" and str(c.value) == "跨车型乱挂":
                    c.fill = red
                if hdr == "优先级" and str(c.value) == "高":
                    c.fill = amber
        if ws.title != "汇总诊断":
            ws.auto_filter.ref = ws.dimensions

print("家族 %d | 跨车型 %d | 大组 %d 家/我方 %d 条 星级中位 %s | 小组 星级中位 %s | 单ASIN 星级中位 %s" % (
    len(F), (F["结构判定"] == "跨车型乱挂").sum(), len(big), big["我方条数"].sum(),
    big["星级中位"].median(), small["星级中位"].median(), round(S["星级"].median(), 2)))
print("输出:", OUT)
