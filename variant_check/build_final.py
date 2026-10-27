# -*- coding: utf-8 -*-
"""合并两轮抓取 → 交付《ASIN前台变体捆绑核查.xlsx》
用法: python build_final.py --p1 full_pd.json --p2 final_pd.json --asins asins.txt --out xxx.xlsx
"""
import argparse
import json
import re
from datetime import date
from pathlib import Path

import pandas as pd
from openpyxl.styles import Alignment, Border, Font, PatternFill, Side
from openpyxl.utils import get_column_letter



MODELS = ["Silverado", "Sierra", "Ram", "F-150", "F150", "F-250", "F250", "F-350", "Ranger",
          "Tundra", "Tacoma", "Camry", "Corolla", "Accord", "Civic", "Malibu", "Impala",
          "Sonata", "Elantra", "Fusion", "Focus", "Mustang", "Altima", "Maxima", "Mazda",
          "Cherokee", "Compass", "Wrangler", "Trailblazer", "Colorado", "Canyon", "Pathfinder",
          "Sentra", "RAV4", "CR-V", "CRV", "Equinox", "Escape", "Journey", "Frontier", "Titan",
          "4Runner", "Highlander", "Escalade", "Yukon", "Tahoe", "Suburban", "Explorer",
          "Expedition", "Outback", "Forester", "Santa Fe", "Tucson", "Sportage", "Optima",
          "Forte", "Jetta", "Passat", "Golf", "Tiguan", "Cruze", "Sonic", "Dart", "Charger",
          "Challenger", "Durango", "Avalanche", "Blazer", "Legacy", "Soul", "Camry", "Avalon",
          "Impreza", "Crosstrek", "Envoy", "Bravada", "Mountaineer", "Navigator", "Armada",
          "Quest", "Odyssey", "Pilot", "Ridgeline", "Sienna", "Prius", "Yaris", "Versa",
          "Cube", "Juke", "Rogue", "Murano", "CX-3", "CX-5", "CX-7", "CX-9", "Mazda3", "Mazda6",
          "Terrain", "Grand Caravan", "Town & Country", "Town and Country", "K1500", "C1500", "C/K", "S10",
          "Grand Cherokee", "Liberty", "Commander", "Nitro", "Caliber", "Avenger", "Sebring",
          "PT Cruiser", "300", "Rio", "Accent", "Veloster", "Lancer", "Mirage", "Outlander",
          "Eclipse", "Xterra", "Armada", "Ridgeline", "Pilot", "Odyssey", "Sienna", "Prius",
          "Yaris", "Versa", "370Z", "350Z", "G35", "G37", "TL", "TSX", "TSX", "ILX", "RDX",
          "MDX", "IS250", "IS350", "ES350", "GS350", "MKX", "MKZ", "Navigator", "Aviator",
          "Escort", "Contour", "Taurus", "Crown Victoria", "Grand Marquis", "Town Car",
          "DeVille", "DTS", "CTS", "SRX", "XLR", "Aveo", "Optra", "Epica", "Lacetti"]


def models_in(text):
    t = " " + str(text) + " "
    hit = []
    for m in MODELS:
        if re.search(r"(?<![A-Za-z])" + re.escape(m) + r"(?![A-Za-z])", t, re.I):
            key = m.upper().replace("-", "").replace(" ", "")
            if key not in [h[0] for h in hit]:
                hit.append((key, m))
    return [h[1] for h in hit]


def num(s):
    m = re.search(r"([\d.]+)", str(s if s is not None else ""))
    return float(m.group(1)) if m else None


def cnt(s):
    m = re.search(r"([\d,]+)", str(s if s is not None else ""))
    return int(m.group(1).replace(",", "")) if m else None


class UF:
    def __init__(self):
        self.p = {}

    def find(self, x):
        self.p.setdefault(x, x)
        while self.p[x] != x:
            self.p[x] = self.p[self.p[x]]
            x = self.p[x]
        return x

    def union(self, a, b):
        ra, rb = self.find(a), self.find(b)
        if ra != rb:
            self.p[rb] = ra


def main():
    ap = argparse.ArgumentParser()
    ap.add_argument("--p1", required=True)
    ap.add_argument("--p2", required=True)
    ap.add_argument("--asins", required=True)
    ap.add_argument("--out", required=True)
    a = ap.parse_args()
    today = str(date.today())

    asins = [x.strip() for x in Path(a.asins).read_text(encoding="utf-8").replace("\n", ",").split(",") if x.strip()]
    p1 = json.loads(Path(a.p1).read_text(encoding="utf-8"))
    p2 = json.loads(Path(a.p2).read_text(encoding="utf-8"))
    r1, r2 = p1.get("rows", {}), p2.get("rows", {})

    def sup_n(root, ours):
        return max([len(mem_map.get(x, set())) for x in ours] + [0])

    rec, mem_map, vals_map = [], {}, {}
    for sid in asins:
        d1, d2 = r1.get(sid), r2.get(sid)
        if not d1:
            rec.append({"ASIN": sid, "前台是否捆绑变体": "未抓到", "前台家族ASIN数": None,
                        "前台落地ASIN": "", "变体维度": "", "是否含自身": "",
                        "核查说明": "未抓到", "前台星级": None, "前台评分人数": None,
                        "标题": "", "品牌": "", "采集日期": today})
            continue
        n1 = int(d1.get("nVariantAsins") or 0)
        n2 = int((d2 or {}).get("nStateAsins") or 0)
        n2d = int((d2 or {}).get("nDomAsins") or 0)
        members = set()
        for x in (d1.get("variantAsins") or []):
            if x:
                members.add(x)
        for x in ((d2 or {}).get("stateAsins") or []):
            if x:
                members.add(x)
        for x in ((d2 or {}).get("domAsins") or []):
            if x:
                members.add(x)
        n = max(n1, n2, n2d, len(members))
        dims = [x for x in (d1.get("dims") or []) if x.get("dim") != "collapsedView"]
        vals = []
        for x in dims:
            for v in (x.get("vals") or []):
                if v not in vals:
                    vals.append(v)
        for v in ((d2 or {}).get("vals") or []):
            if v not in vals:
                vals.append(v)
        dimnames = "/".join(x["dim"] for x in dims) or "/".join((d2 or {}).get("dimNames") or [])
        note = []
        if d1.get("landedAsin") and d1["landedAsin"] != sid:
            note.append("前台落地 %s" % d1["landedAsin"])
        if n1 < 2 and n >= 2:
            note.append("首轮漏判→复抓修正为 %d 个" % n)
        if n == 1:
            only = next(iter(members), "")
            note.append("有变体选择器但仅1个选项(%s)" % (
                ("指向 %s" % only) if only and only != sid else "即自身"))
            if only and only != sid:
                note.append("⚠️残留在他人变体组")
        elif n == 0:
            note.append("前台无变体选择器")
        rec.append({"ASIN": sid,
                    "前台是否捆绑变体": "是" if n >= 2 else "否",
                    "前台家族ASIN数": max(n, 1),
                    "前台落地ASIN": d1.get("landedAsin", ""),
                    "变体维度": dimnames,
                    "是否含自身": ("是" if sid in members else ("否" if n >= 2 else "—")),
                    "核查说明": "；".join(note) or "正常",
                    "前台星级": num(d1.get("rating")) or num((d2 or {}).get("rating")),
                    "前台评分人数": cnt(d1.get("ratingCount")) or cnt((d2 or {}).get("ratingCount")),
                    "标题": d1.get("title", ""),
                    "品牌": (d1.get("brand") or "").replace("Visit the ", "").replace(" Store", ""),
                    "采集日期": today})
        if members:
            mem_map[sid] = members
        if vals:
            vals_map[sid] = vals

    det = pd.DataFrame(rec)
    det["前台家族ASIN数"] = pd.to_numeric(det["前台家族ASIN数"], errors="coerce")

    uf = UF()
    for sid, mem in mem_map.items():
        uf.find(sid)
        for m in mem:
            uf.union(sid, m)
    comp = {}
    for sid in det["ASIN"]:
        comp.setdefault(uf.find(sid), []).append(sid)

    # 只给「捆绑」的家族顺序编号（按我方ASIN数降序）
    bundled = [(root, ours) for root, ours in comp.items()
               if len(ours) >= 2 and sup_n(root, ours) >= 2]
    bundled.sort(key=lambda kv: -len(set(kv[1])))
    famno, fammem = {}, {}
    for i, (root, ours) in enumerate(bundled, 1):
        mem = set()
        for x in ours:
            mem |= mem_map.get(x, set())
        famno[root] = i
        fammem[root] = mem

    def root_of(sid):
        return uf.find(sid)

    bundled_asins = set(det.loc[det["前台是否捆绑变体"] == "是", "ASIN"])
    solo = sorted([x for x in bundled_asins if root_of(x) not in famno])

    det["家族编号"] = det["ASIN"].apply(
        lambda s: famno.get(root_of(s), ("—(与清单外ASIN同组)" if s in bundled_asins else "")))
    det["同家族我方ASIN数"] = det["ASIN"].apply(lambda s: len(set(comp.get(root_of(s), [s]))))
    det["同家族清单内其他ASIN"] = det["ASIN"].apply(
        lambda s: "/".join(sorted([x for x in set(comp.get(root_of(s), [])) if x != s])))
    det["家族成员ASIN"] = det["ASIN"].apply(
        lambda s: "/".join(sorted(fammem.get(root_of(s), mem_map.get(s, set())))))
    det["自身在家族选项内"] = det.apply(
        lambda r: ("是" if r["ASIN"] in mem_map.get(r["ASIN"], set()) else
                   ("否" if r["前台是否捆绑变体"] == "是" else "—")), axis=1)
    det.loc[det["自身在家族选项内"] == "否", "核查说明"] = det.loc[
        det["自身在家族选项内"] == "否", "核查说明"].astype(str).str.replace("正常", "", regex=False).str.strip("；;") + \
        "；⚠️页面展示该变体家族但自身不在选项中(疑似被并入/隐藏)"

    for x in solo:
        i = det.index[det["ASIN"] == x][0]
        old = str(det.at[i, "核查说明"]).strip()
        old = "" if old in ("正常", "nan") else old.rstrip("；;")
        det.at[i, "核查说明"] = ("%s；与清单外ASIN同组(共%d个)" % (
            old, int(det.at[i, "前台家族ASIN数"] or 0))).lstrip("；")
    det = det.sort_values(["前台是否捆绑变体", "家族编号", "ASIN"],
                          ascending=[True, True, True]).reset_index(drop=True)
    cols = ["ASIN", "前台是否捆绑变体", "家族编号", "前台家族ASIN数", "同家族我方ASIN数",
            "自身在家族选项内", "前台落地ASIN", "变体维度", "同家族清单内其他ASIN", "家族成员ASIN",
            "前台星级", "前台评分人数", "品牌", "标题", "核查说明", "采集日期"]
    det = det[[c for c in cols if c in det.columns]]
    yes = det[det["前台是否捆绑变体"] == "是"].copy()
    no = det[det["前台是否捆绑变体"] == "否"].copy()

    # 家族表
    fams = []
    for root, ours in sorted(bundled, key=lambda kv: famno[kv[0]]):
        ours = sorted(set(ours) & bundled_asins)
        sub = det[det["ASIN"].isin(ours)]
        n = sub["前台家族ASIN数"].max()
        n = int(n) if pd.notna(n) else len(ours)
        mem = sorted(fammem[root])
        n = max(n, len(mem))
        notin = [x for x in ours if x not in set(mem)]
        fams.append({"家族编号": famno[root], "前台家族ASIN数": n,
                     "清单内我方ASIN数": len(ours),
                     "超出清单的成员数": max(n - len(ours), 0),
                     "我方不在选项内的ASIN": "/".join(notin) or "",
                     "变体维度": next((x for x in sub["变体维度"] if x), ""),
                     "变体值数": len({v for x in ours for v in vals_map.get(x, [])}),
                     "变体值示例(前8)": " ｜ ".join(
                         list(dict.fromkeys([v for x in ours for v in vals_map.get(x, [])]))[:8]),
                     "清单内我方ASIN": "/".join(ours),
                     "前台星级": sub["前台星级"].dropna().iloc[0] if sub["前台星级"].notna().any() else None,
                     "前台评分人数": sub["前台评分人数"].dropna().iloc[0] if sub["前台评分人数"].notna().any() else None,
                     "家族成员ASIN": "/".join(mem)[:1200]})
    famdf = pd.DataFrame(fams)

    if len(famdf):
        cross, mlist = [], []
        for _, r in famdf.iterrows():
            ours = str(r["清单内我方ASIN"]).split("/")
            txt = str(r.get("变体值示例(前8)", "")) + " " + " ".join(
                det.loc[det["ASIN"].isin(ours), "标题"].astype(str).tolist())
            ms = models_in(txt)
            mlist.append("/".join(ms[:8]))
            cross.append("是" if len(ms) >= 2 else ("待人工核对" if ms else "未识别"))
        famdf["疑似跨车型"] = cross
        famdf["识别到的车型"] = mlist
        famdf = famdf[["家族编号", "前台家族ASIN数", "清单内我方ASIN数", "超出清单的成员数",
                       "疑似跨车型", "识别到的车型", "变体维度", "变体值数", "变体值示例(前8)",
                       "清单内我方ASIN", "我方不在选项内的ASIN", "前台星级", "前台评分人数",
                       "家族成员ASIN"]]

    fix = det[det["核查说明"].str.contains("漏判|残留|不在选项内|前台落地", na=False)][
        ["ASIN", "前台是否捆绑变体", "家族编号", "前台家族ASIN数", "核查说明"]].reset_index(drop=True)

    # 汇总
    sizes = famdf["前台家族ASIN数"].value_counts().sort_index(ascending=False) if len(famdf) else pd.Series(dtype=int)
    srows = [["核查ASIN总数", len(det), "源表 工作簿4.xlsx"],
             ["前台捆绑在变体组内", len(yes), "变体选择器可选 ASIN ≥ 2"],
             ["前台未捆绑(单ASIN)", len(no), "无选择器或仅1个选项"],
             ["涉及的变体家族数", len(famdf), "家族内我方链接≥2条（自己人同组）"],
             ["最大家族ASIN数", int(famdf["前台家族ASIN数"].max()) if len(famdf) else 0, "家族1"],
             ["仅与清单外ASIN同组的ASIN数", len(solo), "同组但清单内无第二条我方链接"],
             ["疑似跨车型乱捆绑的家族数",
              int((famdf["疑似跨车型"] == "是").sum()) if len(famdf) else 0,
              "同一变体组内挂 ≥2 种车型（如 Silverado 与 Ram 同组）"],
             ["捆绑ASIN合计", len(yes), "家族内我方 %d + 仅与外部同组 %d" % (len(yes) - len(solo), len(solo))],
             ["核查时间", today, "agent-browser 直抓 amazon.com 前台"],
             ["口径说明", "以详情页变体选择器实际可切换的不同 ASIN 数判定，≥2 即捆绑；"
                         "首轮全部235条抓取后再对53条“疑似单ASIN”复抓校验，修正1条", ""]]
    summ = pd.DataFrame(srows, columns=["指标", "数值", "说明"])
    dist = pd.DataFrame({"家族ASIN数": [int(i) for i in sizes.index],
                         "家族数": [int(v) for v in sizes.values]})

    out = Path(a.out)
    with pd.ExcelWriter(out, engine="openpyxl") as xw:
        summ.to_excel(xw, sheet_name="汇总", index=False, startrow=0)
        dist.to_excel(xw, sheet_name="汇总", index=False, startrow=len(summ) + 3)
        yes.to_excel(xw, sheet_name="捆绑清单", index=False)
        famdf.to_excel(xw, sheet_name="变体家族", index=False)
        no.to_excel(xw, sheet_name="未捆绑清单", index=False)
        fix.to_excel(xw, sheet_name="修正与异常说明", index=False)

        hf, hfill = Font(bold=True, color="FFFFFF"), PatternFill("solid", fgColor="2F5597")
        thin = Side(style="thin", color="BFBFBF")
        bd = Border(left=thin, right=thin, top=thin, bottom=thin)
        lefty = ("ASIN", "前台落地ASIN", "标题", "核查说明", "清单内我方ASIN",
                 "家族成员ASIN", "同家族清单内其他ASIN", "指标", "说明", "口径说明")
        warn = PatternFill("solid", fgColor="FFF2CC")
        for ws in xw.book.worksheets:
            ws.freeze_panes = "A2"
            if ws.title == "汇总":
                ws.column_dimensions["A"].width = 24
                ws.column_dimensions["B"].width = 14
                ws.column_dimensions["C"].width = 60
                for c in ws[1]:
                    c.font, c.fill, c.border = hf, hfill, bd
                    c.alignment = Alignment(horizontal="center", vertical="center")
                for row in ws.iter_rows(min_row=2, max_row=len(summ) + 1):
                    for i, c in enumerate(row):
                        c.border = bd
                        c.alignment = Alignment(horizontal="left" if i != 1 else "center",
                                                vertical="center", wrap_text=True)
                for c in ws[len(summ) + 4]:
                    c.font, c.fill, c.border = hf, hfill, bd
                    c.alignment = Alignment(horizontal="center", vertical="center")
                continue
            for c in ws[1]:
                c.font, c.fill, c.border = hf, hfill, bd
                c.alignment = Alignment(horizontal="center", vertical="center", wrap_text=True)
            ws.row_dimensions[1].height = 30
            for j, col in enumerate(ws.iter_cols(min_row=1, max_row=1), start=1):
                nm = str(col[0].value or "")
                if nm == "标题":
                    w = 44
                elif "成员ASIN" in nm or "其他ASIN" in nm or "我方ASIN" in nm:
                    w = 42
                elif "说明" in nm:
                    w = 40
                elif "ASIN" in nm:
                    w = 15
                else:
                    w = 13
                ws.column_dimensions[get_column_letter(j)].width = w
            for row in ws.iter_rows(min_row=2):
                for i, c in enumerate(row):
                    c.border = bd
                    hdr = str(ws.cell(row=1, column=i + 1).value or "")
                    c.alignment = Alignment(horizontal="left" if any(k in hdr for k in lefty) else "center",
                                            vertical="center", wrap_text=True)
            ws.auto_filter.ref = ws.dimensions

    print("总计 %d | 捆绑 %d | 未捆绑 %d | 家族数 %d" % (len(det), len(yes), len(no), len(famdf)))
    print("家族规模分布:", {int(k): int(v) for k, v in sizes.items()})
    print("输出:", out)


if __name__ == "__main__":
    main()
