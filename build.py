"""內銷週報產生器

用法：
    python build.py 銷貨明細.xlsx            -> 產出 index.html
    python build.py 銷貨明細.xlsx 2026-10-19 -> 指定截止日（預設為資料最後一天）

需要套件：pandas、openpyxl
"""
import json
import re
import sys
from pathlib import Path

import pandas as pd

# ---- 可調整的設定 ----
# 第 3 段「客戶動態」不列入分析的客戶（簡稱）與業務；第 1、2 段的銷售額為全體內銷，不受影響
EXCLUDE = ["辦公經費", "零售", "呂友助", "正洋電商"]
EXCLUDE_REPS = ["黃啓靖"]
# 月目標（元）。有設定就用目標算完成率；沒設定則以「去年同月全月銷售額」為基準
TARGETS = {
    # "2026-10": 20_000_000,
}
RENAME = {"那波麗士股份有限公司(合點壽司)": "那波麗士(合點壽司)"}


def short(name):
    name = re.sub(r"(股份有限公司|有限公司|企業社|公司)$", "", str(name))
    return RENAME.get(name, name)


def load(path):
    df = pd.read_excel(path)
    df["d"] = pd.to_datetime(df["單據日期"])
    df = df[df["交易幣別"] == "TWD"].copy()          # 內銷
    df["c"] = df["客戶名稱"].map(short)
    return df


def total(df, a, b):
    return float(df[(df.d >= a) & (df.d <= b)]["表身本幣金額"].sum())


def headline(df, ref):
    tw0, lw1, lw0 = ref - pd.Timedelta(days=6), ref - pd.Timedelta(days=7), ref - pd.Timedelta(days=13)
    m0 = ref.replace(day=1)
    pm1 = m0 - pd.Timedelta(days=1)
    pm0 = pm1.replace(day=1)
    ly0 = m0 - pd.DateOffset(years=1)
    ly1 = ly0 + pd.offsets.MonthEnd(0)
    key = f"{ref:%Y-%m}"
    if key in TARGETS:
        target, basis = float(TARGETS[key]), "月目標"
    else:
        target, basis = total(df, ly0, ly1), f"去年{ly0.month}月全月"
    return {
        "tw": total(df, tw0, ref), "lw": total(df, lw0, lw1),
        "tw_range": f"{tw0:%m/%d}–{ref:%m/%d}", "lw_range": f"{lw0:%m/%d}–{lw1:%m/%d}",
        "pm": total(df, pm0, pm1), "pm_label": f"{pm0.month}月",
        "mtd": total(df, m0, ref), "mtd_range": f"{m0:%m/%d}–{ref:%m/%d}", "m_label": f"{m0.month}月",
        "ly_month": total(df, ly0, ly1), "ly_label": f"{ly0:%Y}年{ly0.month}月",
        "target": target, "target_basis": basis,
        "days_passed": ref.day, "days_in_month": (m0 + pd.offsets.MonthEnd(0)).day,
    }


def price_volume(df, ref):
    """本週銷售額下滑客戶，下展到品項的價量分析"""
    df = df[~df["c"].isin(EXCLUDE) & ~df["業務人員姓名"].isin(EXCLUDE_REPS)]
    tw0, lw1, lw0 = ref - pd.Timedelta(days=6), ref - pd.Timedelta(days=7), ref - pd.Timedelta(days=13)
    A = df[(df.d >= lw0) & (df.d <= lw1)]
    B = df[(df.d >= tw0) & (df.d <= ref)]
    k = ["c", "產品名稱", "交易單位名稱"]
    j = A.groupby(k).agg(q0=("發貨數量", "sum"), s0=("表身本幣金額", "sum")).join(
        B.groupby(k).agg(q1=("發貨數量", "sum"), s1=("表身本幣金額", "sum")), how="outer"
    ).fillna(0).reset_index()
    rep = pd.concat([A, B]).sort_values("d").groupby("c")["業務人員姓名"].last()
    out = []
    for c, g in j.groupby("c"):
        s0, s1 = g.s0.sum(), g.s1.sum()
        if s1 >= s0:
            continue
        items = []
        for _, x in g.iterrows():
            if x.q0 > 0 and x.q1 <= 0:
                attr = "全數停購" if s1 <= 0 else "部份停購"   # 客戶整週沒出貨 vs 仍有買其他品項
            elif x.q0 <= 0 and x.q1 > 0:
                attr = "新增品項"
            else:
                attr = "減量" if x.q1 < x.q0 else "增量"
            p0 = x.s0 / x.q0 if x.q0 > 0 else None
            p1 = x.s1 / x.q1 if x.q1 > 0 else None
            price = (p1 - p0) * x.q1 if p0 and p1 else 0
            items.append({"item": x["產品名稱"], "unit": x["交易單位名稱"],
                          "q0": round(float(x.q0), 2), "q1": round(float(x.q1), 2),
                          "p0": round(p0, 2) if p0 else None, "p1": round(p1, 2) if p1 else None,
                          "s0": round(x.s0), "s1": round(x.s1), "diff": round(x.s1 - x.s0),
                          "attr": attr, "price": round(price)})
        items.sort(key=lambda r: r["diff"])
        out.append({"rep": rep[c], "c": c, "s0": round(s0), "s1": round(s1),
                    "diff": round(s1 - s0), "stopped": bool(s1 <= 0), "items": items})
    out.sort(key=lambda r: r["diff"])
    return {"rows": out, "cust_lw": int(A.c.nunique()), "cust_tw": int(B.c.nunique())}


def main():
    df = load(Path(sys.argv[1]))
    ref = pd.Timestamp(sys.argv[2]) if len(sys.argv) > 2 else df.d.max().normalize()
    data = {"ref": f"{ref:%Y/%m/%d}", "head": headline(df, ref), "pv": price_volume(df, ref),
            "exclude": EXCLUDE + [f"業務 {r}" for r in EXCLUDE_REPS]}
    here = Path(__file__).parent
    tpl = (here / "template.html").read_text(encoding="utf-8")
    (here / "index.html").write_text(tpl.replace("/*__DATA__*/null", json.dumps(data, ensure_ascii=False)),
                                     encoding="utf-8")
    print(f"已產出 index.html（截止日 {data['ref']}）")


if __name__ == "__main__":
    main()
