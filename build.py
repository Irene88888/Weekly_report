"""內銷週報產生器

用法：
    python build.py 銷貨明細.xlsx            -> 產出 index.html
    python build.py 銷貨明細.xlsx 2026-10-19 -> 指定截止日（預設為資料最後一天）

本週／上週為完整一週（週一～週五），取截止日當天或之前最近的週五為本週結束日；
本月累計則統計到截止日。

需要套件：pandas、openpyxl
"""
import json
import re
import sys
from pathlib import Path

import pandas as pd

# ---- 可調整的設定 ----
# 第 3 段「客戶動態」不列入分析的客戶（簡稱）與業務；目前全部納入，與第 1 段本週銷售額一致
# 若要排除，例如：EXCLUDE = ["辦公經費", "零售", "呂友助", "正洋電商"]、EXCLUDE_REPS = ["黃啓靖"]
# （有排除時，歸因彙總會自動加上「排除客戶」對帳列）
EXCLUDE = []
EXCLUDE_REPS = []
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


def weeks(ref):
    """完整一週＝週一到週五。取截止日當天或之前最近的週五為本週結束日"""
    fri = ref - pd.Timedelta(days=(ref.weekday() - 4) % 7)
    return fri - pd.Timedelta(days=4), fri, fri - pd.Timedelta(days=11), fri - pd.Timedelta(days=7)


def headline(df, ref):
    tw0, tw1, lw0, lw1 = weeks(ref)
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
        "tw": total(df, tw0, tw1), "lw": total(df, lw0, lw1),
        "tw_range": f"{tw0:%m/%d}–{tw1:%m/%d}", "lw_range": f"{lw0:%m/%d}–{lw1:%m/%d}",
        "pm": total(df, pm0, pm1), "pm_label": f"{pm0.month}月",
        "mtd": total(df, m0, ref), "mtd_range": f"{m0:%m/%d}–{ref:%m/%d}", "m_label": f"{m0.month}月",
        "ly_month": total(df, ly0, ly1), "ly_label": f"{ly0:%Y}年{ly0.month}月",
        "target": target, "target_basis": basis,
        "days_passed": ref.day, "days_in_month": (m0 + pd.offsets.MonthEnd(0)).day,
    }


def price_volume(df, ref):
    """本週客戶動態：兩週有交易的全部客戶，下展到品項的價量分析"""
    tw0, tw1, lw0, lw1 = weeks(ref)
    exm = df["c"].isin(EXCLUDE) | df["業務人員姓名"].isin(EXCLUDE_REPS)
    X = df[exm]
    xa = X[(X.d >= lw0) & (X.d <= lw1)].groupby("c")["表身本幣金額"].sum()
    xb = X[(X.d >= tw0) & (X.d <= tw1)].groupby("c")["表身本幣金額"].sum()
    excluded = [{"c": c, "s0": round(float(xa.get(c, 0))), "s1": round(float(xb.get(c, 0)))}
                for c in sorted(set(xa.index) | set(xb.index)) if xa.get(c, 0) or xb.get(c, 0)]
    df = df[~exm]
    A = df[(df.d >= lw0) & (df.d <= lw1)]
    B = df[(df.d >= tw0) & (df.d <= tw1)]
    k = ["c", "產品名稱", "交易單位名稱"]
    j = A.groupby(k).agg(q0=("發貨數量", "sum"), s0=("表身本幣金額", "sum")).join(
        B.groupby(k).agg(q1=("發貨數量", "sum"), s1=("表身本幣金額", "sum")), how="outer"
    ).fillna(0).reset_index()
    rep = pd.concat([A, B]).sort_values("d").groupby("c")["業務人員姓名"].last()
    out = []
    for c, g in j.groupby("c"):
        s0, s1 = g.s0.sum(), g.s1.sum()
        if round(s0) == 0 and round(s1) == 0:
            continue                                # 兩週金額皆為 0（如沖銷、贈品），不列入
        if s1 <= 0:
            status = "停單"
        elif s0 <= 0:
            status = "新出貨"
        else:
            status = "下滑" if s1 < s0 else "成長" if s1 > s0 else "持平"
        prev = df[(df.c == c) & (df.d < lw0)]
        items = []
        for _, x in g.iterrows():
            if x.q0 > 0 and x.q1 <= 0:
                attr = "全數停購" if s1 <= 0 else "部份停購"   # 客戶整週沒出貨 vs 仍有買其他品項
            elif x.q0 <= 0 and x.q1 > 0:
                attr = "新出貨客戶" if s0 <= 0 else "新增品項"  # 上週整週沒出貨的客戶 vs 老客戶多買新品項
            else:
                attr = "減量" if x.q1 < x.q0 else "增量" if x.q1 > x.q0 else "持平"
            p0 = x.s0 / x.q0 if x.q0 > 0 else None
            p1 = x.s1 / x.q1 if x.q1 > 0 else None
            price = (p1 - p0) * x.q1 if p0 and p1 else 0
            items.append({"item": x["產品名稱"], "unit": x["交易單位名稱"],
                          "q0": round(float(x.q0), 2), "q1": round(float(x.q1), 2),
                          "s0": round(x.s0), "s1": round(x.s1), "diff": round(x.s1 - x.s0),
                          "attr": attr, "price": round(price)})
        items.sort(key=lambda r: r["diff"])
        out.append({"rep": rep[c], "c": c, "s0": round(s0), "s1": round(s1), "diff": round(s1 - s0),
                    "status": status, "first": bool(status == "新出貨" and prev.empty),
                    "last": None if prev.empty else f"{prev.d.max():%Y/%m/%d}", "items": items})
    out.sort(key=lambda r: r["diff"])
    return {"rows": out, "excluded": excluded, "cust_lw": sum(r["s0"] > 0 for r in out), "cust_tw": sum(r["s1"] > 0 for r in out)}


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
