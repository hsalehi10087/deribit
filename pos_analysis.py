#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
پایش و پرستاری پایانه‌های فروش (POS) و درگاه‌های پرداخت (IPG)

- POS فقط با POS و IPG فقط با IPG مقایسه می‌شود.
- مبلغ تراکنش و معدل هر مشتری در نمودارهای جداگانه نمایش داده می‌شوند.
- کد شعبه در همه فهرست‌ها آمده تا هر گزارش به شعبه مربوطه داده شود.

ورودی : هر تعداد فایل اکسل ماهانه (POS و/یا IPG)
خروجی : فایل اکسل گزارش

اجرا:
    python pos_analysis.py file1.xlsx file2.xlsx ... -o pos_monitoring_report.xlsx
"""
import argparse
import re

import numpy as np
import pandas as pd
from openpyxl import Workbook
from openpyxl.chart import BarChart, LineChart, Reference
from openpyxl.formatting.rule import FormulaRule
from openpyxl.worksheet.datavalidation import DataValidation
from openpyxl.worksheet.table import Table, TableStyleInfo
from openpyxl.styles import Alignment, Border, Font, PatternFill, Side
from openpyxl.utils import get_column_letter
from openpyxl.utils.dataframe import dataframe_to_rows

# ------------------------------------------------------------------
# تنظیمات و آستانه‌ها
# ------------------------------------------------------------------
CONFIG = {
    "rial_per_toman": 10,                  # داده‌ها ریال هستند؛ هر تومان = ۱۰ ریال
    "min_avg_balance_toman": 30_000_000,   # حداقل معدل برای طرح وام صبا
    "drop_pct": 0.30,                      # افت ≥ ۳۰٪ نسبت به دوره قبل = افت شدید
    "stable_band": 0.10,                   # تغییر بین ±۱۰٪ = ثابت
    "weak_terminal_ratio": 0.20,           # دستگاهی با کارکرد < ۲۰٪ قوی‌ترین، ضعیف
    "top_n": 10,                           # تعداد مشتری در هر نمودار
    "risk_weight_churn": 40,
    "risk_weight_decline": 30,
    "risk_weight_low_volume": 30,
}

# رنگ‌های گزارش
FILL = {
    "header": PatternFill("solid", fgColor="1F4E78"),
    "title": PatternFill("solid", fgColor="DDEBF7"),
    "churn": PatternFill("solid", fgColor="FFC7CE"),
    "drop": PatternFill("solid", fgColor="F8CBAD"),
    "down": PatternFill("solid", fgColor="FFF2CC"),
    "up": PatternFill("solid", fgColor="C6EFCE"),
    "w1": PatternFill("solid", fgColor="F4B183"),
    "w2": PatternFill("solid", fgColor="FFE699"),
    "w3": PatternFill("solid", fgColor="DDEBF7"),
    "saba": PatternFill("solid", fgColor="E2EFDA"),
}
THIN = Side(style="thin", color="7F7F7F")
BORDER = Border(left=THIN, right=THIN, top=THIN, bottom=THIN)


# ------------------------------------------------------------------
# خواندن فایل‌ها
# ------------------------------------------------------------------
def clean_id(x):
    """کد مشتری/حساب/شعبه را به رشته تمیز تبدیل می‌کند."""
    if pd.isna(x):
        return ""
    if isinstance(x, float) and x.is_integer():
        x = int(x)
    out = str(x).strip().replace(" ", "")
    if out.isdigit() and len(out) == 10 and out.startswith("9"):
        out = "0" + out          # شماره موبایل بدون صفر اول ذخیره شده است
    return out


def to_num(s):
    return pd.to_numeric(s, errors="coerce").fillna(0)


def find_col(df, name, last=False, prefix=False):
    idx = [i for i, c in enumerate(df.columns)
           if (str(c).startswith(name) if prefix else str(c).strip() == name)]
    if not idx:
        raise KeyError(f"ستون پیدا نشد: {name}")
    return df.iloc[:, idx[-1] if last else idx[0]]


def read_snapshot(path):
    raw = pd.read_excel(path, header=None, dtype=object)
    title = str(raw.iloc[0].dropna().iloc[0]).strip()
    m = re.search(r"(\d{6})\s*$", title)
    if not m:
        raise ValueError(f"دوره در عنوان فایل پیدا نشد: {path}")
    period = f"{m.group(1)[:4]}/{m.group(1)[4:]}"
    kind = "POS" if "POS" in title else ("IPG" if "IPG" in title else "?")
    is_header = raw.apply(lambda r: r.astype(str).str.strip().eq("نوع تراکنش").any(), axis=1)
    hdr = raw.index[is_header][0]
    df = raw.loc[hdr + 1:].reset_index(drop=True)
    df.columns = [str(c).strip() for c in raw.loc[hdr]]
    return kind, period, df


def parse_pos(df, period):
    out = pd.DataFrame({
        "period": period,
        "terminal": to_num(find_col(df, "کد ترمینال", last=True)).astype("int64"),
        "branch": find_col(df, "کد شعبه").map(clean_id),
        "cust": find_col(df, "شماره تعریف مشتری").map(clean_id),
        "acct": find_col(df, "شماره حساب").map(clean_id),
        "store": find_col(df, "نام فروشگاه").fillna("").astype(str).str.strip(),
        "owner": (find_col(df, "نام").fillna("").astype(str) + " "
                  + find_col(df, "نام خانوادگی").fillna("").astype(str)).str.strip(),
        "mobile": find_col(df, "شماره تلفن همراه").map(clean_id),
        "guild": find_col(df, "نام صنف").fillna("").astype(str).str.strip(),
        "model": find_col(df, "مدل پایانه").fillna("").astype(str).str.strip(),
        "cnt": to_num(find_col(df, "تعداد تراکنشها")),
        "amt": to_num(find_col(df, "مبلغ تراکنش ها")),
        "bal": to_num(find_col(df, "میانگین یک ماهه حساب")),
    })
    return out[~out["cust"].isin(["", "0"])].reset_index(drop=True)


def parse_ipg(df, period):
    out = pd.DataFrame({
        "period": period,
        "branch": find_col(df, "کد شعبه").map(clean_id),
        "acct": find_col(df, "شماره حساب").map(clean_id),
        "owner": find_col(df, "نام دارنده حساب").fillna("").astype(str).str.strip(),
        "store": find_col(df, "نام فروشگاه").fillna("").astype(str).str.strip(),
        "cnt": to_num(find_col(df, "تعداد تراکنش موفق")),
        "amt": to_num(find_col(df, "مبلغ تراکنش")),
        "bal": to_num(find_col(df, "میانگین مانده حساب طی ماه", prefix=True)),
    })
    return out[~out["acct"].isin(["", "0"])].reset_index(drop=True)


# ------------------------------------------------------------------
# ابزارهای مشترک
# ------------------------------------------------------------------
def trend_label(churn, g, band, drop):
    return np.select(
        [churn, g.isna(), g <= -drop, g <= -band, g >= band],
        ["ریزش (بدون تراکنش)", "جدید / بازگشته", "افت شدید", "رو به کاهش", "رو به رشد"],
        default="ثابت",
    )


def priority_table(cand_idx, value_m, weights_df, weights, action_map):
    """امتیاز اولویت و موج تماس: ۶۰٪ ارزش مشتری + ۴۰٪ نوع علت."""
    rw = pd.Series([max([weights[c] for c in weights_df.columns if weights_df.at[i, c]], default=0)
                    for i in weights_df.index], index=weights_df.index)
    cand = rw.index[(rw > 0).values]
    cand = [i for i in cand if i in cand_idx]
    if not cand:
        return cand, pd.Series(dtype=float), rw
    vpct = value_m.loc[cand].rank(pct=True)
    score = {i: round(60 * vpct[i] + 40 * rw[i]) for i in cand}
    return cand, pd.Series(score), rw


def wave(score):
    return "موج ۱ (فوری)" if score >= 70 else ("موج ۲" if score >= 40 else "موج ۳")


# ------------------------------------------------------------------
# تحلیل POS (فقط دستگاه‌های POS)
# ------------------------------------------------------------------
def analyse_pos(pos, cfg):
    periods = sorted(pos["period"].unique())
    L, P = periods[-1], periods[-2]
    rial = cfg["rial_per_toman"]
    min_bal_rial = cfg["min_avg_balance_toman"] * rial
    drop, band = cfg["drop_pct"], cfg["stable_band"]
    M = lambda s: (s / (rial * 1e6)).round(2)

    cp = pos.groupby(["cust", "period"]).agg(
        terminals=("terminal", "nunique"), cnt=("cnt", "sum"), amt=("amt", "sum")).reset_index()
    bal_cp = (pos.drop_duplicates(["cust", "period", "acct"])
              .groupby(["cust", "period"])["bal"].sum().reset_index())
    cp = cp.merge(bal_cp, on=["cust", "period"], how="left")

    def wide(col):
        return cp.pivot(index="cust", columns="period", values=col).reindex(columns=periods)

    present = wide("cnt").notna()
    terms = wide("terminals").fillna(0)
    cnt = wide("cnt").fillna(0)
    amt = wide("amt").fillna(0)
    bal = wide("bal").fillna(0)
    active = cnt > 0
    active_L = active[L]
    churn = (~active_L) & active.drop(columns=[L]).any(axis=1)

    # تغییر کارکرد فقط روی پایانه‌های مشترک دو دوره
    tp = pos.groupby(["cust", "terminal", "period"]).agg(cnt=("cnt", "sum"), amt=("amt", "sum")).reset_index()
    a = tp[tp.period == P].set_index(["cust", "terminal"])[["cnt", "amt"]]
    b = tp[tp.period == L].set_index(["cust", "terminal"])[["cnt", "amt"]]
    common = a.index.intersection(b.index)
    ca = a.loc[common].groupby(level=0).sum()
    cb = b.loc[common].groupby(level=0).sum()

    g_amt = (amt[L] / amt[P].replace(0, np.nan) - 1)
    g_cnt = (cnt[L] / cnt[P].replace(0, np.nan) - 1)
    g_amt = (cb["amt"] / ca["amt"].replace(0, np.nan) - 1).reindex(amt.index).combine_first(g_amt)
    g_cnt = (cb["cnt"] / ca["cnt"].replace(0, np.nan) - 1).reindex(amt.index).combine_first(g_cnt)
    g_amt[churn] = -1.0
    g_cnt[churn] = -1.0
    trend = pd.Series(trend_label(churn, g_amt, band, drop), index=amt.index)

    # معدل و روند آن (آخرین مقدار مثبت ثبت‌شده)
    def bal_info(i):
        vals = [bal.at[i, p] for p in periods if present.at[i, p] and bal.at[i, p] > 0]
        if not vals:
            return 0.0, False, 0.0
        up = len(vals) >= 2 and vals[-1] > vals[0] and vals[-1] >= vals[-2]
        return vals[-1], up, vals[0]

    info = [bal_info(i) for i in amt.index]
    bal_last = pd.Series([x[0] for x in info], index=amt.index)
    bal_up = pd.Series([x[1] for x in info], index=amt.index)
    bal_first = pd.Series([x[2] for x in info], index=amt.index)

    potential = (bal_last >= min_bal_rial) & bal_up
    decline_flag = churn | (g_amt <= -drop).fillna(False)
    saba = potential & decline_flag

    # پایانه‌های غایب (دوره آخر نیستند)
    tw = tp.pivot_table(index=["cust", "terminal"], columns="period", values="amt",
                        aggfunc="sum").reindex(columns=periods)
    seen = tw.notna()
    tw_info = pos.sort_values("period").groupby(["cust", "terminal"])[["store", "model", "branch"]].last().reindex(tw.index)
    term_last_seen = [next((p for p in reversed(periods) if seen.at[k, p]), "") for k in tw.index]
    term_missing_L = (seen.any(axis=1) & ~seen[L]).values
    term_df = pd.DataFrame({
        "کد شعبه": tw_info["branch"].fillna("").values,
        "کد مشتری": [k[0] for k in tw.index],
        "کد پایانه": [k[1] for k in tw.index],
        "نام فروشگاه": tw_info["store"].fillna("").values,
        "مدل پایانه": tw_info["model"].fillna("").values,
        "آخرین دوره دیده‌شده": term_last_seen,
        "وضعیت در آخرین دوره": np.where(seen[L].values, "فعال", "غایب"),
    })
    for p in periods:
        term_df[f"مبلغ {p} (میلیون تومان)"] = M(tw[p].fillna(0)).values
    term_df = term_df.sort_values(["وضعیت در آخرین دوره", "کد شعبه"]).reset_index(drop=True)

    # نام، شعبه و اطلاعات نمایشی
    names = pos.sort_values("period").groupby("cust")[["owner", "store", "guild", "mobile", "branch"]].last()
    names = names.reindex(amt.index).fillna("")

    last_active_s = pd.Series({i: next((p for p in reversed(periods) if active.at[i, p]), "")
                               for i in amt.index})
    months_out = pd.Series({i: periods.index(L) - periods.index(last_active_s[i]) if last_active_s[i] else np.nan
                            for i in amt.index})

    # نمره ریسک
    q25 = amt[L][active_L].quantile(0.25)
    decl = np.clip(-g_amt.fillna(0), 0, 1)
    risk = (cfg["risk_weight_churn"] * churn.astype(int)
            + cfg["risk_weight_decline"] * decl
            + cfg["risk_weight_low_volume"] * (active_L & (amt[L] <= q25)).astype(int))
    risk = risk.clip(0, 100).round(0)

    # برتر / ضعیف / پتانسیل
    top_idx = amt[L][active_L].sort_values(ascending=False).index[: cfg["top_n"]]
    weak_all = churn | trend.isin(["افت شدید"])
    weak_rank = pd.DataFrame({"risk": risk, "amtP": amt[P]})[weak_all]
    weak_top = weak_rank.sort_values(["risk", "amtP"], ascending=False).index[: cfg["top_n"]]
    pot_top = bal_last[potential].sort_values(ascending=False).index[: cfg["top_n"]]

    # برچسب‌ها
    tags = pd.DataFrame(False, index=amt.index, columns=[
        "برتر", "ضعیف", "کاندید صبا", "دارای پتانسیل", "ریزش", "چند پوز"])
    tags.loc[top_idx, "برتر"] = True
    tags["ضعیف"] = weak_all
    tags["کاندید صبا"] = saba
    tags["دارای پتانسیل"] = potential & ~saba
    tags["ریزش"] = churn
    tags["چند پوز"] = terms[L] >= 2
    labels = tags.apply(lambda r: "، ".join([c for c in tags.columns if r[c]]), axis=1)

    # جدول اصلی مشتریان
    cust = pd.DataFrame({
        "کد شعبه": names["branch"].values,
        "کد مشتری": amt.index,
        "صاحب پایانه": names["owner"].values,
        "نام فروشگاه": names["store"].values,
        "صنف": names["guild"].values,
        "موبایل": names["mobile"].values,
        f"تعداد پایانه ({L})": terms[L].astype(int).values,
    }, index=amt.index)
    for p in periods:
        cust[f"تراکنش {p}"] = cnt[p].astype(int).values
        cust[f"مبلغ {p} (میلیون تومان)"] = M(amt[p]).values
        cust[f"معدل {p} (میلیون تومان)"] = M(bal[p]).values
    cust["تغییر مبلغ (%)"] = g_amt.round(4).values
    cust["تغییر تراکنش (%)"] = g_cnt.round(4).values
    cust["وضعیت روند"] = trend.values
    cust["آخرین دوره فعال"] = last_active_s.values
    cust["نمره ریسک (۰-۱۰۰)"] = risk.values
    cust["دسته‌ها"] = labels.values
    cust = cust.sort_values("نمره ریسک (۰-۱۰۰)", ascending=False).reset_index(drop=True)

    # فهرست‌های اختصاصی
    churn_df = pd.DataFrame({
        "کد شعبه": names["branch"][churn].values,
        "کد مشتری": amt.index[churn], "صاحب پایانه": names["owner"][churn].values,
        "نام فروشگاه": names["store"][churn].values, "موبایل": names["mobile"][churn].values,
        "آخرین دوره فعال": last_active_s[churn].values,
        f"مبلغ آخرین دوره فعال (میلیون تومان)": M(amt.apply(lambda r: r[[p for p in periods if r[p] > 0][-1]] if (r > 0).any() else 0, axis=1)[churn]).values,
        "معدل آخرین دوره (میلیون تومان)": M(bal_last[churn]).values,
        "نمره ریسک": risk[churn].values,
    }).sort_values("مبلغ آخرین دوره فعال (میلیون تومان)", ascending=False)

    def bal_block(mask):
        df = pd.DataFrame({
            "کد شعبه": names["branch"][mask].values,
            "کد مشتری": amt.index[mask], "صاحب پایانه": names["owner"][mask].values,
            "نام فروشگاه": names["store"][mask].values, "موبایل": names["mobile"][mask].values,
            "معدل اول (میلیون تومان)": M(bal_first[mask]).values,
            "معدل آخر (میلیون تومان)": M(bal_last[mask]).values,
            "تغییر معدل (%)": (bal_last[mask] / bal_first[mask].replace(0, np.nan) - 1).round(3).values,
            f"مبلغ {L} (میلیون تومان)": M(amt[L][mask]).values,
            "تغییر مبلغ (%)": g_amt[mask].round(3).values,
            "وضعیت روند": trend[mask].values,
            "تعداد پایانه": terms[L][mask].astype(int).values,
        })
        return df.sort_values("معدل آخر (میلیون تومان)", ascending=False)

    saba_df = bal_block(saba)
    potential_df = bal_block(potential)
    potential_df["کاندید صبا"] = np.where(saba[potential].values, "بله", "خیر")

    # چند پوز (آخرین دوره‌ای که جزئیات چند پایانه دارد)
    counts = pos.groupby("period").agg(c=("cust", "nunique"), t=("terminal", "nunique"))
    detail_p = [p for p in periods if counts.at[p, "t"] > counts.at[p, "c"]]
    D = detail_p[-1] if detail_p else L
    Dprev = [p for p in periods if p < D][-1] if any(p < D for p in periods) else None
    td = tp[tp.period == D].copy()
    td = td[td.groupby("cust").terminal.transform("nunique") >= 2].copy()
    td_info = pos[pos.period == D].groupby(["cust", "terminal"]).agg(
        store=("store", "last"), model=("model", "last"), branch=("branch", "last")).reset_index()
    td = td.merge(td_info, on=["cust", "terminal"], how="left")
    if Dprev:
        prev = tp[tp.period == Dprev][["cust", "terminal", "amt"]].rename(columns={"amt": "amt_prev"})
        td = td.merge(prev, on=["cust", "terminal"], how="left")
    else:
        td["amt_prev"] = np.nan
    td["max_amt"] = td.groupby("cust").amt.transform("max")
    td["min_amt"] = td.groupby("cust").amt.transform("min")
    td["ratio"] = (td.amt / td.max_amt).round(3)
    td["weakest"] = (td.amt == td.min_amt) & (td["ratio"] < cfg["weak_terminal_ratio"])
    td["chg"] = np.where(td.amt_prev > 0, td.amt / td.amt_prev - 1, np.nan)
    td = td.sort_values(["cust", "amt"], ascending=[True, True])
    multi_df = pd.DataFrame({
        "کد شعبه": td.branch, "کد مشتری": td.cust, "کد پایانه": td.terminal,
        "نام فروشگاه": td.store, "مدل پایانه": td.model,
        "تراکنش": td.cnt.astype(int),
        f"مبلغ {D} (میلیون تومان)": M(td.amt),
        f"مبلغ {Dprev or '-'} (میلیون تومان)": M(td.amt_prev),
        "تغییر نسبت به دوره قبل (%)": td.chg.round(3),
        "نسبت به قوی‌ترین پایانه": td.ratio,
        "وضعیت": np.where(td.weakest, "ضعیف‌ترین دستگاه مشتری", ""),
    }).reset_index(drop=True)

    # فهرست تماس POS
    miss_map = {}
    for k, is_missing in zip(tw.index, term_missing_L):
        if is_missing:
            miss_map.setdefault(k[0], []).append(str(k[1]))
    miss_count = pd.Series({c: len(v) for c, v in miss_map.items()}).reindex(amt.index, fill_value=0)
    value_m = amt.max(axis=1) / (rial * 1e6)
    reasons = pd.DataFrame({
        "خرابی دستگاه احتمالی": active_L & (miss_count > 0),
        "قهر/ترک احتمالی (معدل قبلی ثبت شده)": churn & (bal_last > 0),
        "افت کارکرد ≥۳۰٪": active_L & (g_amt <= -drop).fillna(False),
        "خروج کامل احتمالی": churn & (bal_last <= 0),
    }, index=amt.index)
    weights = {"خرابی دستگاه احتمالی": 1.0, "قهر/ترک احتمالی (معدل قبلی ثبت شده)": 0.9,
               "افت کارکرد ≥۳۰٪": 0.7, "خروج کامل احتمالی": 0.4}
    action_map = {
        "خرابی دستگاه احتمالی": "بررسی فنی پایانه‌های غایب و تماس برای رفع مشکل",
        "قهر/ترک احتمالی (معدل قبلی ثبت شده)": "تماس برای بازگشت مشتری و رفع نارضایتی",
        "افت کارکرد ≥۳۰٪": "تماس برای بررسی علت کاهش کارکرد",
        "خروج کامل احتمالی": "تماس برای تأیید ترک و وضعیت حساب",
    }
    cand, scores, rw = priority_table(set(amt.index), value_m, reasons, weights, action_map)
    call_rows = []
    for i in cand:
        rs = [c for c in reasons.columns if reasons.at[i, c]]
        primary = max(rs, key=lambda c: weights[c])
        sc = int(scores[i])
        call_rows.append({
            "موج تماس": wave(sc),
            "امتیاز اولویت (۰-۱۰۰)": sc,
            "کد شعبه": names.at[i, "branch"],
            "اقدام پیشنهادی": action_map[primary],
            "علت‌های احتمالی": "، ".join(rs),
            "کد مشتری": i,
            "صاحب پایانه": names.at[i, "owner"],
            "نام فروشگاه": names.at[i, "store"],
            "موبایل": names.at[i, "mobile"],
            "آخرین دوره فعال": last_active_s[i],
            "دوره‌های غیرفعال": int(months_out[i]) if not pd.isna(months_out[i]) else "",
            "پایانه‌های غایب": "، ".join(miss_map.get(i, [])),
            "حداکثر مبلغ ماهانه (میلیون تومان)": round(value_m[i], 2),
            "معدل آخر (میلیون تومان)": M(pd.Series([bal_last[i]]))[0],
            f"مبلغ {L} (میلیون تومان)": M(pd.Series([amt[L][i]]))[0],
            "وضعیت روند": trend[i],
            "علت عدم کار (پس از تماس)": "",
        })
    call_df = pd.DataFrame(call_rows)
    if not call_df.empty:
        call_df = call_df.sort_values("امتیاز اولویت (۰-۱۰۰)", ascending=False).reset_index(drop=True)

    # نمودارها: مبلغ و معدل جداگانه
    def chart_pack(idx):
        return {
            "custs": [f"{names.at[i, 'store'] or names.at[i, 'owner']} ({i})" for i in idx],
            "amt": [[round(amt.at[i, p] / (rial * 1e6), 2) for p in periods] for i in idx],
            "bal": [[round(bal.at[i, p] / (rial * 1e6), 2) for p in periods] for i in idx],
        }

    charts = {
        "برتر": chart_pack(top_idx),
        "ضعیف": chart_pack(weak_top),
        "پتانسیل": chart_pack(pot_top),
    }

    summary = {
        "periods": periods, "L": L, "P": P, "D": D, "Dprev": Dprev,
        "rows": [
            ("کل مشتریان POS (همه دوره‌ها)", len(amt)),
            (f"مشتریان فعال در {L}", int(active_L.sum())),
            (f"ریزش (فعال قبلی، بدون تراکنش در {L})", int(churn.sum())),
            ("مشتریان ضعیف (ریزش + افت شدید ≥۳۰٪)", int(weak_all.sum())),
            ("کاندید وام صبا", int(saba.sum())),
            ("مشتریان دارای پتانسیل (معدل بالا و افزایشی)", int(potential.sum())),
            (f"مشتریان چندپوز ({D})", int(td.cust.nunique())),
            (f"دستگاه‌های ضعیف‌تر در چندپوزها ({D})", int(td.weakest.sum())),
            ("فهرست تماس POS (کل)", len(call_df)),
        ],
        "period_stats": counts,
        "reason_counts": {c: int(reasons[c].sum()) for c in reasons.columns},
    }
    cust_i = cust.set_index("کد مشتری")
    grp_cols = ["کد شعبه", "صاحب پایانه", "نام فروشگاه", f"مبلغ {L} (میلیون تومان)",
                f"معدل {L} (میلیون تومان)", "تغییر مبلغ (%)", "وضعیت روند", "نمره ریسک (۰-۱۰۰)"]
    groups = {"برتر": cust_i.loc[list(top_idx), grp_cols].reset_index(),
              "ضعیف": cust_i.loc[list(weak_top), grp_cols].reset_index(),
              "پتانسیل": cust_i.loc[list(pot_top), grp_cols].reset_index()}
    return {
        "summary": summary, "customers": cust, "churn": churn_df, "saba": saba_df,
        "potential": potential_df, "multi": multi_df, "terminals": term_df,
        "calls": call_df, "charts": charts, "groups": groups,
    }


# ------------------------------------------------------------------
# تحلیل IPG (فقط درگاه‌های پرداخت)
# ------------------------------------------------------------------
def analyse_ipg(ipg, cfg):
    periods = sorted(ipg["period"].unique())
    L = periods[-1]
    P = periods[-2] if len(periods) >= 2 else None
    rial = cfg["rial_per_toman"]
    drop, band = cfg["drop_pct"], cfg["stable_band"]
    M = lambda s: (s / (rial * 1e6)).round(2)

    g = ipg.groupby(["acct", "period"]).agg(cnt=("cnt", "sum"), amt=("amt", "sum"),
                                            bal=("bal", "max")).reset_index()
    W = lambda c: g.pivot(index="acct", columns="period", values=c).reindex(columns=periods).fillna(0)
    cnt, amt, bal = W("cnt"), W("amt"), W("bal")
    names = ipg.sort_values("period").groupby("acct")[["owner", "store", "branch"]].last()
    names = names.reindex(amt.index).fillna("")
    active = amt > 0
    active_L = active[L]
    churn = (~active_L) & active.drop(columns=[L]).any(axis=1) if P else pd.Series(False, index=amt.index)

    if P:
        g_amt = amt[L] / amt[P].replace(0, np.nan) - 1
    else:
        g_amt = pd.Series(np.nan, index=amt.index)
    g_amt[churn] = -1.0
    trend = pd.Series(trend_label(churn, g_amt, band, drop), index=amt.index)

    bal_last = pd.Series([next((bal.at[i, p] for p in reversed(periods) if bal.at[i, p] > 0), 0.0)
                          for i in amt.index], index=amt.index)
    last_active_s = pd.Series({i: next((p for p in reversed(periods) if active.at[i, p]), "")
                               for i in amt.index})

    accounts = pd.DataFrame({
        "کد شعبه": names["branch"].values,
        "کد حساب": amt.index,
        "دارنده حساب": names["owner"].values,
        "نام فروشگاه": names["store"].values,
    }, index=amt.index)
    for p in periods:
        accounts[f"تراکنش {p}"] = cnt[p].astype(int).values
        accounts[f"مبلغ {p} (میلیون تومان)"] = M(amt[p]).values
        accounts[f"معدل {p} (میلیون تومان)"] = M(bal[p]).values
    accounts["تغییر مبلغ (%)"] = g_amt.round(4).values
    accounts["وضعیت روند"] = trend.values
    accounts["آخرین دوره فعال"] = last_active_s.values
    accounts = accounts.sort_values(f"مبلغ {L} (میلیون تومان)", ascending=False).reset_index(drop=True)

    reasons = pd.DataFrame({
        "قهر/ترک احتمالی (معدل قبلی ثبت شده)": churn & (bal_last > 0),
        "افت کارکرد ≥۳۰٪": active_L & (g_amt <= -drop).fillna(False),
        "خروج کامل احتمالی": churn & (bal_last <= 0),
    }, index=amt.index)
    weights = {"قهر/ترک احتمالی (معدل قبلی ثبت شده)": 0.9, "افت کارکرد ≥۳۰٪": 0.7,
               "خروج کامل احتمالی": 0.4}
    action_map = {
        "قهر/ترک احتمالی (معدل قبلی ثبت شده)": "تماس با صاحب درگاه برای بازگشت و بررسی علت قطع",
        "افت کارکرد ≥۳۰٪": "تماس برای بررسی علت کاهش تراکنش‌های درگاه",
        "خروج کامل احتمالی": "تماس برای تأیید ترک و وضعیت حساب",
    }
    value_m = amt.max(axis=1) / (rial * 1e6)
    cand, scores, rw = priority_table(set(amt.index), value_m, reasons, weights, action_map)
    call_rows = []
    for i in cand:
        rs = [c for c in reasons.columns if reasons.at[i, c]]
        primary = max(rs, key=lambda c: weights[c])
        sc = int(scores[i])
        call_rows.append({
            "موج تماس": wave(sc),
            "امتیاز اولویت (۰-۱۰۰)": sc,
            "کد شعبه": names.at[i, "branch"],
            "اقدام پیشنهادی": action_map[primary],
            "علت‌های احتمالی": "، ".join(rs),
            "کد حساب": i,
            "دارنده حساب": names.at[i, "owner"],
            "نام فروشگاه": names.at[i, "store"],
            "آخرین دوره فعال": last_active_s[i],
            "حداکثر مبلغ ماهانه (میلیون تومان)": round(value_m[i], 2),
            "معدل آخر (میلیون تومان)": M(pd.Series([bal_last[i]]))[0],
            f"مبلغ {L} (میلیون تومان)": M(pd.Series([amt[L][i]]))[0],
            "وضعیت روند": trend[i],
            "علت عدم کار (پس از تماس)": "",
        })
    call_df = pd.DataFrame(call_rows)
    if not call_df.empty:
        call_df = call_df.sort_values("امتیاز اولویت (۰-۱۰۰)", ascending=False).reset_index(drop=True)

    def chart_pack(idx):
        return {
            "custs": [f"{names.at[i, 'store'] or names.at[i, 'owner']} ({i})" for i in idx],
            "amt": [[round(amt.at[i, p] / (rial * 1e6), 2) for p in periods] for i in idx],
            "bal": [[round(bal.at[i, p] / (rial * 1e6), 2) for p in periods] for i in idx],
        }

    top_idx = amt[L][active_L].sort_values(ascending=False).index[: cfg["top_n"]]
    weak_idx = list(call_df["کد حساب"][:cfg["top_n"]]) if not call_df.empty else []
    charts = {"برتر": chart_pack(top_idx), "ضعیف": chart_pack(weak_idx)}

    summary = {
        "periods": periods,
        "rows": [
            ("کل حساب‌های درگاه (همه دوره‌ها)", len(amt)),
            (f"حساب‌های فعال در {L}", int(active_L.sum())),
            (f"قطع یا ریزش در {L}", int(churn.sum())),
            ("حساب‌های با افت کارکرد ≥۳۰٪", int(reasons["افت کارکرد ≥۳۰٪"].sum())),
            ("فهرست تماس درگاه (کل)", len(call_df)),
        ],
    }
    acc_i = accounts.set_index("کد حساب")
    ipg_cols = ["کد شعبه", "دارنده حساب", "نام فروشگاه", f"مبلغ {L} (میلیون تومان)",
                f"معدل {L} (میلیون تومان)", "تغییر مبلغ (%)", "وضعیت روند"]
    groups = {"برتر": acc_i.loc[list(top_idx), ipg_cols].reset_index(),
              "ضعیف": acc_i.loc[list(weak_idx), ipg_cols].reset_index()}
    return {"summary": summary, "accounts": accounts, "groups": groups, "calls": call_df, "charts": charts}


# ------------------------------------------------------------------
# نوشتن اکسل: جدول‌های خوانا، نمای مشتری، داشبورد
# ------------------------------------------------------------------
BODY_FONT = Font(name="Calibri", size=12)
HEAD_FONT = Font(name="Calibri", size=12, bold=True, color="FFFFFF")
TITLE_FONT = Font(name="Calibri", size=18, bold=True, color="1F4E78")
SUB_FONT = Font(name="Calibri", size=14, bold=True, color="1F4E78")
LABEL_FONT = Font(name="Calibri", size=12, bold=True)
_table_counter = [0]

TREND_FILL = {
    "ریزش (بدون تراکنش)": FILL["churn"], "افت شدید": FILL["drop"],
    "رو به کاهش": FILL["down"], "رو به رشد": FILL["up"],
    "قهر/ترک احتمالی (معدل قبلی ثبت شده)": FILL["churn"],
    "خروج کامل احتمالی": FILL["drop"],
}
WAVE_FILL = {"موج ۱ (فوری)": FILL["w1"], "موج ۲": FILL["w2"], "موج ۳": FILL["w3"]}


def new_sheet(wb, name):
    ws = wb.create_sheet(name[:31])
    ws.sheet_view.rightToLeft = True
    return ws


def fill_cells(ws, first_row, last_row, col_idx, mapping):
    for r in range(first_row, last_row + 1):
        v = ws.cell(row=r, column=col_idx).value
        if v in mapping:
            ws.cell(row=r, column=col_idx).fill = mapping[v]


def write_table(ws, df, start_row=1, title=None, as_table=True):
    """جدول خوانا با هدر رنگی، حاشیه، فونت بزرگ و (در صورت نیاز) جدول اکسل. برمی‌گرداند ردیف آخر."""
    r = start_row
    if title:
        ws.cell(row=r, column=1, value=title).font = SUB_FONT
        ws.row_dimensions[r].height = 28
        r += 1
    if df.empty:
        ws.cell(row=r, column=1, value="موردی یافت نشد").font = BODY_FONT
        return r
    header_row = r
    for c_idx, col in enumerate(df.columns, start=1):
        cell = ws.cell(row=r, column=c_idx, value=str(col))
        cell.font, cell.fill, cell.border = HEAD_FONT, FILL["header"], BORDER
        cell.alignment = Alignment(horizontal="center", vertical="center", wrap_text=True)
    ws.row_dimensions[header_row].height = 48
    for row in dataframe_to_rows(df, index=False, header=False):
        r += 1
        for c_idx, val in enumerate(row, start=1):
            if isinstance(val, float) and (np.isnan(val) or np.isinf(val)):
                val = None
            cell = ws.cell(row=r, column=c_idx, value=val)
            cell.font, cell.border = BODY_FONT, BORDER
            cell.alignment = Alignment(vertical="center")
        ws.row_dimensions[r].height = 24
    last = r
    for c_idx, col in enumerate(df.columns, start=1):
        letter = get_column_letter(c_idx)
        want = min(50, max(16, len(str(col)) * 1.4))
        cur = ws.column_dimensions[letter].width or 0
        ws.column_dimensions[letter].width = max(cur, want)
        fmt = None
        if "%" in str(col):
            fmt = "0.0%"
        elif "میلیون" in str(col):
            fmt = "#,##0.00"
        elif any(k in str(col) for k in ["تراکنش", "تعداد", "نمره", "امتیاز", "دوره‌های"]):
            fmt = "#,##0"
        if fmt:
            for rr in range(header_row + 1, last + 1):
                ws[f"{letter}{rr}"].number_format = fmt
        if col in ("وضعیت روند", "وضعیت"):
            fill_cells(ws, header_row + 1, last, c_idx, TREND_FILL)
        if col == "موج تماس":
            fill_cells(ws, header_row + 1, last, c_idx, WAVE_FILL)
    if as_table and last > header_row:
        _table_counter[0] += 1
        tbl = Table(displayName=f"Tbl{_table_counter[0]}",
                    ref=f"A{header_row}:{get_column_letter(len(df.columns))}{last}")
        tbl.tableStyleInfo = TableStyleInfo(name="TableStyleLight9", showRowStripes=True)
        ws.add_table(tbl)
    return last


def attach_calls(base, calls, key):
    """ستون‌های تماس (موج، امتیاز، علت، اقدام) را کنار جدول اصلی می‌آورد."""
    extra = ["موج تماس", "امتیاز اولویت (۰-۱۰۰)", "علت‌های احتمالی", "اقدام پیشنهادی"]
    if calls.empty:
        out = base.copy()
        for c in extra:
            out[c] = ""
    else:
        out = base.merge(calls[[key] + extra], on=key, how="left")
        for c in extra:
            out[c] = out[c].fillna("")
    return out[[key] + [c for c in out.columns if c != key]]


def build_viewer(wb, sheet_name, title, data_df, data_sheet, periods, info_cols, id_col):
    """نمای تک‌مشتری: با انتخاب کد، اطلاعات، پیشنهاد و نمودار مبلغ/معدل همان مشتری نشان داده می‌شود."""
    ds = sheet(wb, data_sheet)
    HDR = 4
    write_table(ds, data_df, HDR)
    n = len(data_df)
    cols = list(data_df.columns)

    def rng(col):
        L_ = get_column_letter(cols.index(col) + 1)
        return f"'{data_sheet}'!${L_}${HDR + 1}:${L_}${HDR + n}"

    id_rng = rng(id_col)
    ws = sheet(wb, sheet_name)
    ws.sheet_view.showGridLines = False
    ws["A1"] = title
    ws["A1"].font = TITLE_FONT
    ws["A3"] = "کد مشتری را از لیست انتخاب کنید ←"
    ws["A3"].font = LABEL_FONT
    ws["B3"] = data_df[id_col].iloc[0]
    ws["B3"].font = Font(name="Calibri", size=16, bold=True, color="1F4E78")
    ws["B3"].fill = PatternFill("solid", fgColor="FFF2CC")
    ws["B3"].alignment = Alignment(horizontal="center")
    thick = Side(style="medium", color="1F4E78")
    ws["B3"].border = Border(left=thick, right=thick, top=thick, bottom=thick)
    ws.row_dimensions[3].height = 30
    dv = DataValidation(type="list", formula1=id_rng, allow_blank=False)
    dv.error = "کد مشتری را از لیست انتخاب کنید."
    ws.add_data_validation(dv)
    dv.add("B3")

    r = 5
    status_row = None
    for col in info_cols:
        lc = ws.cell(row=r, column=1, value=col)
        lc.font, lc.fill, lc.border = LABEL_FONT, FILL["title"], BORDER
        ref_v = f'INDEX({rng(col)},MATCH($B$3,{id_rng},0))'
        vc = ws.cell(row=r, column=2, value=f'=IFERROR(IF({ref_v}="","-",{ref_v}),"-")')
        vc.font, vc.border = BODY_FONT, BORDER
        vc.alignment = Alignment(horizontal="right")
        if "%" in col:
            vc.number_format = "0.0%"
        ws.row_dimensions[r].height = 24
        if col == "وضعیت روند":
            status_row = r
        r += 1

    sr = status_row or 5
    lc = ws.cell(row=r, column=1, value="پیشنهاد هوشمند")
    lc.font, lc.fill, lc.border = LABEL_FONT, PatternFill("solid", fgColor="E2EFDA"), BORDER
    vc = ws.cell(row=r, column=2, value=(
        f'=IF(ISNUMBER(SEARCH("ریزش",B{sr})),"بدون تراکنش؛ تماس برای بازگشت و بررسی علت قطع",'
        f'IF(ISNUMBER(SEARCH("افت",B{sr})),"کارکرد کاهش یافته؛ تماس برای بررسی علت و معرفی طرح‌ها",'
        f'IF(ISNUMBER(SEARCH("رو به کاهش",B{sr})),"روند نزولی؛ در دوره بعد پایش شود","وضعیت عادی؛ ادامه پایش")))'))
    vc.font, vc.fill, vc.border = Font(name="Calibri", size=12, bold=True), PatternFill("solid", fgColor="E2EFDA"), BORDER
    vc.alignment = Alignment(wrap_text=True, vertical="center")
    ws.row_dimensions[r].height = 40
    ws.conditional_formatting.add(f"B{sr}", FormulaRule(formula=[f'ISNUMBER(SEARCH("ریزش",B{sr}))'], fill=FILL["churn"]))
    ws.conditional_formatting.add(f"B{sr}", FormulaRule(formula=[f'ISNUMBER(SEARCH("افت",B{sr}))'], fill=FILL["drop"]))
    ws.conditional_formatting.add(f"B{sr}", FormulaRule(formula=[f'ISNUMBER(SEARCH("رو به رشد",B{sr}))'], fill=FILL["up"]))

    r += 3
    hdr = r
    heads = ["دوره", "مبلغ تراکنش (میلیون تومان)", "معدل حساب (میلیون تومان)", "تعداد تراکنش"]
    for j, h in enumerate(heads, start=1):
        c = ws.cell(row=hdr, column=j, value=h)
        c.font, c.fill, c.border = HEAD_FONT, FILL["header"], BORDER
        c.alignment = Alignment(horizontal="center", vertical="center", wrap_text=True)
    ws.row_dimensions[hdr].height = 40
    for p in periods:
        r += 1
        ws.cell(row=r, column=1, value=p).font = BODY_FONT
        ws.cell(row=r, column=2, value=f'=IFERROR(INDEX({rng(f"مبلغ {p} (میلیون تومان)")},MATCH($B$3,{id_rng},0)),0)')
        ws.cell(row=r, column=3, value=f'=IFERROR(INDEX({rng(f"معدل {p} (میلیون تومان)")},MATCH($B$3,{id_rng},0)),0)')
        ws.cell(row=r, column=4, value=f'=IFERROR(INDEX({rng(f"تراکنش {p}")},MATCH($B$3,{id_rng},0)),0)')
        for j in range(1, 5):
            cell = ws.cell(row=r, column=j)
            cell.border = BORDER
            cell.font = BODY_FONT
            if j in (2, 3):
                cell.number_format = "#,##0.00"
            if j == 4:
                cell.number_format = "#,##0"
        ws.row_dimensions[r].height = 24
    last = r
    ws.column_dimensions["A"].width = 40
    ws.column_dimensions["B"].width = 42
    ws.column_dimensions["C"].width = 30
    ws.column_dimensions["D"].width = 20

    bar = BarChart()
    bar.type = "col"
    bar.title = "مبلغ تراکنش (ستون) و معدل حساب (خط)"
    bar.add_data(Reference(ws, min_col=2, min_row=hdr, max_row=last), titles_from_data=True)
    bar.set_categories(Reference(ws, min_col=1, min_row=hdr + 1, max_row=last))
    bar.y_axis.title = "مبلغ تراکنش (میلیون تومان)"
    bar.x_axis.title = "دوره"
    line = LineChart()
    line.add_data(Reference(ws, min_col=3, min_row=hdr, max_row=last), titles_from_data=True)
    line.y_axis.axId = 200
    line.y_axis.title = "معدل حساب (میلیون تومان)"
    line.y_axis.crosses = "max"
    bar.x_axis.delete = False
    bar.y_axis.delete = False
    line.y_axis.delete = False
    bar += line
    bar.height, bar.width = 12, 24
    bar.legend.position = "b"
    ws.add_chart(bar, "F3")
    return ws


def write_kpi(ws, df, start_row, title):
    """جدول خلاصه: دو ستون، فونت بزرگ، ردیف‌های بلند و مقدار وسط‌چین."""
    ws.cell(row=start_row, column=1, value=title).font = SUB_FONT
    ws.row_dimensions[start_row].height = 30
    r = start_row + 1
    for j, h in enumerate(df.columns, start=1):
        c = ws.cell(row=r, column=j, value=h)
        c.font, c.fill, c.border = HEAD_FONT, FILL["header"], BORDER
        c.alignment = Alignment(horizontal="center", vertical="center")
    ws.row_dimensions[r].height = 30
    for k, (label, val) in enumerate(df.itertuples(index=False), start=1):
        r += 1
        band = PatternFill("solid", fgColor="F2F2F2") if k % 2 == 0 else None
        a = ws.cell(row=r, column=1, value=label)
        b = ws.cell(row=r, column=2, value=val)
        a.font = Font(name="Calibri", size=13)
        b.font = Font(name="Calibri", size=14, bold=True, color="1F4E78")
        b.alignment = Alignment(horizontal="center", vertical="center")
        b.number_format = "#,##0"
        for c in (a, b):
            c.border = BORDER
            c.alignment = Alignment(vertical="center", horizontal=c.alignment.horizontal, wrap_text=True)
            if band:
                c.fill = band
        ws.row_dimensions[r].height = 28
    ws.column_dimensions["A"].width = max(ws.column_dimensions["A"].width or 0, 62)
    ws.column_dimensions["B"].width = max(ws.column_dimensions["B"].width or 0, 22)
    return r


def add_dashboard(wb, res_pos, res_ipg, cfg):
    ws = sheet(wb, "داشبورد")
    ws.sheet_view.zoomScale = 100
    ws["A1"] = "داشبورد پایش و پرستاری پایانه‌ها و درگاه‌ها"
    ws["A1"].font = Font(bold=True, size=16, color="1F4E78")
    sp = res_pos["summary"]
    row = 3
    ws.cell(row=row, column=1, value="دوره‌های POS").font = Font(bold=True)
    ws.cell(row=row, column=2, value="، ".join(sp["periods"]))
    row += 1
    if res_ipg:
        ws.cell(row=row, column=1, value="دوره‌های IPG").font = Font(bold=True)
        ws.cell(row=row, column=2, value="، ".join(res_ipg["summary"]["periods"]))
        row += 1
    row += 1

    kpi_pos = pd.DataFrame(sp["rows"], columns=["شاخص POS", "مقدار"])
    row = write_kpi(ws, kpi_pos, row, "خلاصه POS (دستگاه‌های پوز)") + 2
    if res_ipg:
        kpi_ipg = pd.DataFrame(res_ipg["summary"]["rows"], columns=["شاخص درگاه", "مقدار"])
        row = write_kpi(ws, kpi_ipg, row, "خلاصه IPG (درگاه‌های پرداخت)") + 2

    # جدول شعبه‌ها: گزارش هر شعبه
    cust = res_pos["customers"]
    L = sp["L"]
    active_col = f"تراکنش {L}"
    branch_rows = {}
    for b, grp in cust.groupby("کد شعبه"):
        branch_rows.setdefault(b, {})
        branch_rows[b]["مشتریان POS فعال"] = int((grp[active_col] > 0).sum())
    calls = res_pos["calls"]
    if not calls.empty:
        for b, grp in calls.groupby("کد شعبه"):
            branch_rows.setdefault(b, {})
            branch_rows[b]["تماس POS - موج ۱"] = int((grp["موج تماس"] == "موج ۱ (فوری)").sum())
            branch_rows[b]["تماس POS - موج ۲"] = int((grp["موج تماس"] == "موج ۲").sum())
            branch_rows[b]["تماس POS - موج ۳"] = int((grp["موج تماس"] == "موج ۳").sum())
    churn = res_pos["churn"]
    for b, grp in churn.groupby("کد شعبه"):
        branch_rows.setdefault(b, {})
        branch_rows[b]["ریزش POS"] = len(grp)
    saba = res_pos["saba"]
    for b, grp in saba.groupby("کد شعبه"):
        branch_rows.setdefault(b, {})
        branch_rows[b]["کاندید صبا"] = len(grp)
    if res_ipg and not res_ipg["calls"].empty:
        for b, grp in res_ipg["calls"].groupby("کد شعبه"):
            branch_rows.setdefault(b, {})
            branch_rows[b]["تماس درگاه - موج ۱"] = int((grp["موج تماس"] == "موج ۱ (فوری)").sum())
            branch_rows[b]["تماس درگاه - موج ۲"] = int((grp["موج تماس"] == "موج ۲").sum())
            branch_rows[b]["تماس درگاه - موج ۳"] = int((grp["موج تماس"] == "موج ۳").sum())

    bdf = pd.DataFrame.from_dict(branch_rows, orient="index").fillna(0)
    bdf.index.name = "کد شعبه"
    bdf = bdf.reset_index()
    bdf.insert(0, "کد شعبه", bdf.pop("کد شعبه"))
    for c in bdf.columns[1:]:
        bdf[c] = bdf[c].astype(int)
    wcols = [c for c in bdf.columns if c.startswith("تماس POS - موج") or c.startswith("تماس درگاه - موج")]
    bdf["جمع تماس‌های فوری"] = bdf.get("تماس POS - موج ۱", 0) + bdf.get("تماس درگاه - موج ۱", 0)
    bdf = bdf.sort_values(["جمع تماس‌های فوری"] + wcols, ascending=False).reset_index(drop=True)
    row = write_table(ws, bdf, row, title="گزارش هر شعبه (برای ارجاع تماس‌ها)") + 2

    # ۱۵ مورد اول فهرست تماس POS
    top = calls.head(15)[["موج تماس", "امتیاز اولویت (۰-۱۰۰)", "کد شعبه", "نام فروشگاه", "موبایل",
                          "علت‌های احتمالی", "اقدام پیشنهادی", "معدل آخر (میلیون تومان)",
                          f"مبلغ {L} (میلیون تومان)"]] if not calls.empty else calls
    row = write_table(ws, top, row, title="۱۵ مشتری اولویت‌دار POS") + 2
    ws.cell(row=row, column=1, value="نکات:").font = Font(bold=True)
    notes = [
        "POS فقط با POS و IPG فقط با IPG مقایسه شده است.",
        "مبلغ تراکنش و معدل هر مشتری در شیت‌های نمودار به‌صورت جداگانه آمده‌اند.",
        "پایانه‌های غایب (در دوره‌های قبل بوده ولی در دوره آخر نیستند) به‌عنوان احتمال خرابی دستگاه آمده‌اند.",
        "اگر فایل آخرین دوره POS دوباره با همه پایانه‌ها صادر شود، فهرست خرابی دستگاه دقیق‌تر می‌شود.",
        f"مبالغ ریال‌اند و در گزارش به میلیون تومان نمایش داده شده‌اند (هر تومان = {cfg['rial_per_toman']} ریال).",
    ]
    for t in notes:
        row += 1
        ws.cell(row=row, column=1, value="• " + t)
    ws.column_dimensions["A"].width = 16
    ws.column_dimensions["B"].width = 20


NOTES = {
    "داشبورد": "این صفحه خلاصه کل گزارش است. جدول «گزارش هر شعبه» نشان می‌دهد هر شعبه چند تماس فوری دارد و هر تماس باید به کدام شعبه ارجاع شود.",
    "نمایش_مشتری_POS": "کد مشتری را از لیست کشویی سلول زرد انتخاب کنید. اطلاعات، وضعیت، پیشنهاد و نمودار مبلغ و معدل همان مشتری نشان داده می‌شود.",
    "فهرست_تماس_POS": "مشتریانی که باید تماس بگیرید، بر اساس امتیاز اولویت مرتب شده‌اند. موج ۱ یعنی تماس فوری. ستون «علت عدم کار» را بعد از تماس پر کنید.",
    "برتر_ضعیف_پتانسیل_POS": "سه فهرست: مشتریان برتر (بیشترین کارکرد)، ضعیف (کارکرد افت کرده یا قطع شده) و پتانسیل (معدل بالا و رو به رشد).",
    "روند_POS": "وضعیت هر مشتری در هر ماه شامل تعداد تراکنش، مبلغ و معدل. ستون «وضعیت روند» نشان می‌دهد کارکرد رو به رشد، ثابت، کاهشی یا قطع است.",
    "ریزش_POS": "مشتریانی که قبلاً تراکنش داشتند ولی در ماه آخر تراکنشی ندارند. از بیشترین مبلغ قبلی به کمترین مرتب شده‌اند.",
    "کاندید_صبا": "مشتریانی با معدل حداقل ۳۰ میلیون تومان که معدلشان رو به افزایش است، ولی کارکرد پایانه‌شان کم شده. برای پیشنهاد وام صبا مناسب‌اند.",
    "پتانسیل_POS": "مشتریانی با معدل بالا (حداقل ۳۰ میلیون تومان) که رو به رشدند. ستون «کاندید صبا» مشخص می‌کند کدام‌ها برای وام مناسب‌اند.",
    "چند_پوز": "مشتریانی که بیش از یک دستگاه POS دارند. برای هر دستگاه مبلغ و نسبت به قوی‌ترین دستگاه نشان داده شده. ردیف‌های «ضعیف‌ترین دستگاه» باید بررسی شوند.",
    "وضعیت_پایانه‌ها": "وضعیت ماهانه هر پایانه. «فعال» یعنی در ماه آخر تراکنش داشته؛ «غایب» یعنی در ماه آخر نیست و احتمال خرابی یا مشکل دارد.",
    "داده_POS": "داده کامل مشتریان POS که نمای مشتری از آن استفاده می‌کند. بهتر است این شیت را تغییر ندهید.",
    "فهرست_تماس_درگاه": "فهرست تماس برای درگاه‌های پرداخت (IPG) که جدا از POS تحلیل شده‌اند. ستون «علت عدم کار» را بعد از تماس پر کنید.",
    "برتر_ضعیف_درگاه": "حساب‌های درگاه با بیشترین مبلغ (برتر) و حساب‌های دارای افت یا قطع (ضعیف).",
    "روند_درگاه": "روند تراکنش‌ها، مبلغ و معدل هر حساب درگاه در ماه‌های موجود.",
    "نمایش_حساب_درگاه": "کد حساب درگاه را از لیست کشویی سلول زرد انتخاب کنید تا اطلاعات و نمودار مبلغ و معدل همان حساب نشان داده شود.",
    "داده_درگاه": "داده کامل حساب‌های درگاه که نمای حساب از آن استفاده می‌کند. بهتر است این شیت را تغییر ندهید.",
    "تنظیمات": "آستانه‌ها و تنظیمات اصلی تحلیل، مثلاً حداقل معدل برای صبا، درصد افت و تبدیل ریال به تومان. مبالغ گزارش به میلیون تومان است.",
}
NOTE_FONT = Font(name="Calibri", size=12, italic=True, color="404040")


def add_note(ws, text):
    c = ws.cell(row=2, column=1, value="راهنما: " + text)
    c.font = NOTE_FONT
    c.alignment = Alignment(vertical="center")
    ws.row_dimensions[2].height = 28


def sheet(wb, name):
    ws = new_sheet(wb, name)
    add_note(ws, NOTES.get(name[:31], ""))
    return ws


def write_groups_sheet(wb, name, title, groups, labels):
    ws = sheet(wb, name)
    ws["A1"] = title
    ws["A1"].font = TITLE_FONT
    r = 4
    for key, label in labels:
        r = write_table(ws, groups[key], r, title=label) + 3
    return ws


def write_report(path, res_pos, res_ipg, cfg):
    wb = Workbook()
    wb.remove(wb.active)
    add_dashboard(wb, res_pos, res_ipg, cfg)

    sp = res_pos["summary"]
    pos_data = attach_calls(res_pos["customers"], res_pos["calls"], "کد مشتری")
    pos_info = ["کد شعبه", "صاحب پایانه", "نام فروشگاه", "موبایل", "وضعیت روند", "تغییر مبلغ (%)",
                "تغییر تراکنش (%)", "آخرین دوره فعال", "نمره ریسک (۰-۱۰۰)", "موج تماس",
                "امتیاز اولویت (۰-۱۰۰)", "علت‌های احتمالی", "اقدام پیشنهادی"]
    build_viewer(wb, "نمایش_مشتری_POS", "نمای مشتری - دستگاه‌های POS", pos_data, "داده_POS",
                 sp["periods"], pos_info, "کد مشتری")

    write_table(sheet(wb, "فهرست_تماس_POS"), res_pos["calls"], 4,
                title="فهرست تماس POS (اولویت‌دار، با کد شعبه)")
    write_groups_sheet(wb, "برتر_ضعیف_پتانسیل_POS", "مشتریان برتر، ضعیف و دارای پتانسیل (POS)",
                       res_pos["groups"], [("برتر", "مشتریان برتر POS"),
                                           ("ضعیف", "مشتریان ضعیف POS"),
                                           ("پتانسیل", "مشتریان دارای پتانسیل POS")])
    write_table(sheet(wb, "روند_POS"), res_pos["customers"], 4, title="روند مشتریان POS")
    write_table(sheet(wb, "ریزش_POS"), res_pos["churn"], 4, title="مشتریان ریزش‌کرده POS")
    write_table(sheet(wb, "کاندید_صبا"), res_pos["saba"], 4,
                title="کاندیدهای وام صبا (معدل بالا + افزایشی + افت مبلغ)")
    write_table(sheet(wb, "پتانسیل_POS"), res_pos["potential"], 4, title="مشتریان دارای پتانسیل POS")
    write_table(sheet(wb, "چند_پوز"), res_pos["multi"], 4,
                title=f"مشتریان چندپوز و ضعیف‌ترین دستگاه ({sp['D']})")
    write_table(sheet(wb, "وضعیت_پایانه‌ها"), res_pos["terminals"], 4, title="وضعیت ماهانه هر پایانه POS")
    if res_ipg:
        ip = res_ipg["summary"]["periods"]
        ipg_data = attach_calls(res_ipg["accounts"], res_ipg["calls"], "کد حساب")
        ipg_info = ["کد شعبه", "دارنده حساب", "نام فروشگاه", "وضعیت روند", "تغییر مبلغ (%)",
                    "آخرین دوره فعال", "موج تماس", "امتیاز اولویت (۰-۱۰۰)", "علت‌های احتمالی",
                    "اقدام پیشنهادی"]
        build_viewer(wb, "نمایش_حساب_درگاه", "نمای حساب - درگاه‌های پرداخت (IPG)", ipg_data,
                     "داده_درگاه", ip, ipg_info, "کد حساب")
        write_table(sheet(wb, "فهرست_تماس_درگاه"), res_ipg["calls"], 4,
                    title="فهرست تماس درگاه پرداخت (IPG، با کد شعبه)")
        write_groups_sheet(wb, "برتر_ضعیف_درگاه", "مشتریان برتر و ضعیف درگاه (IPG)",
                           res_ipg["groups"], [("برتر", "حساب‌های برتر درگاه"),
                                               ("ضعیف", "حساب‌های ضعیف درگاه")])
        write_table(sheet(wb, "روند_درگاه"), res_ipg["accounts"], 4, title="روند حساب‌های درگاه (IPG)")

    cfg_df = pd.DataFrame([(k, v) for k, v in cfg.items()], columns=["پارامتر", "مقدار"])
    write_table(sheet(wb, "تنظیمات"), cfg_df, 4, title="تنظیمات و آستانه‌ها")
    wb.calculation.fullCalcOnLoad = True
    wb.save(path)


# ------------------------------------------------------------------
def run(files, out):
    """اجرای کامل: POS با POS و IPG با IPG."""
    pos_list, ipg_list = [], []
    for f in files:
        kind, period, df = read_snapshot(f)
        if kind == "POS":
            pos_list.append(parse_pos(df, period))
        elif kind == "IPG":
            ipg_list.append(parse_ipg(df, period))
        else:
            print(f"[نادیده گرفته شد] {f}")

    if not pos_list:
        raise SystemExit("هیچ فایل POS پیدا نشد.")
    pos = pd.concat(pos_list, ignore_index=True)
    res_pos = analyse_pos(pos, CONFIG)
    res_ipg = None
    if ipg_list:
        res_ipg = analyse_ipg(pd.concat(ipg_list, ignore_index=True), CONFIG)
    write_report(out, res_pos, res_ipg, CONFIG)

    print("گزارش ساخته شد:", out)
    for label, val in res_pos["summary"]["rows"]:
        print(f"  POS - {label}: {val}")
    if res_ipg:
        for label, val in res_ipg["summary"]["rows"]:
            print(f"  IPG - {label}: {val}")


def main():
    ap = argparse.ArgumentParser(description="تحلیل پایش پایانه‌های فروش و درگاه‌ها")
    ap.add_argument("files", nargs="+", help="فایل‌های اکسل ماهانه (POS و/یا IPG)")
    ap.add_argument("-o", "--out", default="pos_monitoring_report.xlsx")
    args = ap.parse_args()
    run(args.files, args.out)


if __name__ == "__main__":
    main()
