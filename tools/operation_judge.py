"""室外機の稼働状況から、分析に使える日・時間帯を判定する。あわせて運転時間の推移をまとめる。

目的
    芋緑化ありの室外機（OACP-8・11）と、同じ区画の芋緑化なしの室外機（OACP-3）を
    同じ条件で比べるため、「2台とも運転している時間」だけを分析に使えるようにする。
    また、季節が進むにつれて室外機の運転時間がどう変わったかを確かめる。

入力
    実測データの Excel（シート「整理済みデータ」、1分値、ヘッダ9行）
    例：python3 tools/operation_judge.py 261013_実測データ_完全版.xlsx 出力.xlsx

処理
    1. 室外機ごとに「排気温度 − 吸気温度（3方位の平均）」を1分ごとに求め、
       10分の移動中央値をとる。1.0℃以上なら「運転中」とする。
       - 11-3 は北側吸気（BH列）を全期間除くため2方位の平均
       - 3-1 北側吸気（F列）は、熱電対の異常期間（下の BAD_3_1_NORTH）を除く
    2. 1時間ごとに、室外機の「判定できた分数」と「運転中の分数」を数える（Python）。
    3. 室外機の判定（運転／停止／断続）と、組（例：8-1 と 3-1）の判定（利用可など）は
       Excel の数式で行う。条件は「判定条件」シートの黄色のセルにあり、変えると判定がすべて更新される。
       - 組が「利用可」＝2台とも「運転」（1時間の80％以上運転中）、かつ平日
    4. 日ごとに、室外機の運転時間・運転の始まりと終わり・日中の平均気温をまとめ（Python）、
       期間ごとの平均とグラフを作る。

出力
    Excel（判定条件、日ごとの判定、時間ごとの判定、時間帯の一覧（列8・列11）、時刻別の利用可時間数、
    運転時間の推移、期間別のまとめ、グラフ（平日）、データの不具合、しきい値の確認）
"""
import sys
import datetime as dt

import numpy as np
import openpyxl
import pandas as pd
from openpyxl.chart import LineChart, Reference, ScatterChart, Series
from openpyxl.formatting.rule import CellIsRule, ColorScaleRule
from openpyxl.styles import Alignment, Border, Font, PatternFill, Side
from openpyxl.utils import get_column_letter as L
from openpyxl.worksheet.datavalidation import DataValidation

# ---------------------------------------------------------------- 設定
# 室外機ごとの列記号：(排気温度, [吸気温度の列])
UNITS = {
    "3-1": ("D", ["E", "F", "G"]), "3-2": ("J", ["K", "L", "M"]), "3-3": ("Q", ["R", "S", "T"]),
    "8-1": ("W", ["X", "Y", "Z"]), "8-2": ("AC", ["AD", "AE", "AF"]), "8-3": ("AK", ["AL", "AM", "AN"]),
    "11-1": ("AR", ["AS", "AT", "AU"]), "11-2": ("AX", ["AY", "AZ", "BA"]), "11-3": ("BF", ["BG", "BI"]),
}
PAIRS = [(f"{c}-{k}", f"3-{k}") for c in ("8", "11") for k in "123"]
COLS = {"列3": ["3-1", "3-2", "3-3"], "列8": ["8-1", "8-2", "8-3"], "列11": ["11-1", "11-2", "11-3"]}
DT_THRESHOLD = 1.0          # 運転中とみなす「排気−吸気」［℃］
ROLL_MIN = 10               # 移動中央値の幅［分］
AIR_TEMP = "BV"             # 屋上の気温（マルチ環境計測計）
# 3-1 北側吸気の熱電対の異常（夜間に周りより +3〜+10℃高い。急に始まり急に戻る）
BAD_3_1_NORTH = [("2026-09-14 18:47", "2026-09-22 11:40"), ("2026-09-25 23:20", "2026-10-01 11:13")]
HOLIDAYS = {dt.date(2026, 8, 11): "山の日", dt.date(2026, 9, 21): "敬老の日",
            dt.date(2026, 9, 22): "国民の休日", dt.date(2026, 9, 23): "秋分の日",
            dt.date(2026, 10, 12): "スポーツの日"}
FIRST_DAY, LAST_DAY = dt.date(2026, 7, 29), dt.date(2026, 10, 13)
PERIODS = [("8月", dt.date(2026, 7, 30), dt.date(2026, 8, 31)),
           ("9月前半", dt.date(2026, 9, 1), dt.date(2026, 9, 15)),
           ("9月後半", dt.date(2026, 9, 16), dt.date(2026, 9, 30)),
           ("10月", dt.date(2026, 10, 1), dt.date(2026, 10, 13))]
WEEK = "月火水木金土日"


def kind(d):
    if d in HOLIDAYS:
        return "祝日"
    return "土日" if d.weekday() >= 5 else "平日"


# ---------------------------------------------------------------- 計算（Python）
def load(path):
    wb = openpyxl.load_workbook(path, read_only=True)
    ws = wb["整理済みデータ"]
    rows = list(ws.iter_rows(min_row=10, values_only=True))
    df = pd.DataFrame(rows, columns=[L(i + 1) for i in range(len(rows[0]))])
    df["ts"] = pd.to_datetime(df["B"].dt.strftime("%Y-%m-%d") + " " + df["C"].astype(str))
    df = df.set_index("ts")
    for c in df.columns[3:]:
        df[c] = pd.to_numeric(df[c], errors="coerce")
    for a, b in BAD_3_1_NORTH:
        df.loc[a:b, "F"] = np.nan
    return df


def diff(df):
    """室外機ごとの「排気−吸気」の10分移動中央値（1分ごと）。"""
    return pd.DataFrame({u: (df[ex] - df[ins].mean(axis=1)).rolling(ROLL_MIN, center=True, min_periods=6).median()
                         for u, (ex, ins) in UNITS.items()})


def hourly_counts(dT, threshold=DT_THRESHOLD):
    hours = pd.date_range(f"{FIRST_DAY} 00:00", f"{LAST_DAY} 23:00", freq="1h")
    out = pd.DataFrame(index=hours)
    for u in UNITS:
        out[(u, "有効")] = dT[u].notna().resample("1h").sum().reindex(hours, fill_value=0)
        out[(u, "運転")] = (dT[u] >= threshold).resample("1h").sum().reindex(hours, fill_value=0)
    return out.astype(int)


def daily_operation(df, dT):
    """日ごとの運転時間［時間］、運転の始まり・終わり［時］、日中の平均気温［℃］。"""
    on = dT >= DT_THRESHOLD
    days = pd.date_range(str(FIRST_DAY), str(LAST_DAY))
    out = pd.DataFrame(index=days)
    air = df[AIR_TEMP]
    a = air[(air.index.hour >= 9) & (air.index.hour < 18)]
    tm = a.groupby(a.index.normalize()).agg(["mean", "count"])
    out["気温"] = tm["mean"].where(tm["count"] >= 0.8 * 540)
    for u in UNITS:
        v = dT[u].notna().groupby(dT.index.normalize()).sum()
        h = on[u].groupby(on.index.normalize()).sum() / 60
        out[(u, "時間")] = h.where(v >= 22 * 60)            # ほぼ丸1日（22時間以上）判定できた日だけ
        st, en = {}, {}
        for d, s in on[u].groupby(on.index.normalize()):
            night = s[(s.index.hour >= 3) & (s.index.hour < 4)]
            if v.get(d, 0) < 22 * 60 or night.mean() > 0.2:   # 夜も運転している日（24時間運転）は空欄
                continue
            day = s[s.index.hour >= 5]
            if day.any():
                t0, t1 = day[day].index[0], day[day].index[-1]
                st[d] = t0.hour + t0.minute / 60
                en[d] = t1.hour + t1.minute / 60
        out[(u, "始まり")] = pd.Series(st)
        out[(u, "終わり")] = pd.Series(en)
    for c, us in COLS.items():
        out[(c, "時間")] = out[[(u, "時間") for u in us]].mean(axis=1)
        out[(c, "始まり")] = out[[(u, "始まり") for u in us]].mean(axis=1)
        out[(c, "終わり")] = out[[(u, "終わり") for u in us]].mean(axis=1)
    return out


def sensitivity(dT):
    """①のしきい値を変えたときの、平日日中の利用可の時間数（組ごと）。"""
    rows, base = [], None
    for th in (0.5, 1.0, 1.5, 2.0):
        c = hourly_counts(dT, th)
        sel = ((c.index.hour >= 9) & (c.index.hour <= 17) & (c.index >= "2026-07-30")
               & np.array([kind(t.date()) == "平日" for t in c.index]))
        cls = {}
        for u in UNITS:
            v, n = c[(u, "有効")], c[(u, "運転")]
            cls[u] = np.where(v < 48, "x", np.where(n >= 0.8 * v, "運転", np.where(n <= 0.2 * v, "停止", "断続")))
        cls = pd.DataFrame(cls, index=c.index)[sel]
        vals = [int(((cls[a] == "運転") & (cls[b] == "運転")).sum()) for a, b in PAIRS]
        if th == DT_THRESHOLD:
            base = cls
        rows.append((th, vals, cls))
    m = base != "x"
    return [(th, vals, float(((cls != base) & m).sum().sum() / m.sum().sum())) for th, vals, cls in rows]


# ---------------------------------------------------------------- Excel 書式
FONT = "游ゴシック"
NAVY, TEXT, MUTED = "22354F", "172033", "5E6B7A"
F_BASE = Font(name=FONT, size=10, color=TEXT)
F_BOLD = Font(name=FONT, size=10, color=TEXT, bold=True)
F_HEAD = Font(name=FONT, size=10, color="FFFFFF", bold=True)
F_TITLE = Font(name=FONT, size=14, color=NAVY, bold=True)
F_NOTE = Font(name=FONT, size=9, color=MUTED)
F_INPUT = Font(name=FONT, size=10, color="0000FF", bold=True)
F_SMALL = Font(name=FONT, size=9, color=TEXT)
FILL_HEAD = PatternFill("solid", fgColor=NAVY)
FILL_INPUT = PatternFill("solid", fgColor="FFF2B3")
FILL_SUB = PatternFill("solid", fgColor="EEF2F6")
THIN = Side(style="thin", color="C9D1DB")
BOX = Border(left=THIN, right=THIN, top=THIN, bottom=THIN)
CENTER = Alignment(horizontal="center", vertical="center", wrap_text=True)
LEFT = Alignment(horizontal="left", vertical="center", wrap_text=True)
# 判定の色（利用可・運転＝緑、停止＝灰、そろわない・断続＝橙、土日祝＝薄い紫、判定不能＝白地に灰文字）
COLORS = {"利用可": "B9E0AE", "運転": "B9E0AE", "両方停止": "D3D8DD", "停止": "D3D8DD",
          "運転がそろわない": "F8CBAD", "断続": "F8CBAD", "土日祝": "E4DCF0", "判定不能": "FFFFFF", "不可": "F8CBAD"}
SERIES_COLOR = {"列3": "5E6B7A", "列8": "01B8AA", "列11": "118DFF"}


def head(ws, row, col, text, width=None):
    c = ws.cell(row=row, column=col, value=text)
    c.font, c.fill, c.alignment, c.border = F_HEAD, FILL_HEAD, CENTER, BOX
    if width:
        ws.column_dimensions[L(col)].width = width
    return c


def cell(ws, row, col, value, font=F_BASE, align=CENTER, fmt=None, fill=None):
    c = ws.cell(row=row, column=col, value=value)
    c.font, c.alignment, c.border = font, align, BOX
    if fmt:
        c.number_format = fmt
    if fill:
        c.fill = fill
    return c


def color_rules(ws, rng, labels):
    for lab in labels:
        ws.conditional_formatting.add(
            rng, CellIsRule(operator="equal", formula=[f'"{lab}"'],
                            fill=PatternFill("solid", fgColor=COLORS[lab]),
                            font=Font(name=FONT, size=9, color=MUTED if lab == "判定不能" else TEXT)))


def nz(v):
    return None if v is None or (isinstance(v, float) and np.isnan(v)) else v


# ---------------------------------------------------------------- シートづくり
def build(counts, daily, sens, path):
    wb = openpyxl.Workbook()
    ws0 = wb.active
    ws0.title = "判定条件"
    wsd = wb.create_sheet("日ごとの判定")
    wsh = wb.create_sheet("時間ごとの判定")
    wsg = {"8": wb.create_sheet("時間帯の一覧_列8"), "11": wb.create_sheet("時間帯の一覧_列11")}
    wst = wb.create_sheet("時刻別の利用可時間数")
    wso = wb.create_sheet("運転時間の推移")
    wsp = wb.create_sheet("期間別のまとめ")
    wsc = wb.create_sheet("グラフ（平日）")
    wsb = wb.create_sheet("データの不具合")
    wss = wb.create_sheet("しきい値の確認")

    # ---- 判定条件
    ws = ws0
    ws.column_dimensions["A"].width = 3
    ws.column_dimensions["B"].width = 46
    ws.column_dimensions["C"].width = 12
    ws.column_dimensions["D"].width = 74
    ws["B1"] = "室外機の稼働状況による、分析に使える日・時間帯の判定"
    ws["B1"].font = F_TITLE
    ws["B2"] = ("対象：2026/7/29〜10/13（1分値）。芋緑化ありの室外機と、同じ区画の芋緑化なし（OACP-3）の室外機を"
                "1組とし、平日に2台とも運転している時間だけを「利用可」とする。黄色のセルを変えると、すべての判定が更新される。")
    ws["B2"].font, ws["B2"].alignment = F_NOTE, LEFT
    ws.merge_cells("B2:D2")
    ws.row_dimensions[2].height = 30

    r = 4
    for j, t in enumerate(["段階", "値", "内容"]):
        head(ws, r, 2 + j, t)
    r += 1
    cell(ws, r, 2, "① 1分ごとの運転判定（Pythonで計算）", F_BOLD, LEFT)
    cell(ws, r, 3, f"{DT_THRESHOLD}℃以上", F_BOLD)
    cell(ws, r, 4, "排気温度 − 吸気温度（3方位の平均）の10分の移動中央値がこの値以上なら「運転中」。"
                   "停止中は約0℃、運転中は2〜10℃に分かれ、その谷が1℃付近にある。11-3は北側を除く2方位。"
                   "3-1北側吸気は異常期間を除く（「データの不具合」）", F_BASE, LEFT)
    ws.row_dimensions[r].height = 45
    r += 1
    inputs = [
        ("valid", "② 1時間の判定に必要な有効データ［分］", 48, "0",
         "1時間のうち、①の判定ができた分がこれ未満なら「判定不能」（これまでの1時間平均の基準と同じ）"),
        ("run", "③ 室外機が「運転」とみなす割合", 0.8, "0%",
         "有効な分のうち、運転中の分がこの割合以上なら「運転」"),
        ("stop", "④ 室外機が「停止」とみなす割合", 0.2, "0%",
         "運転中の分がこの割合以下なら「停止」。③と④の間は「断続」（1時間の中で運転と停止を繰り返す）"),
        ("holiday", "⑤ 土日祝の扱い", "除く", "@",
         "「除く」なら土日・祝日はすべて使わない（これまでの勤務日の扱いと同じ）。「含める」なら運転していれば使う"),
        ("start", "⑥ 日中の始まり［時台］", 9, "0", "日ごとの判定で数える時間帯の始まり（これまでの「日中」と同じ9時台）"),
        ("end", "⑦ 日中の終わり［時台］", 17, "0", "日ごとの判定で数える時間帯の終わり（17時台＝18時まで）"),
        ("days", "⑧ 日として使うのに必要な利用可の時間数", 7, "0",
         "⑥〜⑦の9時間のうち、利用可がこの時間数以上ある日を「利用可」の日とする（これまでの「9時間中7時間以上」と同じ）"),
    ]
    ref = {}
    for key, a, v, fmt, c in inputs:
        cell(ws, r, 2, a, F_BOLD, LEFT)
        cell(ws, r, 3, v, F_INPUT, CENTER, fmt, FILL_INPUT)
        cell(ws, r, 4, c, F_BASE, LEFT)
        ws.row_dimensions[r].height = 32
        ref[key] = f"判定条件!$C${r}"
        if key == "holiday":
            dv = DataValidation(type="list", formula1='"除く,含める"', allow_blank=False)
            ws.add_data_validation(dv)
            dv.add(f"C{r}")
        r += 1
    rules = [("組が「利用可」となる条件", "2台とも③の「運転」（1時間の80％以上運転中）の時間。1時間平均は運転状態がそろった値になる"),
             ("列8・列11の値の出し方", "区画ごと（組ごと）に分析し、列の値は3区画の結果を平均する。3区画が同じ時間にそろう必要はない")]
    for a, c in rules:
        cell(ws, r, 2, a, F_BOLD, LEFT)
        cell(ws, r, 3, "", F_BASE)
        cell(ws, r, 4, c, F_BASE, LEFT)
        ws.row_dimensions[r].height = 30
        r += 1

    r += 1
    ws.cell(row=r, column=2, value="判定の名前").font = F_BOLD
    r += 1
    legend = [("利用可", "平日に2台とも「運転」。分析に使う"),
              ("両方停止", "2台とも停止。運転中の比較には使わない（夜間など）"),
              ("運転がそろわない", "片方だけ運転、またはどちらかが断続。使わない"),
              ("土日祝", "⑤が「除く」のとき、土日・祝日の時間。使わない"),
              ("判定不能", "排気温度・吸気温度が欠けていて運転状態がわからない。使わない")]
    for lab, txt in legend:
        cell(ws, r, 2, lab, F_BOLD, CENTER, fill=PatternFill("solid", fgColor=COLORS[lab]))
        cell(ws, r, 3, "", F_BASE)
        cell(ws, r, 4, txt, F_BASE, LEFT)
        r += 1
    r += 1
    ws.cell(row=r, column=2, value="結果のまとめ（7/30〜10/13、日中）").font = F_BOLD
    r += 1
    head(ws, r, 2, "組")
    head(ws, r, 3, "利用可の時間数")
    head(ws, r, 4, "利用可の日数（⑧を満たす日）")
    sum_row = r + 1
    r += 1
    for a, b in PAIRS:
        cell(ws, r, 2, f"{a}と{b}", F_BOLD, LEFT)
        r += 1
    ws.cell(row=r, column=2, value="7/29は設置当日のため数えない").font = F_NOTE

    # ---- 時間ごとの判定
    ws = wsh
    hrs = counts.index
    R0, RN = 3, 3 + len(hrs) - 1
    for j, (t, w) in enumerate([("日付", 11), ("時台", 6), ("曜日", 6), ("区分", 8)]):
        head(ws, 1, 1 + j, t, w)
        ws.merge_cells(start_row=1, start_column=1 + j, end_row=2, end_column=1 + j)
    col = 5
    ucol = {}
    for u in UNITS:
        ws.merge_cells(start_row=1, start_column=col, end_row=1, end_column=col + 2)
        head(ws, 1, col, f"室外機 {u}")
        for j, t in enumerate(["有効［分］", "運転中［分］", "判定"]):
            head(ws, 2, col + j, t, 7 if j < 2 else 8)
        ucol[u] = col
        col += 3
    ws.merge_cells(start_row=1, start_column=col, end_row=1, end_column=col + len(PAIRS) - 1)
    head(ws, 1, col, "組の判定")
    pcol = {}
    for a, b in PAIRS:
        head(ws, 2, col, f"{a}と{b}", 14)
        pcol[(a, b)] = col
        col += 1
    ncol = {}
    for c in ("8", "11"):
        ws.merge_cells(start_row=1, start_column=col, end_row=2, end_column=col)
        head(ws, 1, col, f"列{c}：利用可の区画数", 10)
        ncol[c] = col
        col += 1
    ws.row_dimensions[2].height = 42
    for i, t in enumerate(hrs):
        rr = R0 + i
        d = t.date()
        ws.cell(row=rr, column=1, value=d).number_format = "m/d"
        ws.cell(row=rr, column=2, value=t.hour)
        ws.cell(row=rr, column=3, value=WEEK[d.weekday()])
        ws.cell(row=rr, column=4, value=kind(d))
        for u in UNITS:
            c0 = ucol[u]
            ws.cell(row=rr, column=c0, value=int(counts.at[t, (u, "有効")]))
            ws.cell(row=rr, column=c0 + 1, value=int(counts.at[t, (u, "運転")]))
            V, N = f"{L(c0)}{rr}", f"{L(c0 + 1)}{rr}"
            ws.cell(row=rr, column=c0 + 2,
                    value=f'=IF({V}<{ref["valid"]},"判定不能",IF({N}>={ref["run"]}*{V},"運転",'
                          f'IF({N}<={ref["stop"]}*{V},"停止","断続")))')
        for a, b in PAIRS:
            JA, JB = f"{L(ucol[a] + 2)}{rr}", f"{L(ucol[b] + 2)}{rr}"
            ws.cell(row=rr, column=pcol[(a, b)],
                    value=f'=IF(AND({ref["holiday"]}="除く",$D{rr}<>"平日"),"土日祝",'
                          f'IF(OR({JA}="判定不能",{JB}="判定不能"),"判定不能",'
                          f'IF(AND({JA}="運転",{JB}="運転"),"利用可",'
                          f'IF(AND({JA}="停止",{JB}="停止"),"両方停止","運転がそろわない"))))')
        for c in ("8", "11"):
            p0 = pcol[(f"{c}-1", "3-1")]
            ws.cell(row=rr, column=ncol[c], value=f'=COUNTIF({L(p0)}{rr}:{L(p0 + 2)}{rr},"利用可")')
    for row in ws.iter_rows(min_row=R0, max_row=RN, max_col=col - 1):
        for c in row:
            c.font, c.alignment = F_SMALL, CENTER
    last = L(col - 1)
    color_rules(ws, f"E{R0}:{last}{RN}", ["利用可", "運転", "両方停止", "停止", "運転がそろわない", "断続", "土日祝", "判定不能"])
    ws.freeze_panes = "E3"
    ws.auto_filter.ref = f"A2:{last}{RN}"
    H = "時間ごとの判定"
    rngA, rngB, rngD = f"{H}!$A${R0}:$A${RN}", f"{H}!$B${R0}:$B${RN}", f"{H}!$D${R0}:$D${RN}"

    def pr(a, b):
        c = L(pcol[(a, b)])
        return f"{H}!${c}${R0}:${c}${RN}"

    # ---- 日ごとの判定
    ws = wsd
    days = pd.date_range(str(FIRST_DAY), str(LAST_DAY)).date
    for j, (t, w) in enumerate([("日付", 9), ("曜日", 5), ("区分", 10)]):
        head(ws, 1, 1 + j, t, w)
        ws.merge_cells(start_row=1, start_column=1 + j, end_row=2, end_column=1 + j)
    col = 4
    dcol = {}
    for a, b in PAIRS:
        ws.merge_cells(start_row=1, start_column=col, end_row=1, end_column=col + 1)
        head(ws, 1, col, f"{a}と{b}")
        head(ws, 2, col, "日中の利用可［時間］", 9)
        head(ws, 2, col + 1, "日の判定", 8)
        dcol[(a, b)] = col
        col += 2
    dncol = {}
    for c in ("8", "11"):
        ws.merge_cells(start_row=1, start_column=col, end_row=2, end_column=col)
        head(ws, 1, col, f"列{c}：利用可の区画数", 10)
        dncol[c] = col
        col += 1
    ws.merge_cells(start_row=1, start_column=col, end_row=2, end_column=col)
    head(ws, 1, col, "備考（データの欠け）", 46)
    note_col = col
    ws.row_dimensions[2].height = 42
    D0 = 3
    DN = D0 + len(days) - 1
    for i, d in enumerate(days):
        rr = D0 + i
        ws.cell(row=rr, column=1, value=d).number_format = "m/d"
        ws.cell(row=rr, column=2, value=WEEK[d.weekday()])
        ws.cell(row=rr, column=3, value=HOLIDAYS.get(d, kind(d)))
        for a, b in PAIRS:
            c0 = dcol[(a, b)]
            ws.cell(row=rr, column=c0,
                    value=f'=COUNTIFS({rngA},$A{rr},{rngB},">="&{ref["start"]},{rngB},"<="&{ref["end"]},{pr(a, b)},"利用可")')
            ws.cell(row=rr, column=c0 + 1, value=f'=IF({L(c0)}{rr}>={ref["days"]},"利用可","不可")')
        for c in ("8", "11"):
            js = [f'{L(dcol[(f"{c}-{k}", f"3-{k}")] + 1)}{rr}' for k in "123"]
            ws.cell(row=rr, column=dncol[c], value=f'=COUNTIF({js[0]},"利用可")+COUNTIF({js[1]},"利用可")+COUNTIF({js[2]},"利用可")')
        ws.cell(row=rr, column=note_col, value=day_note(counts, d))
    for row in ws.iter_rows(min_row=D0, max_row=DN, max_col=note_col):
        for c in row:
            c.font, c.alignment, c.border = F_SMALL, CENTER, BOX
    for rr in range(D0, DN + 1):
        ws.cell(row=rr, column=note_col).alignment = LEFT
        if ws.cell(row=rr, column=3).value != "平日":
            for cc in (1, 2, 3):
                ws.cell(row=rr, column=cc).fill = FILL_SUB
    color_rules(ws, f"D{D0}:{L(note_col - 1)}{DN}", ["利用可", "不可"])
    ws.freeze_panes = "D3"

    # 判定条件のまとめ（数式）
    rr = sum_row
    for a, b in PAIRS:
        cell(ws0, rr, 3, f'=COUNTIFS({rngA},">="&DATE(2026,7,30),{rngB},">="&{ref["start"]},{rngB},"<="&{ref["end"]},{pr(a, b)},"利用可")')
        c1 = L(dcol[(a, b)] + 1)
        cell(ws0, rr, 4, f'=COUNTIFS(日ごとの判定!${c1}${D0 + 1}:${c1}${DN},"利用可")')
        rr += 1

    # ---- 時間帯の一覧（日付×時刻）
    sym = '{"利用可","両方停止","運転がそろわない","判定不能","土日祝"}'
    for c, wg in wsg.items():
        wg["A1"] = f"列{c}：組ごとの判定（○＝利用可、停＝両方停止、×＝運転がそろわない、－＝判定不能、休＝土日祝）"
        wg["A1"].font = F_BOLD
        wg.column_dimensions["A"].width = 9
        wg.column_dimensions["B"].width = 4
        col = 3
        for k in "123":
            a, b = f"{c}-{k}", f"3-{k}"
            wg.merge_cells(start_row=2, start_column=col, end_row=2, end_column=col + 23)
            head(wg, 2, col, f"{a}と{b}")
            src = L(pcol[(a, b)])
            for h in range(24):
                head(wg, 3, col + h, h, 2.6).font = Font(name=FONT, size=7, color="FFFFFF", bold=True)
                for i in range(len(days)):
                    wg.cell(row=4 + i, column=col + h,
                            value=f'=IFERROR(CHOOSE(MATCH(INDEX({H}!${src}${R0}:${src}${RN},{i * 24 + h + 1}),'
                                  f'{sym},0),"○","停","×","－","休"),"")')
            col += 25
            wg.column_dimensions[L(col - 1)].width = 1.5
        for i, d in enumerate(days):
            x = wg.cell(row=4 + i, column=1, value=d)
            x.number_format, x.font = "m/d", Font(name=FONT, size=8, color=TEXT)
            y = wg.cell(row=4 + i, column=2, value=WEEK[d.weekday()])
            y.font = Font(name=FONT, size=8, color="E66C64" if kind(d) != "平日" else TEXT)
        rng = f"C4:{L(col - 2)}{3 + len(days)}"
        for row in wg.iter_rows(min_row=4, max_row=3 + len(days), min_col=3, max_col=col - 2):
            for x in row:
                x.font, x.alignment = Font(name=FONT, size=7, color=TEXT), CENTER
        for lab, s in (("利用可", "○"), ("両方停止", "停"), ("運転がそろわない", "×"), ("土日祝", "休")):
            wg.conditional_formatting.add(rng, CellIsRule(operator="equal", formula=[f'"{s}"'],
                                                          fill=PatternFill("solid", fgColor=COLORS[lab])))
        wg.freeze_panes = "C4"

    # ---- 時刻別の利用可時間数
    ws = wst
    ws["A1"] = "時刻ごとの利用可の時間数（7/30〜10/13の合計）。どの時間帯が分析に使えるかを見る"
    ws["A1"].font = F_BOLD
    head(ws, 2, 1, "組", 14)
    for h in range(24):
        head(ws, 2, 2 + h, f"{h}時台", 6)
    for j, (a, b) in enumerate(PAIRS):
        rr = 3 + j
        cell(ws, rr, 1, f"{a}と{b}", F_BOLD)
        for h in range(24):
            cell(ws, rr, 2 + h, f'=COUNTIFS({rngA},">="&DATE(2026,7,30),{rngB},{h},{pr(a, b)},"利用可")')
    ws.conditional_formatting.add(f"B3:Y{2 + len(PAIRS)}",
                                  ColorScaleRule(start_type="num", start_value=0, start_color="FFFFFF",
                                                 end_type="max", end_color="5FAF52"))

    # ---- 運転時間の推移（Pythonで計算した日ごとの値）
    ws = wso
    ws["A1"] = ("日ごとの運転時間（1日のうち「運転中」だった時間の合計）と、運転の始まり・終わり。Pythonで1分値から計算した値。"
                "22時間以上判定できた日だけ。始まり・終わりは、夜も運転している日（24時間運転）は空欄")
    ws["A1"].font, ws["A1"].alignment = F_NOTE, LEFT
    ws.merge_cells("A1:Z1")
    ws.row_dimensions[1].height = 28
    names = list(UNITS) + list(COLS)
    for j, (t, w) in enumerate([("日付", 8), ("曜日", 5), ("区分", 9), ("日中の平均気温［℃］", 9)]):
        head(ws, 2, 1 + j, t, w)
        ws.merge_cells(start_row=2, start_column=1 + j, end_row=3, end_column=1 + j)
    blocks = [("1日の運転時間［時間］", "時間", "0.0"), ("運転の始まり［時刻］", "始まり", "h:mm"), ("運転の終わり［時刻］", "終わり", "h:mm")]
    col = 5
    ocol = {}
    for title, key, fmt in blocks:
        ws.merge_cells(start_row=2, start_column=col, end_row=2, end_column=col + len(names) - 1)
        head(ws, 2, col, title)
        for j, n in enumerate(names):
            head(ws, 3, col + j, n + ("平均" if n in COLS else ""), 6.5)
            ocol[(key, n)] = col + j
        col += len(names)
    O0 = 4
    ON = O0 + len(daily) - 1
    for i, (d, row) in enumerate(daily.iterrows()):
        rr = O0 + i
        dd = d.date()
        ws.cell(row=rr, column=1, value=dd).number_format = "m/d"
        ws.cell(row=rr, column=2, value=WEEK[dd.weekday()])
        ws.cell(row=rr, column=3, value=HOLIDAYS.get(dd, kind(dd)))
        x = ws.cell(row=rr, column=4, value=nz(row["気温"]))
        x.number_format = "0.0"
        for key, _, fmt in [(b[1], b[0], b[2]) for b in blocks]:
            for n in names:
                v = nz(row[(n, key)])
                if v is not None and key != "時間":
                    v = v / 24          # Excel の時刻（1日＝1）
                c = ws.cell(row=rr, column=ocol[(key, n)], value=v)
                c.number_format = fmt
    for row in ws.iter_rows(min_row=O0, max_row=ON, max_col=col - 1):
        for c in row:
            c.font, c.alignment, c.border = F_SMALL, CENTER, BOX
    for rr in range(O0, ON + 1):
        if ws.cell(row=rr, column=3).value != "平日":
            for cc in range(1, col):
                ws.cell(row=rr, column=cc).fill = FILL_SUB
    ws.freeze_panes = "E4"
    OS = "運転時間の推移"

    # ---- 期間別のまとめ（数式）
    ws = wsp
    ws["A1"] = "期間ごとの平日の平均。運転時間・始まり・終わりは「運転時間の推移」、利用可の割合は「時間ごとの判定」から数式で集計"
    ws["A1"].font = F_BOLD
    hdr = ["期間", "始め", "終わり", "平日の日数", "日中の平均気温［℃］"]
    hdr += [f"{c}：1日の運転時間［時間］" for c in COLS]
    hdr += [f"{c}：運転の始まり" for c in COLS] + [f"{c}：運転の終わり" for c in COLS]
    hdr += [f"{a}と{b}：利用可の割合" for a, b in PAIRS]
    for j, t in enumerate(hdr):
        head(ws, 2, 1 + j, t, 10 if j > 2 else 9)
    ws.row_dimensions[2].height = 58
    od = f"{OS}!$A${O0}:$A${ON}"
    ok = f"{OS}!$C${O0}:$C${ON}"
    for i, (name, a, b) in enumerate(PERIODS):
        rr = 3 + i
        cell(ws, rr, 1, name, F_BOLD)
        cell(ws, rr, 2, a, fmt="m/d")
        cell(ws, rr, 3, b, fmt="m/d")
        cond = f'{od},">="&$B{rr},{od},"<="&$C{rr},{ok},"平日"'
        cell(ws, rr, 4, f"=COUNTIFS({cond})")
        cell(ws, rr, 5, f"=IFERROR(AVERAGEIFS({OS}!$D${O0}:$D${ON},{cond}),\"－\")", fmt="0.0")
        j = 6
        for key, fmt in (("時間", "0.0"), ("始まり", "h:mm"), ("終わり", "h:mm")):
            for c in COLS:
                cl = L(ocol[(key, c)])
                cell(ws, rr, j, f'=IFERROR(AVERAGEIFS({OS}!${cl}${O0}:${cl}${ON},{cond}),"－")', fmt=fmt)
                j += 1
        for a, b in PAIRS:
            hc = f'{rngA},">="&$B{rr},{rngA},"<="&$C{rr},{rngD},"平日",{rngB},">="&{ref["start"]},{rngB},"<="&{ref["end"]},{pr(a, b)}'
            cell(ws, rr, j, f'=IFERROR(COUNTIFS({hc},"利用可")/COUNTIFS({hc},"<>判定不能"),"データなし")', fmt="0%")
            j += 1
    ws.cell(row=8, column=1, value=("利用可の割合＝平日の日中（⑥〜⑦）のうち、判定できた時間に対する「利用可」の時間の割合。"
                                   "列11は9/15〜10/1のデータがない。列3の運転時間は3-3の24時間運転（7/29〜8/14、9/2〜9/7）を含む")).font = F_NOTE
    ws.freeze_panes = "B3"

    # ---- グラフ（平日）：Pythonで計算した平日の値。欠けは空欄（グラフでは線が途切れる）
    ws = wsc
    ws["A1"] = "平日の1日の運転時間（3台の平均）と日中の平均気温。「運転時間の推移」の平日の行を写した値（欠けは空欄）"
    ws["A1"].font = F_BOLD
    for j, (t, w) in enumerate([("日付", 8), ("日中の平均気温［℃］", 10), ("列3", 8), ("列8", 8), ("列11", 8)]):
        head(ws, 2, 1 + j, t, w)
    wk = daily[[kind(d.date()) == "平日" for d in daily.index]]
    wk = wk.loc["2026-07-30":]
    for i, (d, row) in enumerate(wk.iterrows()):
        rr = 3 + i
        cell(ws, rr, 1, f"{d.month}/{d.day}", F_SMALL)
        cell(ws, rr, 2, nz(round(row["気温"], 1)) if not np.isnan(row["気温"]) else None, F_SMALL, fmt="0.0")
        for j, c in enumerate(COLS):
            v = row[(c, "時間")]
            cell(ws, rr, 3 + j, None if np.isnan(v) else round(float(v), 2), F_SMALL, fmt="0.0")
    G0, GN = 3, 3 + len(wk) - 1

    ch = LineChart()
    ch.title = "平日の1日の運転時間（3台の平均）と日中の平均気温"
    ch.y_axis.title = "運転時間［時間］"
    ch.y_axis.scaling.min, ch.y_axis.scaling.max = 0, 24
    ch.height, ch.width = 9, 26
    for j, c in enumerate(COLS):
        s = Series(Reference(ws, min_col=3 + j, min_row=G0, max_row=GN), title=c)
        s.graphicalProperties.line.solidFill = SERIES_COLOR[c]
        s.graphicalProperties.line.width = 22000
        s.smooth = False
        ch.series.append(s)
    ch.set_categories(Reference(ws, min_col=1, min_row=G0, max_row=GN))
    t2 = LineChart()
    s = Series(Reference(ws, min_col=2, min_row=G0, max_row=GN), title="日中の平均気温（右軸）")
    s.graphicalProperties.line.solidFill = "ED7D31"
    s.graphicalProperties.line.dashStyle = "dash"
    s.graphicalProperties.line.width = 15000
    s.smooth = False
    t2.series.append(s)
    t2.y_axis.axId = 200
    t2.y_axis.title = "気温［℃］"
    t2.y_axis.crosses = "max"
    t2.y_axis.scaling.min, t2.y_axis.scaling.max = 15, 35
    t2.y_axis.majorGridlines = None
    ch += t2
    ch.display_blanks = "gap"
    for ax in (ch.x_axis, ch.y_axis, t2.y_axis):
        ax.delete = False
    ch.legend.position = "b"
    ws.add_chart(ch, "G2")

    sc = ScatterChart()
    sc.title = "日中の平均気温と1日の運転時間（平日）"
    sc.style = 13
    sc.x_axis.title = "日中の平均気温［℃］"
    sc.y_axis.title = "運転時間［時間］"
    sc.x_axis.scaling.min, sc.x_axis.scaling.max = 18, 34
    sc.y_axis.scaling.min, sc.y_axis.scaling.max = 0, 24
    sc.height, sc.width = 9, 14
    xs = Reference(ws, min_col=2, min_row=G0, max_row=GN)
    for j, c in enumerate(COLS):
        s = Series(Reference(ws, min_col=3 + j, min_row=G0, max_row=GN), xs, title=c)
        s.marker.symbol = "circle"
        s.marker.size = 5
        s.marker.graphicalProperties.solidFill = SERIES_COLOR[c]
        s.marker.graphicalProperties.line.solidFill = SERIES_COLOR[c]
        s.graphicalProperties.line.noFill = True
        sc.series.append(s)
    for ax in (sc.x_axis, sc.y_axis):
        ax.delete = False
    sc.legend.position = "b"
    ws.add_chart(sc, "G21")

    # ---- データの不具合
    ws = wsb
    ws["A1"] = "運転の判定に関わるデータの欠け・異常"
    ws["A1"].font = F_BOLD
    for j, (t, w) in enumerate([("対象", 16), ("期間", 30), ("内容", 60), ("扱い", 40)]):
        head(ws, 2, 1 + j, t, w)
    issues = [
        ("3-1 北側吸気（F列）", "9/14 18:47〜9/22 11:40", "夜間（全台停止）に周りの吸気温度より＋3.4〜＋4.4℃高い。急に高くなり、急に戻る", "この期間は使わない（3-1は正面・南側の2方位で計算）。今回見つけた異常"),
        ("3-1 北側吸気（F列）", "9/25 23:20〜10/1 11:13", "同じく＋6.4〜＋9.9℃高い。10/1のデータ回収で戻っている", "同上"),
        ("8-2 排気温度（AC列）", "8/14 15:10〜8/28 10:25", "設置不良で欠損（これまでと同じ）", "8-2は判定不能。組8-2と3-2は使えない"),
        ("3-2 排気温度（J列）", "8/30〜9/7、9/14〜9/22ごろ", "センサー不良で欠損が多い（エラー値・起こり得ない値を空欄にしたもの）", "欠けた時間は判定不能。組8-2・11-2と3-2は使えない時間が多い"),
        ("OACP-11 の熱電対（AR〜BK列）", "9/14 17:23〜10/1 15:01", "元データがない", "列11はこの期間すべて判定不能"),
        ("8-2 の運転", "8/8〜8/14", "8-1・8-3は運転しているが、8-2は日中も停止（データはある）", "判定どおり「運転がそろわない」として使わない"),
        ("3-3 の運転", "7/29〜8/14、9/2〜9/7", "夜間も運転（24時間運転）", "夜間は8-3・11-3が停止しているため「運転がそろわない」"),
        ("排気が吸気より低い時間", "9月後半〜10月（主に3-1・3-3）", "停止中に排気が吸気より1〜3℃低い時間がある。雨の後に多く、熱電対がぬれた影響とみられる。暖房運転のような大きな低下は見られない", "「停止」として扱う（運転の判定には影響しない）"),
        ("ロガーの回収", "8/14・8/28・9/7・9/14・10/1・10/13", "回収作業で数分〜1時間ほど記録が止まる", "有効データが48分未満の時間は判定不能"),
        ("10/13", "14時台以降", "記録の終わり（3列14:16、8列13:43、11列14:01）", "判定不能"),
        ("（参考）屋上の気象", "9/14、10/13", "マルチ環境計測計のデータがない（運転の判定には使わない）", "天気・気温を使う分析ではこの2日間を除く"),
    ]
    for i, row in enumerate(issues):
        for j, v in enumerate(row):
            cell(ws, 3 + i, 1 + j, v, F_BASE, LEFT)
        ws.row_dimensions[3 + i].height = 32

    # ---- しきい値の確認（Pythonで計算した値）
    ws = wss
    ws["A1"] = "①のしきい値（排気−吸気）を変えたときの、平日日中（9〜17時台）の利用可の時間数（2台とも80％以上運転。Pythonで計算した値）"
    ws["A1"].font = F_BOLD
    head(ws, 2, 1, "しきい値［℃］", 12)
    for j, (a, b) in enumerate(PAIRS):
        head(ws, 2, 2 + j, f"{a}と{b}", 11)
    head(ws, 2, 2 + len(PAIRS), "室外機の判定が1.0℃のときと変わる割合", 18)
    ws.row_dimensions[2].height = 32
    for i, (th, vals, change) in enumerate(sens):
        cell(ws, 3 + i, 1, th, F_BOLD, fmt="0.0")
        for j, v in enumerate(vals):
            cell(ws, 3 + i, 2 + j, v)
        cell(ws, 3 + i, 2 + len(PAIRS), change, fmt="0.0%")
    ws.cell(row=4 + len(sens), column=1,
            value="1.0℃は、停止中（約0℃）と運転中（2〜10℃）の分布の谷にあたる。運転中も排気−吸気は0℃と5〜10℃の間を行き来するため、"
                  "しきい値を変えて判定が変わるのは主に運転の始まり・終わりと低負荷の時間").font = F_NOTE

    for w in wb.worksheets:
        w.sheet_view.zoomScale = 90
    wb.save(path)


def day_note(counts, d):
    """日中（9〜17時台）に判定不能が多い室外機を書く。"""
    t0 = pd.Timestamp(d) + pd.Timedelta(hours=9)
    sl = counts.loc[t0:t0 + pd.Timedelta(hours=8)]
    if d == FIRST_DAY:
        return "設置当日（16:14から記録）。使わない"
    miss = [u for u in UNITS if (sl[(u, "有効")] < 48).sum() >= 3]
    if not miss:
        return ""
    grp = []
    if all(u in miss for u in ("11-1", "11-2", "11-3")):
        grp.append("列11のデータなし")
        miss = [u for u in miss if not u.startswith("11-")]
    if miss:
        grp.append("・".join(miss) + "の排気・吸気温度が欠け（判定不能）")
    return "、".join(grp)


def main(src, dst):
    df = load(src)
    dT = diff(df)
    build(hourly_counts(dT), daily_operation(df, dT), sensitivity(dT), dst)


if __name__ == "__main__":
    main(sys.argv[1], sys.argv[2])
