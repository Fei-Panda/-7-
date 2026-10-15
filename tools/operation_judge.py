"""室外機の稼働状況から、分析に使える日・時間帯を判定する。

目的
    芋緑化ありの室外機（OACP-8・11）と、同じ区画の芋緑化なしの室外機（OACP-3）を
    同じ条件で比べるため、「2台とも運転している時間」だけを分析に使えるようにする。
    その判定を1時間ごと・1日ごとに行い、Excel にまとめる。

入力
    実測データの Excel（シート「整理済みデータ」、1分値、ヘッダ9行）
    例：python3 tools/operation_judge.py 261013_実測データ_完全版.xlsx 出力.xlsx

処理
    1. 室外機ごとに「排気温度 − 吸気温度（3方位の平均）」を1分ごとに求め、
       10分の移動中央値をとる。1.0℃以上なら「運転中」とする。
       - 11-3 は北側吸気（BH列）を全期間除くため2方位の平均
       - 3-1 北側吸気（F列）は、熱電対の異常期間（下の BAD_3_1_NORTH）を除く
    2. 1時間ごとに、室外機の「判定に使える分数」と「運転中の分数」を数える。
       組（例：8-1 と 3-1）ごとに「2台とも判定できる分数」と「2台とも運転中の分数」を数える。
    3. 判定（運転／停止／断続、利用可／不可）は Excel の数式で行う。
       しきい値は「判定条件」シートの黄色のセルにあり、変えると判定がすべて更新される。

出力
    Excel（判定条件、日ごとの判定、時間ごとの判定、時間帯の一覧（列8・列11）、
    時刻別の利用可時間数、データの不具合、しきい値の確認）
"""
import sys
import datetime as dt

import numpy as np
import openpyxl
import pandas as pd
from openpyxl.formatting.rule import CellIsRule, FormulaRule
from openpyxl.styles import Alignment, Border, Font, PatternFill, Side
from openpyxl.utils import get_column_letter as L

# ---------------------------------------------------------------- 設定
# 室外機ごとの列記号：(排気温度, [吸気温度の列])
UNITS = {
    "3-1": ("D", ["E", "F", "G"]), "3-2": ("J", ["K", "L", "M"]), "3-3": ("Q", ["R", "S", "T"]),
    "8-1": ("W", ["X", "Y", "Z"]), "8-2": ("AC", ["AD", "AE", "AF"]), "8-3": ("AK", ["AL", "AM", "AN"]),
    "11-1": ("AR", ["AS", "AT", "AU"]), "11-2": ("AX", ["AY", "AZ", "BA"]), "11-3": ("BF", ["BG", "BI"]),
}
PAIRS = [(f"{c}-{k}", f"3-{k}") for c in ("8", "11") for k in "123"]
DT_THRESHOLD = 1.0          # 運転中とみなす「排気−吸気」［℃］
ROLL_MIN = 10               # 移動中央値の幅［分］
# 3-1 北側吸気の熱電対の異常（夜間に周りより +3〜+10℃高い。急に始まり急に戻る）
BAD_3_1_NORTH = [("2026-09-14 18:47", "2026-09-22 11:40"), ("2026-09-25 23:20", "2026-10-01 11:13")]
HOLIDAYS = {dt.date(2026, 8, 11): "山の日", dt.date(2026, 9, 21): "敬老の日",
            dt.date(2026, 9, 22): "国民の休日", dt.date(2026, 9, 23): "秋分の日",
            dt.date(2026, 10, 12): "スポーツの日"}
FIRST_DAY, LAST_DAY = dt.date(2026, 7, 29), dt.date(2026, 10, 13)
WEEK = "月火水木金土日"


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


def hourly_counts(df, threshold=DT_THRESHOLD):
    hours = pd.date_range(f"{FIRST_DAY} 00:00", f"{LAST_DAY} 23:00", freq="1h")
    valid, run = {}, {}
    for u, (ex, ins) in UNITS.items():
        d = (df[ex] - df[ins].mean(axis=1)).rolling(ROLL_MIN, center=True, min_periods=6).median()
        valid[u] = d.notna()
        run[u] = d >= threshold
    out = pd.DataFrame(index=hours)
    for u in UNITS:
        out[(u, "有効")] = valid[u].resample("1h").sum().reindex(hours, fill_value=0)
        out[(u, "運転")] = run[u].resample("1h").sum().reindex(hours, fill_value=0)
    for a, b in PAIRS:
        both_valid = valid[a] & valid[b]
        out[(a, b, "両方有効")] = both_valid.resample("1h").sum().reindex(hours, fill_value=0)
        out[(a, b, "両方運転")] = (run[a] & run[b]).resample("1h").sum().reindex(hours, fill_value=0)
    return out.astype(int)


# ---------------------------------------------------------------- Excel 書式
FONT = "游ゴシック"
NAVY, TEXT, MUTED = "22354F", "172033", "5E6B7A"
F_BASE = Font(name=FONT, size=10, color=TEXT)
F_BOLD = Font(name=FONT, size=10, color=TEXT, bold=True)
F_HEAD = Font(name=FONT, size=10, color="FFFFFF", bold=True)
F_TITLE = Font(name=FONT, size=14, color=NAVY, bold=True)
F_NOTE = Font(name=FONT, size=9, color=MUTED)
F_INPUT = Font(name=FONT, size=10, color="0000FF", bold=True)
FILL_HEAD = PatternFill("solid", fgColor=NAVY)
FILL_INPUT = PatternFill("solid", fgColor="FFF2B3")
FILL_SUB = PatternFill("solid", fgColor="EEF2F6")
THIN = Side(style="thin", color="C9D1DB")
BOX = Border(left=THIN, right=THIN, top=THIN, bottom=THIN)
CENTER = Alignment(horizontal="center", vertical="center", wrap_text=True)
LEFT = Alignment(horizontal="left", vertical="center", wrap_text=True)
# 判定の色（利用可＝緑、停止＝灰、そろわない＝橙、判定不能＝白地に灰文字）
COLORS = {"利用可": "B9E0AE", "運転": "B9E0AE", "両方停止": "D3D8DD", "停止": "D3D8DD",
          "運転がそろわない": "F8CBAD", "断続": "F8CBAD", "判定不能": "FFFFFF", "不可": "F8CBAD"}


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


# ---------------------------------------------------------------- シートづくり
def build(counts, sens, path):
    wb = openpyxl.Workbook()
    ws0 = wb.active
    ws0.title = "判定条件"
    wsd = wb.create_sheet("日ごとの判定")
    wsh = wb.create_sheet("時間ごとの判定")
    wsg8 = wb.create_sheet("時間帯の一覧_列8")
    wsg11 = wb.create_sheet("時間帯の一覧_列11")
    wst = wb.create_sheet("時刻別の利用可時間数")
    wsb = wb.create_sheet("データの不具合")
    wss = wb.create_sheet("しきい値の確認")

    # ---- 判定条件
    ws = ws0
    ws.column_dimensions["A"].width = 3
    ws.column_dimensions["B"].width = 44
    ws.column_dimensions["C"].width = 12
    ws.column_dimensions["D"].width = 70
    ws["B1"] = "室外機の稼働状況による、分析に使える日・時間帯の判定"
    ws["B1"].font = F_TITLE
    ws["B2"] = ("対象：2026/7/29〜10/13（1分値）。芋緑化ありの室外機と、同じ区画の芋緑化なし（OACP-3）の室外機を"
                "1組とし、2台とも運転している時間だけを「利用可」とする。黄色のセルを変えると、すべての判定が更新される。")
    ws["B2"].font, ws["B2"].alignment = F_NOTE, LEFT
    ws.merge_cells("B2:D2")
    ws.row_dimensions[2].height = 30

    r = 4
    for j, t in enumerate(["段階", "値", "内容"]):
        head(ws, r, 2 + j, t)
    fixed = [
        ("① 1分ごとの運転判定（Pythonで計算）", f"{DT_THRESHOLD}℃以上",
         "排気温度 − 吸気温度（3方位の平均）の10分の移動中央値がこの値以上なら「運転中」。"
         "停止中は約0℃、運転中は2〜10℃に分かれ、その谷が1℃付近にある。11-3は北側を除く2方位。"
         "3-1北側吸気は異常期間を除く（「データの不具合」）"),
    ]
    r += 1
    for a, b, c in fixed:
        cell(ws, r, 2, a, F_BOLD, LEFT)
        cell(ws, r, 3, b, F_BOLD)
        cell(ws, r, 4, c, F_BASE, LEFT)
        ws.row_dimensions[r].height = 45
        r += 1
    inputs = [
        ("valid", "② 1時間の判定に必要な有効データ［分］", 48, "1時間のうち、①の判定ができた分がこれ未満なら「判定不能」（これまでの1時間平均の基準と同じ）"),
        ("run", "③ 室外機が「運転」とみなす割合", 0.8, "有効な分のうち、運転中の分がこの割合以上なら「運転」"),
        ("stop", "④ 室外機が「停止」とみなす割合", 0.2, "運転中の分がこの割合以下なら「停止」。③と④の間は「断続」（1時間の中で運転と停止を繰り返す）"),
        ("both", "⑤ 組が「利用可」となる、2台とも運転中の分数［分］", 30,
         "その1時間に、2台とも運転中の分がこれ以上あれば「利用可」。吸気温度差はこの「2台とも運転中」の分だけで平均する"),
        ("start", "⑥ 日中の始まり［時台］", 9, "日ごとの判定で数える時間帯の始まり（これまでの「日中」と同じ9時台）"),
        ("end", "⑦ 日中の終わり［時台］", 17, "日ごとの判定で数える時間帯の終わり（17時台＝18時まで）"),
        ("days", "⑧ 日として使うのに必要な利用可の時間数", 7, "⑥〜⑦の9時間のうち、利用可がこの時間数以上ある日を「利用可」の日とする（これまでの「9時間中7時間以上」と同じ）"),
    ]
    ref = {}
    for key, a, v, c in inputs:
        cell(ws, r, 2, a, F_BOLD, LEFT)
        x = cell(ws, r, 3, v, F_INPUT, CENTER, "0%" if isinstance(v, float) else "0", FILL_INPUT)
        cell(ws, r, 4, c, F_BASE, LEFT)
        ws.row_dimensions[r].height = 32
        ref[key] = f"判定条件!$C${r}"
        r += 1

    r += 1
    ws.cell(row=r, column=2, value="判定の名前").font = F_BOLD
    r += 1
    legend = [("利用可", "2台とも運転中の分が⑤以上。分析に使う"),
              ("両方停止", "2台とも停止。運転中の比較には使わない（夜間・休日など）"),
              ("運転がそろわない", "片方だけ運転、または断続で2台とも運転中の分が⑤未満。使わない"),
              ("判定不能", "排気温度・吸気温度が欠けていて運転状態がわからない。使わない")]
    for lab, txt in legend:
        cell(ws, r, 2, lab, F_BOLD, CENTER, fill=PatternFill("solid", fgColor=COLORS[lab]))
        cell(ws, r, 3, "", F_BASE)
        cell(ws, r, 4, txt, F_BASE, LEFT)
        r += 1

    # 結果のまとめ（数式）
    r += 1
    ws.cell(row=r, column=2, value="結果のまとめ（7/30〜10/13、日中）").font = F_BOLD
    r += 1
    head(ws, r, 2, "組")
    head(ws, r, 3, "利用可の時間数")
    head(ws, r, 4, "利用可の日数（⑧を満たす日）")
    sum_rows = r + 1
    r += 1
    summary_targets = [f"{a}と{b}" for a, b in PAIRS] + ["列8（3区画とも利用可の日）", "列11（3区画とも利用可の日）"]
    for t in summary_targets:
        cell(ws, r, 2, t, F_BOLD, LEFT)
        r += 1
    summary_end = r - 1

    # ---- 時間ごとの判定
    ws = wsh
    hrs = counts.index
    nrow = len(hrs)
    R0 = 3                       # データの最初の行
    RN = R0 + nrow - 1
    head(ws, 1, 1, "日付", 11)
    head(ws, 1, 2, "時台", 6)
    head(ws, 1, 3, "曜日", 6)
    head(ws, 1, 4, "区分", 8)
    for c in range(1, 5):
        ws.merge_cells(start_row=1, start_column=c, end_row=2, end_column=c)
    col = 5
    ucol = {}
    for u in UNITS:
        ws.merge_cells(start_row=1, start_column=col, end_row=1, end_column=col + 2)
        head(ws, 1, col, f"室外機 {u}")
        for j, t in enumerate(["有効［分］", "運転中［分］", "判定"]):
            head(ws, 2, col + j, t, 7 if j < 2 else 8)
        ucol[u] = col
        col += 3
    pcol = {}
    for a, b in PAIRS:
        ws.merge_cells(start_row=1, start_column=col, end_row=1, end_column=col + 2)
        head(ws, 1, col, f"組 {a}と{b}")
        for j, t in enumerate(["2台とも有効［分］", "2台とも運転中［分］", "判定"]):
            head(ws, 2, col + j, t, 8 if j < 2 else 14)
        pcol[(a, b)] = col
        col += 3
    ccol = {}
    for c in ("8", "11"):
        ws.merge_cells(start_row=1, start_column=col, end_row=2, end_column=col)
        head(ws, 1, col, f"列{c}（3区画とも利用可）", 12)
        ccol[c] = col
        col += 1
    ws.row_dimensions[2].height = 42

    for i, t in enumerate(hrs):
        rr = R0 + i
        d = t.date()
        kind = HOLIDAYS.get(d, "土日" if d.weekday() >= 5 else "平日")
        ws.cell(row=rr, column=1, value=d).number_format = "m/d"
        ws.cell(row=rr, column=2, value=t.hour)
        ws.cell(row=rr, column=3, value=WEEK[d.weekday()])
        ws.cell(row=rr, column=4, value="平日" if kind == "平日" else ("土日" if kind == "土日" else "祝日"))
        for u in UNITS:
            c0 = ucol[u]
            v, n = int(counts.at[t, (u, "有効")]), int(counts.at[t, (u, "運転")])
            ws.cell(row=rr, column=c0, value=v)
            ws.cell(row=rr, column=c0 + 1, value=n)
            V, N = f"{L(c0)}{rr}", f"{L(c0 + 1)}{rr}"
            ws.cell(row=rr, column=c0 + 2,
                    value=f'=IF({V}<{ref["valid"]},"判定不能",IF({N}>={ref["run"]}*{V},"運転",'
                          f'IF({N}<={ref["stop"]}*{V},"停止","断続")))')
        for a, b in PAIRS:
            c0 = pcol[(a, b)]
            ws.cell(row=rr, column=c0, value=int(counts.at[t, (a, b, "両方有効")]))
            ws.cell(row=rr, column=c0 + 1, value=int(counts.at[t, (a, b, "両方運転")]))
            JA, JB = f"{L(ucol[a] + 2)}{rr}", f"{L(ucol[b] + 2)}{rr}"
            BR = f"{L(c0 + 1)}{rr}"
            ws.cell(row=rr, column=c0 + 2,
                    value=f'=IF(OR({JA}="判定不能",{JB}="判定不能"),"判定不能",IF({BR}>={ref["both"]},"利用可",'
                          f'IF(AND({JA}="停止",{JB}="停止"),"両方停止","運転がそろわない")))')
        for c in ("8", "11"):
            js = [f'{L(pcol[(f"{c}-{k}", f"3-{k}")] + 2)}{rr}' for k in "123"]
            ws.cell(row=rr, column=ccol[c],
                    value=f'=IF(AND({js[0]}="利用可",{js[1]}="利用可",{js[2]}="利用可"),"利用可","不可")')
    for row in ws.iter_rows(min_row=R0, max_row=RN, max_col=col - 1):
        for c in row:
            c.font, c.alignment = Font(name=FONT, size=9, color=TEXT), CENTER
    last = L(col - 1)
    color_rules(ws, f"E{R0}:{last}{RN}", ["利用可", "運転", "両方停止", "停止", "運転がそろわない", "断続", "判定不能", "不可"])
    ws.freeze_panes = "E3"
    ws.auto_filter.ref = f"A2:{last}{RN}"

    H = "時間ごとの判定"
    rngA, rngB = f"{H}!$A${R0}:$A${RN}", f"{H}!$B${R0}:$B${RN}"

    def pr(a, b):
        c = L(pcol[(a, b)] + 2)
        return f"{H}!${c}${R0}:${c}${RN}"

    # ---- 日ごとの判定
    ws = wsd
    days = pd.date_range(str(FIRST_DAY), str(LAST_DAY)).date
    head(ws, 1, 1, "日付", 9)
    head(ws, 1, 2, "曜日", 5)
    head(ws, 1, 3, "区分", 10)
    for c in range(1, 4):
        ws.merge_cells(start_row=1, start_column=c, end_row=2, end_column=c)
    col = 4
    dcol = {}
    for a, b in PAIRS:
        ws.merge_cells(start_row=1, start_column=col, end_row=1, end_column=col + 1)
        head(ws, 1, col, f"{a}と{b}")
        head(ws, 2, col, "日中の利用可［時間］", 9)
        head(ws, 2, col + 1, "日の判定", 8)
        dcol[(a, b)] = col
        col += 2
    for c in ("8", "11"):
        ws.merge_cells(start_row=1, start_column=col, end_row=2, end_column=col)
        head(ws, 1, col, f"列{c}（3区画とも利用可の日）", 11)
        dcol[c] = col
        col += 1
    ws.merge_cells(start_row=1, start_column=col, end_row=2, end_column=col)
    head(ws, 1, col, "備考（データの欠け）", 46)
    note_col = col
    ws.row_dimensions[2].height = 42
    D0 = 3
    for i, d in enumerate(days):
        rr = D0 + i
        kind = HOLIDAYS.get(d, "土日" if d.weekday() >= 5 else "平日")
        ws.cell(row=rr, column=1, value=d).number_format = "m/d"
        ws.cell(row=rr, column=2, value=WEEK[d.weekday()])
        ws.cell(row=rr, column=3, value=kind)
        for a, b in PAIRS:
            c0 = dcol[(a, b)]
            ws.cell(row=rr, column=c0,
                    value=f'=COUNTIFS({rngA},$A{rr},{rngB},">="&{ref["start"]},{rngB},"<="&{ref["end"]},{pr(a, b)},"利用可")')
            ws.cell(row=rr, column=c0 + 1, value=f'=IF({L(c0)}{rr}>={ref["days"]},"利用可","不可")')
        for c in ("8", "11"):
            js = [f'{L(dcol[(f"{c}-{k}", f"3-{k}")] + 1)}{rr}' for k in "123"]
            ws.cell(row=rr, column=dcol[c],
                    value=f'=IF(AND({js[0]}="利用可",{js[1]}="利用可",{js[2]}="利用可"),"利用可","不可")')
        ws.cell(row=rr, column=note_col, value=day_note(counts, d))
    DN = D0 + len(days) - 1
    for row in ws.iter_rows(min_row=D0, max_row=DN, max_col=note_col):
        for c in row:
            c.font, c.alignment, c.border = Font(name=FONT, size=9, color=TEXT), CENTER, BOX
    for rr in range(D0, DN + 1):
        ws.cell(row=rr, column=note_col).alignment = LEFT
        if ws.cell(row=rr, column=3).value != "平日":
            for cc in (1, 2, 3):
                ws.cell(row=rr, column=cc).fill = FILL_SUB
    color_rules(ws, f"D{D0}:{L(note_col - 1)}{DN}", ["利用可", "不可"])
    ws.freeze_panes = "D3"

    # 判定条件のまとめ（数式）
    ws = ws0
    D = "日ごとの判定"
    rr = sum_rows
    for a, b in PAIRS:
        c0 = dcol[(a, b)]
        hcol = f"{H}!${L(pcol[(a, b)] + 2)}${R0}:${L(pcol[(a, b)] + 2)}${RN}"
        cell(ws, rr, 3, f'=COUNTIFS({rngA},">="&DATE(2026,7,30),{rngB},">="&{ref["start"]},{rngB},"<="&{ref["end"]},{hcol},"利用可")')
        cell(ws, rr, 4, f'=COUNTIFS({D}!${L(c0 + 1)}${D0 + 1}:${L(c0 + 1)}${DN},"利用可")')
        rr += 1
    for c in ("8", "11"):
        hcol = f"{H}!${L(ccol[c])}${R0}:${L(ccol[c])}${RN}"
        cell(ws, rr, 3, f'=COUNTIFS({rngA},">="&DATE(2026,7,30),{rngB},">="&{ref["start"]},{rngB},"<="&{ref["end"]},{hcol},"利用可")')
        cell(ws, rr, 4, f'=COUNTIFS({D}!${L(dcol[c])}${D0 + 1}:${L(dcol[c])}${DN},"利用可")')
        rr += 1
    ws.cell(row=rr, column=2, value="7/29は設置当日のため数えない。列8・列11の時間数は、その1時間に3組とも利用可の時間").font = F_NOTE

    # ---- 時間帯の一覧（日付×時刻）
    for c, wsg in (("8", wsg8), ("11", wsg11)):
        wsg["A1"] = f"列{c}：組ごとの判定（○＝利用可、停＝両方停止、×＝運転がそろわない、－＝判定不能）"
        wsg["A1"].font = F_BOLD
        wsg.column_dimensions["A"].width = 9
        wsg.column_dimensions["B"].width = 4
        col = 3
        for k in "123":
            a, b = f"{c}-{k}", f"3-{k}"
            wsg.merge_cells(start_row=2, start_column=col, end_row=2, end_column=col + 23)
            head(wsg, 2, col, f"{a}と{b}")
            src = L(pcol[(a, b)] + 2)
            for h in range(24):
                head(wsg, 3, col + h, h, 2.6).font = Font(name=FONT, size=7, color="FFFFFF", bold=True)
                for i, d in enumerate(days):
                    rr = 4 + i
                    idx = i * 24 + h
                    wsg.cell(row=rr, column=col + h,
                             value=f'=IFERROR(CHOOSE(MATCH(INDEX({H}!${src}${R0}:${src}${RN},{idx + 1}),'
                                   f'{{"利用可","両方停止","運転がそろわない","判定不能"}},0),"○","停","×","－"),"")')
            col += 25
            wsg.column_dimensions[L(col - 1)].width = 1.5
        for i, d in enumerate(days):
            rr = 4 + i
            x = wsg.cell(row=rr, column=1, value=d)
            x.number_format, x.font = "m/d", Font(name=FONT, size=8, color=TEXT)
            y = wsg.cell(row=rr, column=2, value=WEEK[d.weekday()])
            y.font = Font(name=FONT, size=8, color="E66C64" if (d.weekday() >= 5 or d in HOLIDAYS) else TEXT)
        last = L(col - 2)
        rng = f"C4:{last}{3 + len(days)}"
        for row in wsg.iter_rows(min_row=4, max_row=3 + len(days), min_col=3, max_col=col - 2):
            for x in row:
                x.font, x.alignment = Font(name=FONT, size=7, color=TEXT), CENTER
        for lab, sym in (("利用可", "○"), ("両方停止", "停"), ("運転がそろわない", "×")):
            wsg.conditional_formatting.add(rng, CellIsRule(operator="equal", formula=[f'"{sym}"'],
                                                           fill=PatternFill("solid", fgColor=COLORS[lab])))
        wsg.freeze_panes = "C4"

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
                                  openpyxl.formatting.rule.ColorScaleRule(start_type="num", start_value=0, start_color="FFFFFF",
                                                                         end_type="max", end_color="5FAF52"))

    # ---- データの不具合
    ws = wsb
    ws["A1"] = "運転の判定に関わるデータの欠け・異常"
    ws["A1"].font = F_BOLD
    for j, (t, w) in enumerate([("対象", 16), ("期間", 30), ("内容", 60), ("扱い", 40)]):
        head(ws, 2, 1 + j, t, w)
    issues = [
        ("3-1 北側吸気（F列）", "9/14 18:47〜9/22 11:40", "夜間（全台停止）に周りの吸気温度より＋3.4〜＋4.4℃高い。急に高くなり、急に戻る", "この期間は使わない（3-1は正面・南側の2方位で計算）。今回見つけた異常"),
        ("3-1 北側吸気（F列）", "9/25 23:20〜10/1 11:13", "同じく＋6.4〜＋9.9℃高い。10/1のデータ回収で戻っている", "同上"),
        ("8-2 排気温度（AC列）", "8/14 15:10〜8/28 10:25", "設置不良で欠損（これまでと同じ）", "8-2は判定不能。組8-2と3-2、列8は使えない"),
        ("3-2 排気温度（J列）", "8/30〜9/7、9/14〜9/22ごろ", "センサー不良で欠損が多い（エラー値・起こり得ない値を空欄にしたもの）", "欠けた時間は判定不能。組8-2・11-2と3-2は使えない時間が多い"),
        ("OACP-11 の熱電対（AR〜BK列）", "9/14 17:23〜10/1 15:01", "元データがない", "列11はこの期間すべて判定不能"),
        ("8-2 の運転", "8/8〜8/14", "8-1・8-3は運転しているが、8-2は日中も停止（データはある）", "判定どおり「運転がそろわない」として使わない"),
        ("3-3 の運転", "7/29〜8/14、9/2〜9/7", "夜間も運転（24時間運転）", "夜間は8-3・11-3が停止しているため「運転がそろわない」"),
        ("ロガーの回収", "8/14・8/28・9/7・9/14・10/1・10/13", "回収作業で数分〜1時間ほど記録が止まる", "有効データが48分未満の時間は判定不能"),
        ("10/13", "14時台以降", "記録の終わり（3列14:16、8列13:43、11列14:01）", "判定不能"),
        ("（参考）屋上の気象", "9/14、10/13", "マルチ環境計測計のデータがない（運転の判定には使わない）", "天気・気温を使う分析ではこの2日間を除く"),
    ]
    for i, row in enumerate(issues):
        for j, v in enumerate(row):
            cell(ws, 3 + i, 1 + j, v, F_BASE, LEFT)
        ws.row_dimensions[3 + i].height = 30

    # ---- しきい値の確認（Pythonで計算した値）
    ws = wss
    ws["A1"] = "①のしきい値（排気−吸気）を変えたときの、日中の利用可の時間数（⑤＝30分、⑥〜⑦＝9〜17時台。Pythonで計算した値）"
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
            value="1.0℃は、停止中（約0℃）と運転中（2〜10℃）の分布の谷にあたる。運転中は排気−吸気が0℃と5〜10℃の間を行き来するため、"
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


def sensitivity(df):
    rows = []
    base = None
    for th in (0.5, 1.0, 1.5, 2.0):
        c = hourly_counts(df, th)
        day = (c.index.hour >= 9) & (c.index.hour <= 17) & (c.index >= "2026-07-30")
        vals = []
        for a, b in PAIRS:
            ok = (c[(a, "有効")] >= 48) & (c[(b, "有効")] >= 48) & (c[(a, b, "両方運転")] >= 30)
            vals.append(int(ok[day].sum()))
        cls = {}
        for u in UNITS:
            v, n = c[(u, "有効")], c[(u, "運転")]
            cls[u] = np.where(v < 48, "x", np.where(n >= 0.8 * v, "運転", np.where(n <= 0.2 * v, "停止", "断続")))
        cls = pd.DataFrame(cls, index=c.index)[day]
        if th == 1.0:
            base = cls
        rows.append((th, vals, cls))
    out = []
    for th, vals, cls in rows:
        m = base != "x"
        out.append((th, vals, float(((cls != base) & m).sum().sum() / m.sum().sum())))
    return out


def main(src, dst):
    df = load(src)
    counts = hourly_counts(df)
    sens = sensitivity(df)
    build(counts, sens, dst)


if __name__ == "__main__":
    main(sys.argv[1], sys.argv[2])
