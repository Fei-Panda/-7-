"""v2 のコメントを反映して v3 を作る。

使い方: python3 tools/revise_v3.py <v2.pptx> <出力.pptx>

v2 は PowerPoint で編集済みのため、生成スクリプトに戻さず XML を直接書き換える。
"""
import os
import re
import shutil
import sys
import tempfile
import zipfile

EMU = 914400
FONT = ('<a:latin typeface="游ゴシック" pitchFamily="34" charset="0"/>'
        '<a:ea typeface="游ゴシック" pitchFamily="34" charset="-122"/>'
        '<a:cs typeface="游ゴシック" pitchFamily="34" charset="-120"/>')
DARK, MUTED, LINE = "172033", "5E6B7A", "C9D1DB"


def emu(v):
    return int(round(v * EMU))


def esc(s):
    return s.replace("&", "&amp;").replace("<", "&lt;").replace(">", "&gt;")


# ---------------------------------------------------------------- XML 部品

def run(text, sz=1050, color=DARK, b=False):
    bold = ' b="1"' if b else ""
    return (f'<a:r><a:rPr lang="ja-JP" altLang="en-US" sz="{sz}"{bold} dirty="0">'
            f'<a:solidFill><a:srgbClr val="{color}"/></a:solidFill>{FONT}</a:rPr>'
            f'<a:t>{esc(text)}</a:t></a:r>')


def para(runs, sz=1050, bullet=False, algn=None, after=300):
    al = f' algn="{algn}"' if algn else ""
    if bullet:
        ppr = (f'<a:pPr marL="152400" indent="-152400"{al}>'
               f'<a:spcAft><a:spcPts val="{after}"/></a:spcAft>'
               '<a:buSzPct val="100000"/><a:buChar char="•"/></a:pPr>')
    else:
        ppr = (f'<a:pPr marL="0" indent="0"{al}>'
               f'<a:spcAft><a:spcPts val="{after}"/></a:spcAft><a:buNone/></a:pPr>')
    return f'<a:p>{ppr}{"".join(runs)}<a:endParaRPr lang="ja-JP" altLang="en-US" sz="{sz}" dirty="0"/></a:p>'


def bullets(items, sz=1050, after=300):
    return [para([run(t, sz)], sz, bullet=True, after=after) for t in items]


def labeled(label, text, sz=1050, label_color="ED7D31", newline=True, after=0):
    """「読み取り」などの見出し＋本文。newline=False なら同じ行に続ける。"""
    if newline:
        return [para([run(label, sz, label_color, b=True)], sz, after=0),
                para([run(text, sz)], sz, after=after)]
    return [para([run(label + "　", sz, label_color, b=True), run(text, sz)], sz, after=after)]


def textbox(sid, name, x, y, w, h, paras, anchor="t"):
    return (f'<p:sp><p:nvSpPr><p:cNvPr id="{sid}" name="{name}"/><p:cNvSpPr txBox="1"/><p:nvPr/></p:nvSpPr>'
            f'<p:spPr><a:xfrm><a:off x="{emu(x)}" y="{emu(y)}"/><a:ext cx="{emu(w)}" cy="{emu(h)}"/></a:xfrm>'
            '<a:prstGeom prst="rect"><a:avLst/></a:prstGeom><a:noFill/><a:ln/></p:spPr>'
            f'<p:txBody><a:bodyPr wrap="square" lIns="0" tIns="0" rIns="0" bIns="0" rtlCol="0" anchor="{anchor}"/>'
            f'<a:lstStyle/>{"".join(paras)}</p:txBody></p:sp>')


def rect(sid, name, x, y, w, h, fill, geom="rect"):
    av = '<a:avLst><a:gd name="adj" fmla="val 2532"/></a:avLst>' if geom == "roundRect" else "<a:avLst/>"
    return (f'<p:sp><p:nvSpPr><p:cNvPr id="{sid}" name="{name}"/><p:cNvSpPr/><p:nvPr/></p:nvSpPr>'
            f'<p:spPr><a:xfrm><a:off x="{emu(x)}" y="{emu(y)}"/><a:ext cx="{emu(w)}" cy="{emu(h)}"/></a:xfrm>'
            f'<a:prstGeom prst="{geom}">{av}</a:prstGeom><a:solidFill><a:srgbClr val="{fill}"/></a:solidFill>'
            f'<a:ln w="12700"><a:solidFill><a:srgbClr val="{fill}"/></a:solidFill></a:ln></p:spPr>'
            '<p:txBody><a:bodyPr/><a:lstStyle/><a:p><a:endParaRPr lang="ja-JP" altLang="en-US"/></a:p></p:txBody></p:sp>')


def table(sid, name, x, y, col_w, row_h, rows, sz=900, header_fill="22354F", bold_last=False):
    """rows[0] は見出し行。各セルは文字列。"""
    def border(tag):
        return (f'<a:{tag} w="9525" cap="flat" cmpd="sng" algn="ctr"><a:solidFill><a:srgbClr val="{LINE}"/>'
                f'</a:solidFill><a:prstDash val="solid"/></a:{tag}>')
    borders = "".join(border(t) for t in ("lnL", "lnR", "lnT", "lnB"))
    grid = "".join(f'<a:gridCol w="{emu(w)}"/>' for w in col_w)
    trs = []
    for ri, row in enumerate(rows):
        head = ri == 0
        last = bold_last and ri == len(rows) - 1
        tcs = []
        for text in row:
            color = "FFFFFF" if head else DARK
            fill = (f'<a:solidFill><a:srgbClr val="{header_fill}"/></a:solidFill>' if head else
                    '<a:solidFill><a:srgbClr val="F3F5F8"/></a:solidFill>' if last else
                    '<a:solidFill><a:srgbClr val="FFFFFF"/></a:solidFill>')
            tcs.append('<a:tc><a:txBody><a:bodyPr/><a:lstStyle/>'
                       f'<a:p><a:pPr marL="0" indent="0" algn="ctr"><a:buNone/></a:pPr>'
                       f'{run(text, sz, color, b=head or last)}</a:p></a:txBody>'
                       f'<a:tcPr marL="45720" marR="45720" marT="18288" marB="18288" anchor="ctr">{borders}{fill}</a:tcPr></a:tc>')
        trs.append(f'<a:tr h="{emu(row_h)}">{"".join(tcs)}</a:tr>')
    w, h = sum(col_w), row_h * len(rows)
    return (f'<p:graphicFrame><p:nvGraphicFramePr><p:cNvPr id="{sid}" name="{name}"/>'
            '<p:cNvGraphicFramePr><a:graphicFrameLocks noGrp="1"/></p:cNvGraphicFramePr><p:nvPr/></p:nvGraphicFramePr>'
            f'<p:xfrm><a:off x="{emu(x)}" y="{emu(y)}"/><a:ext cx="{emu(w)}" cy="{emu(h)}"/></p:xfrm>'
            '<a:graphic><a:graphicData uri="http://schemas.openxmlformats.org/drawingml/2006/table">'
            f'<a:tbl><a:tblPr firstRow="1"/><a:tblGrid>{grid}</a:tblGrid>{"".join(trs)}</a:tbl>'
            '</a:graphicData></a:graphic></p:graphicFrame>')


# ---------------------------------------------------------------- スライド操作

class Slide:
    def __init__(self, root, n):
        self.path = os.path.join(root, "ppt", "slides", f"slide{n}.xml")
        with open(self.path, encoding="utf-8") as f:
            self.xml = f.read()

    def save(self):
        with open(self.path, "w", encoding="utf-8") as f:
            f.write(self.xml)

    def _shapes(self):
        return list(re.finditer(r'<p:(sp|graphicFrame|cxnSp|pic)>.*?</p:\1>', self.xml, re.S))

    def shape(self, sid):
        for m in self._shapes():
            if re.search(rf'<p:cNvPr id="{sid}"[ />]', m.group(0)):
                return m
        raise KeyError(f"{self.path}: id={sid} がない")

    def replace_shape(self, sid, new):
        m = self.shape(sid)
        self.xml = self.xml[:m.start()] + new + self.xml[m.end():]

    def delete(self, *sids):
        for sid in sids:
            self.replace_shape(sid, "")

    def add(self, *shapes_xml):
        i = self.xml.rindex("</p:spTree>")
        self.xml = self.xml[:i] + "".join(shapes_xml) + self.xml[i:]

    def set_paras(self, sid, paras):
        """テキストボックスの段落を差し替える（bodyPr などはそのまま）。"""
        m = self.shape(sid)
        x = m.group(0)
        new = "".join(paras)
        # 文字サイズを指定していない（既定の 10.5pt の）ときは、元の文字サイズに合わせる
        orig = re.search(r'<a:r><a:rPr [^>]*sz="(\d+)"', x)
        if orig and orig.group(1) != "1050" and set(re.findall(r'sz="(\d+)"', new)) == {"1050"}:
            new = new.replace('sz="1050"', f'sz="{orig.group(1)}"')
        paras = [new]
        start = x.index("<a:lstStyle/>") + len("<a:lstStyle/>")
        end = x.index("</p:txBody>")
        self.replace_shape(sid, x[:start] + "".join(paras) + x[end:])

    def set_text(self, sid, old, new):
        m = self.shape(sid)
        x = m.group(0)
        if f"<a:t>{old}</a:t>" not in x:
            raise KeyError(f"{self.path}: id={sid} に「{old}」がない")
        self.replace_shape(sid, x.replace(f"<a:t>{old}</a:t>", f"<a:t>{esc(new)}</a:t>", 1))

    def set_xfrm(self, sid, x=None, y=None, w=None, h=None):
        m = self.shape(sid)
        s = m.group(0)
        o = re.search(r'<a:off x="(-?\d+)" y="(-?\d+)"/><a:ext cx="(\d+)" cy="(\d+)"/>', s)
        vals = [int(v) for v in o.groups()]
        for i, v in enumerate((x, y, w, h)):
            if v is not None:
                vals[i] = emu(v)
        s = s[:o.start()] + (f'<a:off x="{vals[0]}" y="{vals[1]}"/><a:ext cx="{vals[2]}" cy="{vals[3]}"/>') + s[o.end():]
        self.replace_shape(sid, s)

    def squeeze_below(self, top, new_top, bottom=5.17):
        """top より下の図形を new_top〜bottom に縦方向に詰める（ページ番号は除く）。"""
        k = (bottom - new_top) / (bottom - top)
        for m in reversed(self._shapes()):
            s = m.group(0)
            o = re.search(r'<a:off x="(-?\d+)" y="(-?\d+)"/><a:ext cx="(\d+)" cy="(\d+)"/>', s)
            if not o:
                continue
            x, y, w, h = (int(v) / EMU for v in o.groups())
            if y < top - 0.05 or y > 5.25:
                continue
            ny, nh = new_top + (y - top) * k, h * k
            s = s[:o.start()] + (f'<a:off x="{o.group(1)}" y="{emu(ny)}"/>'
                                 f'<a:ext cx="{o.group(3)}" cy="{emu(nh)}"/>') + s[o.end():]
            self.xml = self.xml[:m.start()] + s + self.xml[m.end():]

    def next_id(self):
        return max(int(v) for v in re.findall(r'<p:cNvPr id="(\d+)"', self.xml)) + 1


# ---------------------------------------------------------------- 修正内容

CALC_NOTE = {
    # 吸気温度差の求め方（コメント：p.11・p.16「すべての吸気温度の平均か、代表値か」）
    "col": ("求め方", "同じ区画・方位の熱電対どうしの差を、列8は9組、列11は8組すべて平均（1点の代表値ではない）"),
    "sec": ("求め方", "同じ方位の熱電対どうしの差を、区画ごとに3方位（11-3は2方位）平均（1点の代表値ではない）"),
}


def add_calc_note(s, kind):
    """方法の下に「求め方」の帯を入れ、その下の図形を詰める。"""
    s.squeeze_below(1.45, 1.75)
    label, text = CALC_NOTE[kind]
    sid = s.next_id()
    s.add(rect(sid, "求め方 背景", 0.4, 1.42, 9.2, 0.27, "F3F5F8", "roundRect"),
          textbox(sid + 1, "求め方", 0.5, 1.42, 9.0, 0.27,
                  [para([run(label + "　", 1000, "118DFF", b=True), run(text, 1000)], 1000, after=0)],
                  anchor="ctr"))


def revise(root):
    S = {n: Slide(root, n) for n in range(1, 29)}

    # ---- コメント（p.1）：結果の文章を平易に。数値は変えず、先に結論を言葉で書く
    s = S[2]
    s.set_paras(9, [para([run("32日中28日で列3より低い。平均0.26℃低く、運転の条件をそろえても0.16℃低い")])])
    s.set_paras(13, [para([run("平均の差は0.06℃と小さく、低いとは言えない。運転の条件をそろえると、逆に0.23℃高い")])])
    s.set_paras(17, [para([run("日中の気温が31℃以上の日は列8が0.27℃高く、28℃未満の日は0.34℃低い。天気や風より気温の影響が大きい")])])
    s.set_paras(21, [para([run("両列とも左の区画は列3より高い。低くなるのは列11の中央（0.59℃低い）と、列8の中央・右（0.38・0.42℃低い）")])])
    s.set_paras(25, [para([run("中央の列8のほうが外側の列11より、どの区画でも高い（0.12〜0.50℃）")])])
    s.set_paras(27, labeled("注意", "排気温度から見ると、列3は休日や夕方（20時ごろまで）も動いている。"
                                   "休日・夕方は列3だけが動く時間が多いため、差が大きく見える",
                            label_color="E66C64", newline=False))

    s = S[7]
    s.set_paras(11, bullets(["熱電対ごとのずれ（偏差の平均）は−0.21〜＋0.29℃の範囲",
                             "列3の9点はそろって高め。平均＋0.17℃（列8は−0.05℃、列11は−0.07℃）",
                             "室外機の運転状態は、排気温度から推定した（p.10）"]))
    s.set_paras(13, labeled("読み取り", "熱電対1本ごとのずれは±0.3℃以内で小さい。ただし列3は全体に約0.2℃高く、"
                                       "熱電対の違いか場所の違いかは区別できない。この夜間の差はp.10で差し引く"))

    s = S[9]
    s.set_text(6, "9組（列11は8組）平均の日中平均。上段：吸気温度、下段：吸気温度差（緑・青：列3より低い日、赤：高い日、灰：休日）",
               "全点平均（列8は9組、列11は8組）の日中平均。上段：吸気温度、下段：吸気温度差（緑・青：低い日、赤：高い日、灰：休日）")
    s.set_paras(14, labeled("列8", "列3より低い日は32日中18日。平均−0.06℃で、差があるとは言えない（95％区間−0.28〜＋0.13℃）",
                            label_color="01B8AA", newline=False))
    s.set_paras(16, labeled("列11", "列3より低い日は32日中28日。平均0.26℃低い（95％区間−0.39〜−0.15℃）",
                            label_color="118DFF", newline=False))

    s = S[10]
    s.set_paras(10, bullets(["両方が動いている時間で比べると、列11は33日中29日で低い（−0.27℃）。列8は逆に高い（＋0.10℃）",
                             "夜間（両方停止）の差：列8は−0.10℃、列11は−0.15℃",
                             "夜間の差を引いた後：列11は0.16℃低いまま（区間−0.29〜−0.03）。列8は0.23℃高い（＋0.07〜＋0.40）"]))
    s.set_paras(26, labeled("読み取り", "列11は条件をそろえても低い。列8は運転中に列3より高い。"
                                       "夜間の差には芋緑化の効果も含まれるかもしれないため、補正後の値は小さめの見積もり"))

    # ---- コメント（p.11）：吸気温度の算出方法を明記
    s = S[11]
    s.set_paras(13, bullets(["列11：どの時刻も列3より低い（中央値−0.15〜−0.34℃）。低い日は66〜93％",
                             "列11で差が大きいのは17時台（−0.34℃）・9時台（−0.33℃）・13時台（−0.30℃）",
                             "列8：中央値は−0.25〜＋0.07℃で、ほぼ差がない。14時台だけ列3より高い"]))
    s.set_paras(15, labeled("読み取り", "時刻による違いは小さく、エアネット分析で見られた14時のピークはない。"
                                       "9時台・17時台にやや差が大きいのは、運転の開始・終了時刻のずれのため（p.10）"))
    add_calc_note(s, "col")

    for n, items in ((12, ["低い日は32日中3日。平均0.63℃高い（95％区間＋0.39〜＋0.85℃）",
                           "低い日は32日中26日。平均0.38℃低い（95％区間−0.60〜−0.19℃）",
                           "低い日は32日中28日。平均0.42℃低い（95％区間−0.63〜−0.24℃）"]),
                     (13, ["低い日は32日中9日。平均0.13℃高いが、差は確かでない（区間−0.08〜＋0.29℃）",
                           "32日すべてで低い。平均0.59℃低い（95％区間−0.68〜−0.49℃）",
                           "低い日は32日中28日。平均0.36℃低い（95％区間−0.50〜−0.24℃）"])):
        for sid, text in zip((11, 16, 21), items):
            m = S[n].shape(sid).group(0)
            label = re.findall(r"<a:t>([^<]*)</a:t>", m)[0].strip()
            if n == 13 and sid == 11:
                label = "やや高い"
            color = re.search(r'<a:srgbClr val="(\w+)"', m[m.index("<a:r>"):]).group(1)
            S[n].set_paras(sid, labeled(label, text, label_color=color, newline=False))

    s = S[14]
    s.set_paras(13, bullets(["列8：左の区画1は列3より高い（中央値＋0.62℃、90％の時間で高い）。区画2・3は低い（−0.35・−0.39℃）",
                             "列11：中央の区画2が最も低い（−0.55℃、92％の時間で低い）。区画3は−0.40℃、区画1は＋0.17℃",
                             "列11の区画2は区画3より0.22℃低い（区間−0.33〜−0.11）。列8の区画2と3には差がない",
                             "同じ区画で比べると、どの区画も列8が列11より高い（＋0.12〜＋0.50℃）"], sz=1000, after=200))
    s.set_paras(15, labeled("読み取り", "最も低くなるのは、列11は中央の区画2、列8は中央・右の区画2・3。"
                                       "左の区画1は両列とも列3より高い。「中央の列ほど低い」という仮説は支持されない"))
    add_calc_note(s, "sec")

    # ---- コメント（p.15）：淡い赤＝晴れの休日。凡例を勤務日／休日に分けて明記
    s = S[15]
    s.delete(10, 11, 12, 13, 14, 15, 16, 17)
    sid = s.next_id()
    legend = [("晴れ", "E66C64"), ("晴れ（休日）", "F6C3BF"), ("曇り", "8A949E"),
              ("曇り（休日）", "D3D8DD"), ("雨", "118DFF"), ("雨（休日）", "A8D3FF")]
    x = 0.75
    for label, color in legend:
        w = 0.17 + 0.135 * len(label) + 0.12
        s.add(rect(sid, f"凡例 {label}", x, 5.0, 0.16, 0.14, color),
              textbox(sid + 1, f"凡例文字 {label}", x + 0.21, 4.96, w - 0.17, 0.22,
                      [para([run(label, 900, MUTED)], 900, after=0)], anchor="ctr"))
        sid += 2
        x += w + 0.05
    s.add(textbox(sid, "凡例の注記", 0.75, 5.17, 5.4, 0.18,
                  [para([run("下段の淡い色は休日（土日・8/11）。濃い色は勤務日", 800, MUTED)], 800, after=0)]))
    s.set_paras(23, bullets(["晴れ（15日）：列8＋0.22℃、列11−0.13℃",
                             "曇り（8日）：列8−0.28℃、列11−0.30℃",
                             "雨（8日）：列8−0.37℃、列11−0.46℃"], sz=850, after=200))
    s.set_paras(25, labeled("読み取り", "晴れが続いた8月中旬〜下旬は、列8が列3より高い。"
                                       "曇り・雨が多い9月は、両列とも列3より低い", sz=900, newline=False))

    # ---- コメント（p.16）：天気別の値も全点平均であることを明記
    s = S[16]
    s.set_paras(13, bullets(["曇り・雨の日：両列ともほぼ毎日、列3より低い（列8 −0.28・−0.37℃、列11 −0.30・−0.46℃）",
                             "晴れの日：列11は15日中11日で低いが、差は−0.13℃と小さい",
                             "晴れの日：列8は15日中13日で列3より高い（＋0.22℃）"]))
    s.set_paras(15, labeled("読み取り", "晴れの日は差が小さく、列8では逆に高くなる。"
                                       "ただし晴れの日は暑い日でもあるため、次の2枚で気温の影響と分けて確かめる"))
    add_calc_note(s, "col")

    s = S[17]
    s.set_paras(13, bullets(["28℃未満の日：列8は9日中8日、列11は9日すべてで低い（−0.34・−0.41℃）",
                             "31℃以上の日：列8は11日中10日で列3より高い（＋0.27℃）。列11も差が小さい（−0.12℃）",
                             "暑い日ほど、両列とも差が小さくなる（列3に近づく）"]))
    s.set_paras(15, labeled("読み取り", "暑い日ほど低減効果が小さい。冷房の負荷が大きく排熱が多い日に、"
                                       "芋緑化側が高くなりやすいと考えられる"))
    add_calc_note(s, "col")

    # ---- コメント（p.18）：「効果あり／見られない」は意味があいまい → 何を判定したかが分かる言葉に
    s = S[18]
    g = s.shape(19).group(0)
    widths = [0.7, 0.75, 1.2, 1.2, 1.25]
    cols = re.findall(r'<a:gridCol w="\d+">', g)
    for old, w in zip(cols, widths):
        g = g.replace(old, f'<a:gridCol w="{emu(w)}">', 1)
    g = g.replace("<a:t>判定</a:t>", "<a:t>差への影響</a:t>", 1)
    g = g.replace("<a:t>効果あり</a:t>", "<a:t>あり</a:t>", 1)
    g = g.replace("<a:t>見られない</a:t>", "<a:t>はっきりしない</a:t>")
    g = g.replace(' err="1"', "")
    s.replace_shape(19, g)
    s.set_paras(11, [para([run("上段：条件が「変化」の分だけ増えたときの吸気温度差の変化、下段：95％区間。"
                               "「あり」＝両列とも区間が0℃を含まない。「はっきりしない」＝区間が0℃を含み、"
                               "影響の向きを決められない（影響がないとは限らない）", 800, MUTED)], 800, after=0)])
    s.set_xfrm(11, h=0.42)
    s.set_xfrm(13, y=4.6, h=0.6)
    s.set_xfrm(12, y=4.55, h=0.66) if _has(s, 12) else None
    s.set_paras(13, labeled("読み取り", "気温が1℃上がると、吸気温度差は約0.05℃大きくなる（低減効果が小さくなる）。"
                                       "20℃の日と35℃の日では約0.8℃違う。照度比・風速・経過日数の影響ははっきりしない",
                            sz=950, newline=False))

    # ---- コメント（p.19）：風速区分ごとの時間数の表、考察を追加
    s = S[19]
    s.set_text(6, "屋上の風速（1時間平均）で4区分。勤務日の日中の1時間平均の吸気温度差（計274〜276時間）",
               "屋上の風速（1時間平均）で4区分。勤務日の日中の1時間ごとの吸気温度差（時間数は右の表）")
    add_calc_note(s, "col")
    s.delete(11, 12, 13, 14, 15)
    sid = s.next_id()
    rows = [["区分", "風速 m/s", "列8", "列11", "割合"],
            ["区分1", "1.0未満", "50", "51", "18％"],
            ["区分2", "1.0〜2.0", "103", "104", "38％"],
            ["区分3", "2.0〜3.0", "74", "74", "27％"],
            ["区分4", "3.0以上", "47", "47", "17％"],
            ["合計", "", "274", "276", "100％"]]
    s.add(textbox(sid, "時間数 見出し", 6.1, 1.78, 3.5, 0.22,
                  [para([run("風速区分ごとの時間数［時間］（勤務日の日中）", 1050, "22354F", b=True)], 1050, after=0)],
                  anchor="ctr"),
          table(sid + 1, "時間数の表", 6.1, 2.02, [0.55, 0.8, 0.75, 0.8, 0.6], 0.2, rows, sz=850, bold_last=True),
          rect(sid + 2, "結果 背景", 6.1, 3.42, 3.5, 0.74, "EAF4FF", "roundRect"),
          textbox(sid + 3, "結果", 6.25, 3.47, 3.2, 0.66,
                  [para([run("結果", 1050, "22354F", b=True)], 1050, after=100)] +
                  bullets(["列8：風が強いほど差が0℃に近づく",
                           "列11：どの区分も約0.2℃低く、変わらない"], sz=950, after=100)),
          rect(sid + 4, "考察 背景", 6.1, 4.22, 3.5, 0.98, "FFF7D1", "roundRect"),
          textbox(sid + 5, "考察", 6.25, 4.27, 3.2, 0.9,
                  [para([run("考察", 950, "ED7D31", b=True)], 950, after=0),
                   para([run("風が強いと通路の空気が入れ替わり、差は0℃に近づくと考えられる。"
                             "列8はその形だが、区分4の72％は晴れの暑い時間で、気温の影響（p.17）が重なる。"
                             "列11や回帰（p.18）では風の影響ははっきりしない。次は風向別に確かめる", 850)], 850, after=0)]))

    s = S[20]
    s.set_paras(12, labeled("読み取り", "9月に低減効果が大きいのは、主に気温が下がったためと考えられる"
                                       "（回帰では経過日数の影響ははっきりしない）。緑被率を加えて再確認する"))

    s = S[22]
    s.set_paras(17, labeled("読み取り", "日差しが強い時間に、芋緑化側が高くなる。朝に芋緑化側が高いのは、"
                                       "列8・11が列3より先に運転を始めるため（p.10）"))
    s = S[23]
    s.set_paras(17, labeled("読み取り", "この2日では風の強い日に低減効果が小さい。しかし気温をそろえた回帰（p.18）とは"
                                       "向きが合わず、1組だけでは風の影響は判断できない"))
    s = S[24]
    s.set_paras(14, labeled("読み取り", "7〜8時に列8・11が運転を始め、9時に列3が始まる（p.10）。"
                                       "この開始時刻のずれで、朝の吸気温度差が大きく上下する"))

    s = S[26]
    m = s.shape(8).group(0)
    s.set_paras(8, bullets(["列11（外側）は列3より低い。32日中28日で低く（平均0.26℃）、条件をそろえても0.16℃低い",
                            "列8（中央）は列3より低いとは言えない。運転中だけで比べると0.10℃、夜間の差を除くと0.23℃高い",
                            "暑い日ほど低減効果が小さい（気温1℃あたり約0.05℃）。晴れの日や9月との違いも、主に気温で説明できる",
                            "同じ列の中では、左の区画が列3より高い。低くなるのは列11の中央、列8の中央・右",
                            "「中央の列ほど低い」という仮説は支持されない"],
                           sz=int(re.search(r'sz="(\d+)"', m[m.index("<a:r>"):]).group(1))))

    s = S[28]
    s.set_paras(13, bullets(["日射遮蔽率は79〜95％で、どの時間も8割以上を遮っている",
                             "遮った日射量と吸気温度差の相関は＋0.31。多く遮った時間ほど、低減効果はやや小さい",
                             "列11で同じ計算をすると相関は＋0.08"]))
    s.set_paras(15, labeled("読み取り", "この期間では「多く遮るほど低減効果が大きい」関係は見られない。"
                                       "日射が強い時間は晴れの時間と重なり、データも59時間と少ない"))

    for s in S.values():
        s.save()

    # ---- コメント（p.23）：線を細くする（代表日の3枚で統一）
    for n in range(34, 44):
        p = os.path.join(root, "ppt", "charts", f"chart{n}.xml")
        with open(p, encoding="utf-8") as f:
            x = f.read()
        with open(p, "w", encoding="utf-8") as f:
            f.write(x.replace('<a:ln w="22225"', '<a:ln w="11430"'))


def _has(s, sid):
    try:
        s.shape(sid)
        return True
    except KeyError:
        return False


def strip_comments(root):
    """対応済みのコメントを取り除く。"""
    shutil.rmtree(os.path.join(root, "ppt", "comments"), ignore_errors=True)
    if os.path.exists(os.path.join(root, "ppt", "authors.xml")):
        os.remove(os.path.join(root, "ppt", "authors.xml"))
    for rels in [os.path.join(root, "ppt", "_rels", "presentation.xml.rels")] + [
            os.path.join(root, "ppt", "slides", "_rels", f) for f in os.listdir(os.path.join(root, "ppt", "slides", "_rels"))]:
        with open(rels, encoding="utf-8") as f:
            x = f.read()
        x = re.sub(r'<Relationship [^>]*(comments|authors)[^>]*/>', "", x)
        with open(rels, "w", encoding="utf-8") as f:
            f.write(x)
    for f in os.listdir(os.path.join(root, "ppt", "slides")):
        if f.endswith(".xml"):
            p = os.path.join(root, "ppt", "slides", f)
            with open(p, encoding="utf-8") as fh:
                x = fh.read()
            x = re.sub(r'<p:ext uri="\{6950BFC3-D8DA-4A85-94F7-54DA5524770B\}">.*?</p:ext>', "", x, flags=re.S)
            x = x.replace("<p:extLst></p:extLst>", "")
            with open(p, "w", encoding="utf-8") as fh:
                fh.write(x)
    ct = os.path.join(root, "[Content_Types].xml")
    with open(ct, encoding="utf-8") as f:
        x = f.read()
    x = re.sub(r'<Override PartName="/ppt/(comments/[^"]*|authors\.xml)"[^>]*/>', "", x)
    with open(ct, "w", encoding="utf-8") as f:
        f.write(x)


def main(src, dst):
    with tempfile.TemporaryDirectory() as tmp:
        with zipfile.ZipFile(src) as z:
            z.extractall(tmp)
            names = z.namelist()
        revise(tmp)
        strip_comments(tmp)
        if os.path.exists(dst):
            os.remove(dst)
        with zipfile.ZipFile(dst, "w", zipfile.ZIP_DEFLATED) as z:
            # [Content_Types].xml を先頭に、元の並びを保つ
            order = [n for n in names if os.path.exists(os.path.join(tmp, n)) and not n.endswith("/")]
            for n in order:
                z.write(os.path.join(tmp, n), n)


if __name__ == "__main__":
    main(sys.argv[1], sys.argv[2])
