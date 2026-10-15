// 日中（9〜18時）の室外機の稼働時間のスライドを作る（v4 と同じ体裁）。
//
// 目的  tools/operating_hours.py が出力した JSON から、稼働時間の推移を PowerPoint のグラフにする。
//       グラフは PowerPoint のネイティブグラフ（右クリック→データの編集で値を確認できる）。
// 入力  operating_hours.py の JSON
// 処理  スライド3枚（列ごとの推移、室外機ごとの推移、研究に使える時間）を作る。
//       欠けた日は空欄にして、線が途切れるようにする。
// 出力  .pptx
//
//   NODE_PATH=<pptxgenjs のある node_modules> node tools/operating_hours_slides.js 入力.json 出力.pptx
const fs = require("fs");
const path = require("path");
const pptxgen = require("pptxgenjs");
const JSZip = require(require.resolve("jszip", { paths: [require.resolve("pptxgenjs")] }));

const [, , SRC, OUT] = process.argv;
const D = JSON.parse(fs.readFileSync(SRC, "utf-8"));

const THEME = {
  name: "屋上芋緑化",
  headFontFace: "游ゴシック",
  bodyFontFace: "游ゴシック",
  colors: {
    dk1: "172033", lt1: "FFFFFF", dk2: "22354F", lt2: "F3F5F8",
    accent1: "01B8AA", accent2: "118DFF", accent3: "E66C64", accent4: "ED7D31",
    accent5: "8E5BC4", accent6: "5E6B7A", hlink: "118DFF", folHlink: "8E5BC4",
  },
};
const HEX = THEME.colors;
const COLOR = { 列3: "5E6B7A", 列8: "01B8AA", 列11: "118DFF" };
const UNIT_COLOR = {
  "3-1": "22354F", "3-2": "5E6B7A", "3-3": "A7B1BC",
  "8-1": "00786E", "8-2": "01B8AA", "8-3": "7FD8CE",
  "11-1": "0B5CAD", "11-2": "118DFF", "11-3": "8CC4FF",
};
const BAND = "E9ECF1";
const FOOTER = "屋上芋緑化｜2026年実測｜分析結果（暫定）";

const pres = new pptxgen();
pres.layout = "LAYOUT_16x9";
pres.title = "室外機の稼働時間の推移";
pres.theme = { headFontFace: THEME.headFontFace, bodyFontFace: THEME.bodyFontFace };
const C = pres.SchemeColor;

pres.defineSlideMaster({
  title: "CONTENT",
  background: { color: HEX.lt1 },
  objects: [
    { rect: { x: 0, y: 0, w: 10, h: 0.62, fill: { color: C.text2 } } },
    { rect: { x: 0, y: 0.62, w: 0.08, h: 5.005, fill: { color: C.accent1 } } },
    { text: { text: FOOTER, options: { x: 0.4, y: 5.3, w: 6, h: 0.25, fontSize: 9, color: C.accent6, margin: 0, valign: "middle" } } },
    { placeholder: { options: { name: "title", type: "title", x: 0.4, y: 0.08, w: 7.9, h: 0.46, fontSize: 20, bold: true, color: C.background1, align: "left", valign: "middle", margin: 0 }, text: "" } },
  ],
  slideNumber: { x: 9.1, y: 5.3, w: 0.5, h: 0.25, fontSize: 9, color: HEX.accent6, align: "right" },
});

pres.addSection({ title: "前提" });

function frame(title, tag, purpose, method) {
  const s = pres.addSlide({ masterName: "CONTENT", sectionTitle: "前提" });
  s.addText(title, { placeholder: "title" });
  s.addShape(pres.shapes.ROUNDED_RECTANGLE, { x: 8.45, y: 0.17, w: 1.2, h: 0.3, rectRadius: 0.15, fill: { color: C.accent6 }, line: { color: C.accent6 }, objectName: "区分の背景" });
  s.addText(tag, { x: 8.45, y: 0.17, w: 1.2, h: 0.3, fontSize: 11, bold: true, color: C.background1, align: "center", valign: "middle", margin: 0, isTextBox: true, objectName: "区分" });
  s.addText([{ text: "目的　", options: { bold: true, color: C.text2 } }, { text: purpose, options: { color: C.text1 } }],
    { x: 0.4, y: 0.72, w: 9.2, h: 0.32, fontSize: 12, margin: 0, valign: "middle", isTextBox: true, objectName: "目的" });
  s.addText([{ text: "方法　", options: { bold: true } }, { text: method }],
    { x: 0.4, y: 1.04, w: 9.2, h: 0.36, fontSize: 10, color: C.accent6, margin: 0, valign: "middle", isTextBox: true, objectName: "方法" });
  return s;
}

// 土日祝の帯（棒）と、稼働時間の線を重ねたグラフ
function trend(s, series, colors, box, opts = {}) {
  const labels = D.labels;
  const band = { type: pres.charts.BAR, data: [{ name: "土日祝", labels, values: D.band }],
    options: { chartColors: [BAND], barGapWidthPct: 0 } };
  const lines = { type: pres.charts.LINE,
    data: series.map(([name, values]) => ({ name, labels, values })),
    options: { chartColors: colors, lineSize: opts.lineSize || 1.75, lineDataSymbol: "none" } };
  s.addChart([band, lines], {
    ...box,
    objectName: opts.name || "グラフ",
    displayBlanksAs: "gap",
    showTitle: !!opts.title, title: opts.title, titleFontSize: 10, titleColor: HEX.dk2, titleFontFace: "+mn-lt", titleAlign: "left",
    showLegend: true, legendPos: opts.legendPos || "t", legendFontSize: 8, legendFontFace: "+mn-lt", legendColor: HEX.dk1,
    catAxisLabelFrequency: 7, catAxisLabelRotate: 0, catAxisLabelFontSize: 8, catAxisLabelColor: HEX.accent6, catAxisLabelFontFace: "+mn-lt",
    catAxisHidden: !!opts.hideCat,
    catAxisLineShow: true, catAxisLineColor: "C9D1DB", catGridLine: { style: "none" },
    valAxisMinVal: 0, valAxisMaxVal: 9, valAxisMajorUnit: 3,
    valAxisLabelFontSize: 8, valAxisLabelColor: HEX.accent6, valAxisLabelFontFace: "+mn-lt",
    valAxisTitle: opts.yTitle || "時間（9時間中）", showValAxisTitle: opts.yTitle !== "", valAxisTitleFontSize: 8, valAxisTitleColor: HEX.accent6, valAxisTitleFontFace: "+mn-lt",
    valGridLine: { color: "E1E6EC", size: 0.5 },
  });
}

const fmt = (v) => (v === null || v === undefined ? "－" : v.toFixed(1));
const cellOpt = { fontSize: 9, color: HEX.dk1, align: "center", valign: "middle", border: { type: "solid", pt: 0.5, color: "C9D1DB" }, margin: 0.02 };
const headOpt = { ...cellOpt, bold: true, color: HEX.lt1, fill: { color: HEX.dk2 } };

function readBox(s, x, y, w, h, text) {
  s.addShape(pres.shapes.ROUNDED_RECTANGLE, { x, y, w, h, rectRadius: 0.06, fill: { color: "FFF7D1" }, line: { color: "FFF7D1" }, objectName: "読み取りの背景" });
  s.addText([{ text: "読み取り", options: { bold: true, color: HEX.accent4, breakLine: true } }, { text }],
    { x: x + 0.12, y: y + 0.05, w: w - 0.24, h: h - 0.1, fontSize: 10, color: C.text1, valign: "top", margin: 0, isTextBox: true, objectName: "読み取り" });
}

// ---------------------------------------------------------------- 1枚目：列ごとの推移
{
  const s = frame("室外機の稼働時間の推移（日中9〜18時）", "前提",
    "研究に使える時間（室外機が運転している時間）が、7/29以降どう変わったかを見る",
    "1時間ごとに「運転」（排気温度−吸気温度が1.0℃以上の分が8割以上）かを判定し、9〜18時の9時間のうち「運転」の時間を日ごとに数えた。各列3台の平均");
  trend(s, Object.keys(COLOR).map((k) => [k + "（3台の平均）", D.col[k]]), Object.values(COLOR),
    { x: 0.4, y: 1.42, w: 9.2, h: 2.72 }, { name: "列ごとの稼働時間" });
  const P = D.periods;
  const rows = [[{ text: "平日の平均［時間］", options: headOpt }, ...P.map((p) => ({ text: p, options: headOpt }))]];
  for (const k of Object.keys(COLOR)) {
    rows.push([{ text: k, options: { ...cellOpt, bold: true, color: COLOR[k] } },
      ...D.summary.col[k].map((r) => ({ text: r.mean === null ? "データなし" : fmt(r.mean), options: cellOpt }))]);
  }
  s.addTable(rows, { x: 0.4, y: 4.24, w: 5.3, colW: [1.5, 0.95, 0.95, 0.95, 0.95], rowH: 0.23, objectName: "期間別の平均" });
  const c = D.summary.col;
  readBox(s, 5.85, 4.24, 3.75, 0.95,
    `9月後半から短くなった。10月の平日は8月と比べ、列3は約半分（${fmt(c["列3"][0].mean)}→${fmt(c["列3"][3].mean)}時間）、` +
    `列8・列11は約2割短い（${fmt(c["列8"][0].mean)}→${fmt(c["列8"][3].mean)}、${fmt(c["列11"][0].mean)}→${fmt(c["列11"][3].mean)}時間）`);
  s.addNotes("灰色の帯は土日祝。線が途切れている日は、判定できた時間が7時間未満の日（8-2・3-2の排気温度の欠測、列11のデータなし（9/15〜10/1）、7/29、10/13）。" +
    "列の値は、その日に値がある室外機だけで平均した。平日の平均は8月＝7/30〜8/31、9月前半＝9/1〜9/15、9月後半＝9/16〜9/30、10月＝10/1〜10/13。");
}

// ---------------------------------------------------------------- 2枚目：室外機ごとの推移
{
  const s = frame("室外機ごとの稼働時間（日中9〜18時）", "前提",
    "どの室外機の稼働時間が、いつから減ったかを見る",
    "9〜18時の9時間のうち「運転」の時間（判定は1枚目と同じ）。判定できた時間が7時間未満の日は空欄にした（線が途切れる）");
  const groups = [["列3（芋緑化なし）", ["3-1", "3-2", "3-3"]], ["列8（芋緑化あり・中央）", ["8-1", "8-2", "8-3"]], ["列11（芋緑化あり・外側）", ["11-1", "11-2", "11-3"]]];
  const ys = [1.42, 2.62, 3.82], hs = [1.12, 1.12, 1.38];
  groups.forEach(([title, us], i) => {
    s.addText(title, { x: 0.4, y: ys[i], w: 4, h: 0.2, fontSize: 9.5, bold: true, color: C.text2, margin: 0, isTextBox: true, objectName: `見出し ${title}` });
    trend(s, us.map((u) => [u, D.unit[u]]), us.map((u) => UNIT_COLOR[u]),
      { x: 0.4, y: ys[i] + 0.18, w: 9.2, h: hs[i] - 0.18 },
      { name: title, legendPos: "r", yTitle: "", lineSize: 1.25, hideCat: i < 2 });
  });
  s.addNotes("縦軸は9時間中の「運転」の時間（0〜9時間）。灰色の帯は土日祝。8-2は8/8〜8/14に日中も停止、8/14〜8/28は排気温度の欠測。" +
    "3-2は8/30〜9/7、9/14〜9/22ごろに排気温度の欠測。列11は9/15〜10/1のデータがない。");
}

// ---------------------------------------------------------------- 3枚目：研究に使える時間
{
  const s = frame("研究に使える時間の推移（平日、日中9〜18時）", "前提",
    "比べる2台がともに運転していて、分析に使える時間がどう変わったかを見る",
    "組（芋緑化ありの室外機と、同じ区画の列3）の2台がともに「運転」の時間を、平日だけ日ごとに数えた。列8・列11は3組の平均。土日祝は分析に使わない");
  trend(s, [["列8（3組の平均）", D.paircol["列8"]], ["列11（3組の平均）", D.paircol["列11"]]], [COLOR["列8"], COLOR["列11"]],
    { x: 0.4, y: 1.44, w: 9.2, h: 2.06 }, { name: "研究に使える時間" });
  const P = D.periods;
  const h1 = [{ text: "組", options: { ...headOpt, rowspan: 2 } },
    { text: "平日の平均［時間］", options: { ...headOpt, colspan: 4 } },
    { text: "7時間以上の平日［日間］", options: { ...headOpt, colspan: 4 } }];
  const h2 = [...P, ...P].map((p) => ({ text: p, options: headOpt }));
  const rows = [h1, h2];
  for (const [name, r] of Object.entries(D.summary.pair)) {
    rows.push([{ text: name, options: { ...cellOpt, bold: true } },
      ...r.map((x) => ({ text: x.mean === null ? "なし" : fmt(x.mean), options: cellOpt })),
      ...r.map((x) => ({ text: x.days7 === null ? "なし" : String(x.days7), options: cellOpt }))]);
  }
  const small = (r) => r.map((c) => ({ ...c, options: { ...c.options, fontSize: 8.5 } }));
  s.addTable(rows.map(small), { x: 0.4, y: 3.58, w: 6.3, colW: [1.06, 0.655, 0.655, 0.655, 0.655, 0.655, 0.655, 0.655, 0.655], rowH: 0.188, objectName: "組ごとの期間別の平均" });
  const pc = D.summary.paircol;
  readBox(s, 6.85, 3.58, 2.75, 1.5,
    `2台とも運転している時間は、8月の平日は1日${fmt(Math.min(pc["列8"][0].mean, pc["列11"][0].mean))}〜${fmt(Math.max(pc["列8"][0].mean, pc["列11"][0].mean))}時間（3組平均）、` +
    `10月は${fmt(Math.min(pc["列8"][3].mean, pc["列11"][3].mean))}〜${fmt(Math.max(pc["列8"][3].mean, pc["列11"][3].mean))}時間。` +
    "日中平均の基準（7時間以上）を満たす日は、10月は各組0〜1日間");
  s.addNotes("「なし」は列11のデータがない期間（9/15〜10/1）。7時間以上の平日＝日中平均を使う分析（9時間中7時間以上）の対象になる日。");
}

(async () => {
  await pres.writeFile({ fileName: OUT });
  // 欠けた日（空の値）の点を消し、PowerPoint で空欄（線が途切れる）として扱われるようにする
  const zip = await JSZip.loadAsync(fs.readFileSync(OUT));
  for (const name of Object.keys(zip.files).filter((n) => /^ppt\/charts\/chart\d+\.xml$/.test(n))) {
    const xml = await zip.file(name).async("string");
    zip.file(name, xml.replace(/<c:pt idx="\d+"><c:v><\/c:v><\/c:pt>/g, ""));
  }
  fs.writeFileSync(OUT, await zip.generateAsync({ type: "nodebuffer", compression: "DEFLATE" }));
  const { applyTheme } = require(path.join(process.env.PPTX_SKILL, "scripts", "apply_theme.js"));
  await applyTheme(OUT, THEME);
  console.log("written", OUT);
})();
