# -*- coding: utf-8 -*-
"""背景①〜④の金額を「百万円」→「億・万円」表記（例:1億7,800万円）へ。
数字に不慣れな人向けの読み上げ表記。画面上の恒等式(売上-原価=総損益 / 人件費+派遣=合計)を保つ。"""
from pptx import Presentation
from pptx.chart.data import CategoryChartData

prs = Presentation("成田_短時間勤務者採用_説明会.pptx")

# 本文ラン単位の置換（現在は百万円表記）
REPL = [
    # --- 特定（長い）パターンを先に。'▲17.3' の総称置換より前に置く ---
    ("（単位：百万円）", "（金額は億・万円で表記）"),
    ("▲17.3百万円 × 12ヶ月", "▲1,730万円 × 12ヶ月"),
    ("売上 178.1  −  売上原価 182.8  =  ", "売上 1億7,800万  −  売上原価 1億8,300万  =  "),
    ("売上総損益 ▲4.8百万円", "売上総損益 ▲500万円"),
    ("販管費（12.6百万円）", "販管費（1,300万円）"),
    ("107.5百万円", "1億700万円"),
    ("54.7百万円", "5,500万円"),
    ("162.1百万円", "1億6,200万円"),
    ("営業損益（百万円/月）", "営業損益（万円/月）"),
    ("（▲58.0百万円/月）", "（▲5,800万円/月）"),
    ("約 ▲2.1", "約 ▲2億1,000"),
    # --- 総称（短い）パターンは後に ---
    ("▲17.3", "▲1,730"),            # slide5/7 の単独コールアウト数値
    (" 百万円 / 月", "万円 / 月"),
    ("百万円 / 月", "万円 / 月"),
    (" 億円", "万円"),
]

for idx in (4, 5, 6, 7):
    for sh in prs.slides[idx].shapes:
        if sh.has_text_frame:
            for para in sh.text_frame.paragraphs:
                for r in para.runs:
                    for old, new in REPL:
                        if old in r.text:
                            r.text = r.text.replace(old, new)

# slide7 グラフ: 百万円 → 万円 (×100)
for sh in prs.slides[6].shapes:
    if sh.has_chart:
        ch = sh.chart
        cats = [c for c in ch.plots[0].categories]
        new_vals = [round(v*100) for v in ch.series[0].values]
        cd = CategoryChartData(); cd.categories = cats
        cd.add_series(ch.series[0].name or "旅客 営業損益", new_vals)
        ch.replace_data(cd)
        print("chart→万円:", list(zip(cats, new_vals)))

# スピーカーノートも億・万円に統一
NOTE_REPL = {
    4: [("売上178百万に対して売上原価が182百万", "売上1億7,800万に対して売上原価が1億8,300万"),
        ("営業損益では月17百万円の赤字", "営業損益では月1,730万円の赤字"),
        ("年に直すと約2.1億円の赤字", "年に直すと約2億1,000万円の赤字")],
    7: [("営業赤字はおよそ月58百万円", "営業赤字はおよそ月5,800万円")],
}
for idx, reps in NOTE_REPL.items():
    s = prs.slides[idx]
    if s.has_notes_slide:
        tf = s.notes_slide.notes_text_frame
        for para in tf.paragraphs:
            for r in para.runs:
                for old, new in reps:
                    if old in r.text:
                        r.text = r.text.replace(old, new)

prs.save("out_deck.pptx")
print("saved")
