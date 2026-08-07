# -*- coding: utf-8 -*-
"""背景①〜④の金額表記を「千円(6桁)」→「百万円」に統一し、読み上げ誤りを防ぐ。
値は ÷1000。年換算のみ従来どおり億円。slide7のグラフ値も百万円へスケール。"""
from pptx import Presentation
from pptx.chart.data import CategoryChartData

prs = Presentation("成田_短時間勤務者採用_説明会.pptx")

# (old, new) 具体的置換。順序重要：長い/特定パターンを先に。
REPL = [
    ("（単位：千円）", "（単位：百万円）"),
    ("▲17,310千円 × 12ヶ月", "▲17.3百万円 × 12ヶ月"),
    ("売上 178,068  −  売上原価 182,818  =  ", "売上 178.1  −  売上原価 182.8  =  "),
    ("売上総損益 ▲4,751千円", "売上総損益 ▲4.8百万円"),
    ("販管費（12,560千円）", "販管費（12.6百万円）"),
    ("107,460千円", "107.5百万円"),
    ("54,675千円", "54.7百万円"),
    ("162,135千円", "162.1百万円"),
    ("（▲57,970千円/月）", "（▲58.0百万円/月）"),
    ("（千円/月）", "（百万円/月）"),
    ("▲17,310", "▲17.3"),        # 残る単独の大きな数字（slide5/7 コールアウト）
    (" 千円 / 月", " 百万円 / 月"),
    ("千円 / 月", "百万円 / 月"),
    ("千円/月", "百万円/月"),
]

for idx in (4, 5, 6, 7):  # 背景①〜④
    for sh in prs.slides[idx].shapes:
        if sh.has_text_frame:
            for para in sh.text_frame.paragraphs:
                for r in para.runs:
                    for old, new in REPL:
                        if old in r.text:
                            r.text = r.text.replace(old, new)

# slide7 の棒グラフ値を ÷1000（百万円）にスケール
s7 = prs.slides[6]
for sh in s7.shapes:
    if sh.has_chart:
        ch = sh.chart
        cats = [c for c in ch.plots[0].categories]
        old_vals = list(ch.series[0].values)
        new_vals = [round(v/1000, 1) for v in old_vals]
        cd = CategoryChartData()
        cd.categories = cats
        cd.add_series(ch.series[0].name or "旅客 営業損益", new_vals)
        ch.replace_data(cd)
        print("chart scaled:", list(zip(cats, new_vals)))

prs.save("out_deck.pptx")
print("saved")
