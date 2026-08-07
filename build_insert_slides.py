# -*- coding: utf-8 -*-
"""既存デッキ 成田_短時間勤務者採用_説明会.pptx のスライド8の後に、
朝集中の実データ4〜6枚を差し込む。すべて実線表データ・推測なし。"""
from pptx import Presentation
from pptx.util import Inches, Pt, Emu
from pptx.dml.color import RGBColor
from pptx.enum.text import PP_ALIGN, MSO_ANCHOR
from pptx.enum.shapes import MSO_SHAPE
from pptx.chart.data import CategoryChartData
from pptx.enum.chart import XL_CHART_TYPE
from pptx.oxml.ns import qn

def C(h): return RGBColor.from_string(h)
NAVY, TEXT, BLUE, GREY, GREY2 = C("1B2A4A"), C("1F2733"), C("2E5AAC"), C("6B7686"), C("9AA6B8")
RED, CARD, LINE, LBLUE, GREEN = C("C0392B"), C("F3F4F7"), C("D9DFEA"), C("E5ECF7"), C("E3F0EA")
AMBER, WHITE = C("B8891B"), C("FFFFFF")
FONT = "Meiryo"

prs = Presentation("成田_短時間勤務者採用_説明会.pptx")
BLANK = prs.slide_layouts[0]

def new_slide():
    s = prs.slides.add_slide(BLANK)
    for ph in list(s.placeholders):
        ph._element.getparent().remove(ph._element)
    return s

def box(s, l, t, w, h, text=None, size=14, color=TEXT, bold=False, align=PP_ALIGN.LEFT,
        anchor=MSO_ANCHOR.TOP, wrap=True, fill=None, line=None, shape=MSO_SHAPE.RECTANGLE,
        line_w=1.0, space_after=2):
    sp = s.shapes.add_shape(shape, Inches(l), Inches(t), Inches(w), Inches(h))
    if fill is None:
        sp.fill.background()
    else:
        sp.fill.solid(); sp.fill.fore_color.rgb = fill
    if line is None:
        sp.line.fill.background()
    else:
        sp.line.color.rgb = line; sp.line.width = Pt(line_w)
    sp.shadow.inherit = False
    tf = sp.text_frame; tf.word_wrap = wrap
    tf.vertical_anchor = anchor
    for m in ("margin_left", "margin_right"):
        setattr(tf, m, Inches(0.08))
    tf.margin_top = Inches(0.03); tf.margin_bottom = Inches(0.03)
    if text is not None:
        lines = text.split("\n") if isinstance(text, str) else text
        for i, ln in enumerate(lines):
            p = tf.paragraphs[0] if i == 0 else tf.add_paragraph()
            p.alignment = align; p.space_after = Pt(space_after); p.space_before = Pt(0)
            r = p.add_run(); r.text = ln
            r.font.name = FONT; r.font.size = Pt(size); r.font.bold = bold
            r.font.color.rgb = color
    return sp

def eyebrow(s, txt): box(s, 0.60, 0.55, 11.8, 0.35, txt, 13, BLUE, True)
def title(s, txt, size=30): box(s, 0.60, 0.92, 12.10, 1.0, txt, size, NAVY, True, anchor=MSO_ANCHOR.TOP)
def footer(s, page):
    box(s, 0.50, 7.02, 6.0, 0.30, "短時間勤務者採用 説明会", 9, GREY)
    box(s, 12.20, 7.02, 0.70, 0.30, str(page), 9, GREY, align=PP_ALIGN.RIGHT)

# ---- 実データ（線表）----
def hm(x): h, m = map(int, x.split(":")); return h*60+m
flights = [  # carrier, 便, 行先, STA, STD, CTR, band
    ("AM", "AM058/057", "MEX", "06:30", "09:35", "06:30", "朝"),
    ("VJ", "VJ822/823", "SGN", "07:40", "08:55", "05:55", "朝"),
    ("VN", "VN318/319", "DAD", "07:35", "09:00", "06:00", "朝"),
    ("VN", "VN310/311", "HAN", "07:35", "09:30", "06:30", "朝"),
    ("VJ", "VJ932/933", "HAN", "08:00", "09:30", "06:30", "朝"),
    ("VN", "VN306/307", "SGN", "08:00", "09:30", "06:30", "朝"),
    ("5J", "5J5062/063", "CEB", "08:10", "08:55", "05:55", "朝"),
    ("UL", "UL454/455", "CMB", "08:10", "11:15", "08:15", "朝"),
    ("VN", "VN316/317", "DAD", "08:30", "10:30", "07:30", "朝"),
    ("RF", "RF392/391", "CJJ", "09:40", "10:40", "08:10", "朝"),
    ("5J", "5J5068/069", "CRK", "10:25", "11:15", "08:15", "朝"),
    ("5J", "5J5054/055", "MNL", "12:25", "13:45", "10:45", "昼"),
    ("VJ", "VJ934/935", "HAN", "15:30", "16:30", "13:30", "夕"),
    ("5J", "5J5056/057", "MNL", "18:00", "19:15", "16:15", "夕"),
]
# 稼働ウィンドウ [min(CTR,STA), STD]
wins = [(min(hm(f[5]), hm(f[3])), hm(f[4]), f[6]) for f in flights]

# 30分刻み稼働便数
t0, t1 = hm("05:30"), hm("20:00")
labels, vals = [], []
t = t0
while t < t1:
    labels.append(f"{t//60:02d}:{t%60:02d}")
    vals.append(sum(1 for s, e, _ in wins if s <= t < e))
    t += 30

# =====================================================================
# SLIDE A ｜ 追加① 需要カーブ（面グラフ）
# =====================================================================
s = new_slide()
eyebrow(s, "背景⑤｜時間帯別の需要（実線表より）")
title(s, "成田の人員需要は、一日を通して一定ではない")
cd = CategoryChartData(); cd.categories = labels
cd.add_series("時間帯別 稼働便数", vals)
gf = s.shapes.add_chart(XL_CHART_TYPE.AREA, Inches(0.55), Inches(2.15), Inches(7.7), Inches(4.35), cd)
ch = gf.chart; ch.has_title = False; ch.has_legend = False
ser = ch.series[0]
ser.format.fill.solid(); ser.format.fill.fore_color.rgb = BLUE
ser.format.line.color.rgb = NAVY; ser.format.line.width = Pt(1.5)
for ax in (ch.category_axis, ch.value_axis):
    ax.tick_labels.font.size = Pt(9); ax.tick_labels.font.name = FONT
    ax.tick_labels.font.color.rgb = GREY
# ラベル間引き（2つ=1時間ごと）
catAx = ch.category_axis._element
for tag, v in (("c:tickLblSkip", "2"), ("c:tickMarkSkip", "2")):
    e = catAx.makeelement(qn(tag), {"val": v}); catAx.append(e)
# 右：ポイント＋キーメッセージ
px = 8.55
box(s, px, 2.20, 4.20, 0.72, ["朝　　大量の人員が必要", "CTR OPEN 05:55〜、到着が集中"], 13, TEXT, True, fill=LBLUE, line=LINE)
box(s, px, 3.02, 4.20, 0.72, ["12時以降　必要人員が急減", "夕方まで低水準が続く"], 13, TEXT, True, fill=CARD, line=LINE)
box(s, px, 3.84, 4.20, 0.72, ["一方 フルタイム社員は月166時間", "朝のピーク後も勤務時間が残る"], 13, TEXT, True, fill=CARD, line=LINE)
box(s, px, 4.85, 4.20, 1.05, "需要の形と、\n雇用の形は合っているか？", 17, WHITE, True,
    align=PP_ALIGN.CENTER, anchor=MSO_ANCHOR.MIDDLE, fill=NAVY, shape=MSO_SHAPE.ROUNDED_RECTANGLE)
box(s, 0.55, 6.55, 7.7, 0.30, "※ 成田空港事業部がハンドリングを担当する便の同時稼働数（CTR OPEN〜出発）。出典：週間線表。", 9, GREY2)
footer(s, 7)

# =====================================================================
# SLIDE B ｜ 追加①補足 便ごとのガント
# =====================================================================
s = new_slide()
eyebrow(s, "背景⑤｜補足：便ごとの地上作業帯")
title(s, "朝に便が集中している（便ごとの作業時間帯）")
GL, GT, GW, GH = 2.35, 2.35, 10.30, 4.05
ax_min, ax_max = hm("05:00"), hm("20:00")
def xin(m): return GL + (m - ax_min) / (ax_max - ax_min) * GW
# 時間目盛り（毎正時の縦線＋ラベル）
for hh in range(5, 21):
    m = hh*60; x = xin(m)
    ln = s.shapes.add_shape(MSO_SHAPE.RECTANGLE, Inches(x), Inches(GT), Emu(9144), Inches(GH))
    ln.fill.solid(); ln.fill.fore_color.rgb = LINE; ln.line.fill.background(); ln.shadow.inherit = False
    box(s, x-0.25, GT-0.30, 0.5, 0.28, f"{hh}", 9, GREY, align=PP_ALIGN.CENTER)
# 昼(12時)強調線
xn = xin(hm("12:00"))
hl = s.shapes.add_shape(MSO_SHAPE.RECTANGLE, Inches(xn), Inches(GT), Emu(15240), Inches(GH))
hl.fill.solid(); hl.fill.fore_color.rgb = GREY2; hl.line.fill.background(); hl.shadow.inherit = False
# バー（1行=1便、開始昇順で上から）
rows = sorted(range(len(flights)), key=lambda i: wins[i][0])
pitch = GH / len(flights)
for r, idx in enumerate(rows):
    f = flights[idx]; s0, e0, band = wins[idx]
    ry = GT + r*pitch + pitch*0.15
    rh = pitch*0.66
    col = BLUE if band == "朝" else (AMBER if band == "夕" else C("D08A1E"))
    box(s, 0.55, ry-0.02, 1.75, rh+0.04, f"{f[0]} {f[1]}", 9, TEXT, False, anchor=MSO_ANCHOR.MIDDLE)
    bar = s.shapes.add_shape(MSO_SHAPE.ROUNDED_RECTANGLE, Inches(xin(s0)), Inches(ry),
                             Inches(max(0.12, xin(e0)-xin(s0))), Inches(rh))
    bar.fill.solid(); bar.fill.fore_color.rgb = col; bar.line.fill.background(); bar.shadow.inherit = False
    bar.adjustments[0] = 0.35
box(s, GL, GT+GH+0.10, GW, 0.30, "各バー＝CTR OPEN〜出発（STD）。青＝朝便、橙＝昼・夕便。12時（灰線）以降はほぼ空く。", 10, GREY)
box(s, 0.55, 6.62, 12, 0.30, "出典：週間線表（成田空港事業部ハンドリング便）。UL＝火・木・土運航。", 9, GREY2)
footer(s, 8)

# =====================================================================
# SLIDE C ｜ 配置候補便テーブル
# =====================================================================
s = new_slide()
eyebrow(s, "現場での活用｜配置候補便（社長案）")
title(s, "短時間勤務者の配置候補便", 30)
box(s, 6.7, 1.05, 6.0, 0.7, "VN / AM / RF / 5J / UL / VJ", 20, BLUE, True, anchor=MSO_ANCHOR.MIDDLE)
heads = ["Carrier", "便名", "行先", "到着", "出発", "CTR OPEN", "時間帯", "配置"]
widths = [1.1, 1.8, 1.0, 1.0, 1.0, 1.3, 1.0, 3.0]
tbl_l, tbl_t = 0.55, 2.15
rows_n = len(flights) + 1
gtbl = s.shapes.add_table(rows_n, len(heads), Inches(tbl_l), Inches(tbl_t),
                          Inches(sum(widths)), Inches(4.35)).table
gtbl.first_row = False; gtbl.horz_banding = False
for j, w in enumerate(widths):
    gtbl.columns[j].width = Inches(w)
def setcell(cell, txt, size=10, color=TEXT, bold=False, fill=WHITE, align=PP_ALIGN.CENTER):
    cell.fill.solid(); cell.fill.fore_color.rgb = fill
    cell.margin_left = Inches(0.04); cell.margin_right = Inches(0.04)
    cell.margin_top = Inches(0.02); cell.margin_bottom = Inches(0.02)
    cell.vertical_anchor = MSO_ANCHOR.MIDDLE
    tf = cell.text_frame; tf.word_wrap = True
    p = tf.paragraphs[0]; p.alignment = align
    r = p.add_run(); r.text = txt; r.font.name = FONT; r.font.size = Pt(size)
    r.font.bold = bold; r.font.color.rgb = color
for j, h in enumerate(heads):
    setcell(gtbl.cell(0, j), h, 10, WHITE, True, NAVY)
place = {"朝": "◯ 朝ピーク配置", "昼": "△ 昼便", "夕": "△ 夕便"}
for i, f in enumerate(flights):
    band = f[6]
    rowfill = WHITE if band == "朝" else C("FBEAD2")
    pl = place[band]
    if f[1] == "VJ934/935": pl = "△ Winter条件付"
    disp = list(f[:6]) + [band, pl]
    for j, v in enumerate(disp):
        setcell(gtbl.cell(i+1, j), str(v), 10,
                TEXT, (j == 0), rowfill)
box(s, 0.55, 6.62, 12.2, 0.30,
    "◯＝朝ピークの配置候補（VN/AM/RF/5J/UL/VJ）。VJ934/935（夕方HAN）はWinter運休時は対象外、運航時は別途候補。時刻は線表の実データ。",
    9, GREY2)
footer(s, 9)

# =====================================================================
# SLIDE D ｜ 追加② 満席のレストラン
# =====================================================================
s = new_slide()
eyebrow(s, "経営者目線①")
title(s, "満席のレストランは、経営に成功しているのか？")
box(s, 0.60, 2.35, 5.7, 3.6, None, fill=GREEN, line=LINE)
box(s, 0.90, 2.55, 5.1, 0.5, "見えている姿", 15, C("2C6E49"), True)
for i, txt in enumerate(["満席（お客様でいっぱい）", "料理も好評", "サービスも高評価", "スタッフも一生懸命"]):
    box(s, 1.05, 3.15+i*0.62, 5.0, 0.5, "・ " + txt, 14, TEXT, False)
box(s, 6.55, 3.55, 0.9, 0.9, "→", 30, GREY2, True, align=PP_ALIGN.CENTER, anchor=MSO_ANCHOR.MIDDLE)
box(s, 7.55, 2.35, 5.15, 3.6, None, fill=C("F7E7E4"), line=C("E4B9B2"))
box(s, 7.85, 2.55, 4.6, 0.5, "しかし――", 15, RED, True)
box(s, 7.85, 3.55, 4.6, 1.2, "売上 ＜ 原価", 34, RED, True, align=PP_ALIGN.CENTER, anchor=MSO_ANCHOR.MIDDLE)
box(s, 7.85, 4.85, 4.6, 0.8, "良い材料・十分な人員で\n原価が売上を上回っていた", 13, TEXT, False, align=PP_ALIGN.CENTER)
box(s, 0.60, 6.15, 12.1, 0.75, "問い：これは経営として「成功」と言えるでしょうか？", 17, WHITE, True,
    align=PP_ALIGN.CENTER, anchor=MSO_ANCHOR.MIDDLE, fill=NAVY, shape=MSO_SHAPE.ROUNDED_RECTANGLE)
footer(s, 10)

# =====================================================================
# SLIDE E ｜ 追加③ 良いサービス≠良い経営
# =====================================================================
s = new_slide()
eyebrow(s, "経営者目線②")
title(s, "良いサービス ≠ 良い経営")
def col(lx, header, items, tail, hcolor, tcolor, tfill):
    box(s, lx, 2.35, 5.85, 0.6, header, 16, hcolor, True, fill=CARD, line=LINE)
    for i, it in enumerate(items):
        box(s, lx+0.15, 3.10+i*0.60, 5.55, 0.5, "・ " + it, 14, TEXT)
    box(s, lx, 5.55, 5.85, 0.75, tail, 16, tcolor, True, align=PP_ALIGN.CENTER,
        anchor=MSO_ANCHOR.MIDDLE, fill=tfill, line=LINE)
col(0.60, "【レストラン】", ["満席", "料理がおいしい", "評判がいい", "十分なスタッフ"],
    "しかし 売上 ＜ 原価", C("2C6E49"), RED, C("F7E7E4"))
col(6.85, "【空港】", ["安定したオペレーション", "十分な人員配置", "高いサービス品質", "現場の努力"],
    "しかし 収入 ＜ 原価", BLUE, RED, C("F7E7E4"))
box(s, 0.60, 6.55, 12.1, 0.42, "良いサービスを提供しているだけでは、事業は継続できない。", 15, NAVY, True, align=PP_ALIGN.CENTER)
footer(s, 11)

# =====================================================================
# SLIDE F ｜ 追加④ 売上を増やすだけを待てない
# =====================================================================
s = new_slide()
eyebrow(s, "経営者目線③")
title(s, "「売上を増やす」だけを、待つことはできない")
box(s, 0.60, 2.35, 5.85, 0.6, "売上を増やす", 16, BLUE, True, fill=LBLUE, line=LINE)
for i, it in enumerate(["営業強化", "新規受託", "増便", "単価改善"]):
    box(s, 0.75, 3.10+i*0.58, 5.55, 0.5, "・ " + it, 14, TEXT)
box(s, 0.60, 5.50, 5.85, 0.6, "→ 当然、必要（時間はかかる）", 14, TEXT, True, align=PP_ALIGN.CENTER,
    anchor=MSO_ANCHOR.MIDDLE, fill=CARD, line=LINE)
box(s, 6.85, 2.35, 5.85, 0.6, "原価を適正化する", 16, C("2C6E49"), True, fill=GREEN, line=LINE)
for i, it in enumerate(["必要人数の見直し", "時間帯別の配置", "雇用形態の組み合わせ", "短時間勤務者の活用"]):
    box(s, 7.00, 3.10+i*0.58, 5.55, 0.5, "・ " + it, 14, TEXT)
box(s, 6.85, 5.50, 5.85, 0.6, "→ 今すぐ、自分たちでできる", 14, C("2C6E49"), True, align=PP_ALIGN.CENTER,
    anchor=MSO_ANCHOR.MIDDLE, fill=GREEN, line=LINE)
box(s, 0.60, 6.45, 12.1, 0.55, "営業と原価改善は、どちらか一方ではない。両方やる。", 16, WHITE, True,
    align=PP_ALIGN.CENTER, anchor=MSO_ANCHOR.MIDDLE, fill=NAVY, shape=MSO_SHAPE.ROUNDED_RECTANGLE)
footer(s, 12)

# =====================================================================
# 既存フッターのページ番号を +6（7以降）
# =====================================================================
for i in range(20):  # 元の20枚のみ
    sl = prs.slides[i]
    for sh in sl.shapes:
        if sh.has_text_frame:
            tf = sh.text_frame; txt = tf.text.strip()
            if txt.isdigit() and Emu(sh.left).inches > 11:
                n = int(txt)
                if n >= 7:
                    for p in tf.paragraphs:
                        for r in p.runs:
                            if r.text.strip().isdigit():
                                r.text = str(n+6)

# =====================================================================
# 並べ替え：新6枚を idx7（背景④）の直後へ
# =====================================================================
sldIdLst = prs.slides._sldIdLst
ids = list(sldIdLst)
new_ids = ids[20:26]
for e in new_ids:
    sldIdLst.remove(e)
ref = ids[8]  # 元 slide9（人員の持ち方）
for e in new_ids:
    ref.addprevious(e)

prs.save("out_deck.pptx")
print("saved out_deck.pptx ; total slides:", len(prs.slides._sldIdLst))
