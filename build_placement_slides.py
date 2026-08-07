# -*- coding: utf-8 -*-
"""朝の短時間シフト割付を2枚、既存デッキに差し込む（課題整理の後・お願いの前）。
実線表データに基づく。必要人数は断定せず支店確定の建て付け。"""
import json
from pptx import Presentation
from pptx.util import Inches, Pt, Emu
from pptx.dml.color import RGBColor
from pptx.enum.text import PP_ALIGN, MSO_ANCHOR
from pptx.enum.shapes import MSO_SHAPE

def C(h): return RGBColor.from_string(h)
NAVY, TEXT, BLUE, GREY, GREY2 = C("1B2A4A"), C("1F2733"), C("2E5AAC"), C("6B7686"), C("9AA6B8")
RED, CARD, LINE, LBLUE, GREEN = C("C0392B"), C("F3F4F7"), C("D9DFEA"), C("E5ECF7"), C("E3F0EA")
GREENTX, AMBER, WHITE = C("2C6E49"), C("D08A1E"), C("FFFFFF")
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
    if fill is None: sp.fill.background()
    else: sp.fill.solid(); sp.fill.fore_color.rgb = fill
    if line is None: sp.line.fill.background()
    else: sp.line.color.rgb = line; sp.line.width = Pt(line_w)
    sp.shadow.inherit = False
    tf = sp.text_frame; tf.word_wrap = wrap; tf.vertical_anchor = anchor
    tf.margin_left = Inches(0.08); tf.margin_right = Inches(0.08)
    tf.margin_top = Inches(0.03); tf.margin_bottom = Inches(0.03)
    if text is not None:
        for i, ln in enumerate(text.split("\n") if isinstance(text, str) else text):
            p = tf.paragraphs[0] if i == 0 else tf.add_paragraph()
            p.alignment = align; p.space_after = Pt(space_after); p.space_before = Pt(0)
            r = p.add_run(); r.text = ln
            r.font.name = FONT; r.font.size = Pt(size); r.font.bold = bold; r.font.color.rgb = color
    return sp

def eyebrow(s, t): box(s, 0.60, 0.55, 11.8, 0.35, t, 13, BLUE, True)
def title(s, t, size=30): box(s, 0.60, 0.92, 12.10, 1.0, t, size, NAVY, True)
def footer(s, page):
    box(s, 0.50, 7.02, 6.0, 0.30, "短時間勤務者採用 説明会", 9, GREY)
    box(s, 12.20, 7.02, 0.70, 0.30, str(page), 9, GREY, align=PP_ALIGN.RIGHT)

def hm(x): h, m = map(int, x.split(":")); return h*60+m

# =====================================================================
# 割付① シフト設計 ＋ タイムライン
# =====================================================================
s = new_slide()
eyebrow(s, "現場での活用｜短時間シフトの設計")
title(s, "朝の山を「早番・中番」の2本で受ける")
box(s, 0.60, 2.10, 6.00, 1.12,
    ["早番A　05:45 – 10:15（4.5h）", "開設波（カウンター開設）＋到着ピーク前半"], 15, WHITE, True,
    anchor=MSO_ANCHOR.MIDDLE, fill=BLUE, shape=MSO_SHAPE.ROUNDED_RECTANGLE)
box(s, 6.75, 2.10, 6.00, 1.12,
    ["中番B　07:30 – 12:00（4.5h）", "到着ピーク＋10:30以降の後半出発"], 15, WHITE, True,
    anchor=MSO_ANCHOR.MIDDLE, fill=GREENTX, shape=MSO_SHAPE.ROUNDED_RECTANGLE)

# --- タイムライン 05:00–12:00 ---
GL, GR = 1.35, 12.55
amin, amax = hm("05:00"), hm("12:00")
def xin(m): return GL + (m - amin) / (amax - amin) * (GR - GL)
axis_y = 5.30
# 時目盛
for hh in range(5, 13):
    x = xin(hh*60)
    box(s, x-0.0, 3.95, 0.01, 1.9, None, fill=LINE)  # 縦グリッド
    box(s, x-0.25, 3.68, 0.5, 0.25, f"{hh}", 9, GREY, align=PP_ALIGN.CENTER)
# 波マーカー
def marker(mm, label, col):
    x = xin(mm)
    box(s, x-0.008, 3.95, 0.016, 1.9, None, fill=col)
    box(s, x-1.0, 3.95, 2.0, 0.26, label, 9, col, True, align=PP_ALIGN.CENTER)
marker(hm("05:55"), "開設波 05:55", GREY2)
marker(hm("08:30"), "到着ピーク 08:30・同時11便", RED)
marker(hm("11:15"), "後半出発 〜11:15", GREY2)
# 帯A / 帯B
box(s, xin(hm("05:45")), 4.35, xin(hm("10:15"))-xin(hm("05:45")), 0.55, "早番A", 12, WHITE, True,
    align=PP_ALIGN.CENTER, anchor=MSO_ANCHOR.MIDDLE, fill=BLUE, shape=MSO_SHAPE.ROUNDED_RECTANGLE)
box(s, xin(hm("07:30")), 5.05, xin(hm("12:00"))-xin(hm("07:30")), 0.55, "中番B", 12, WHITE, True,
    align=PP_ALIGN.CENTER, anchor=MSO_ANCHOR.MIDDLE, fill=GREENTX, shape=MSO_SHAPE.ROUNDED_RECTANGLE)
box(s, 0.60, 6.20, 12.1, 0.55, "重なる 07:30–10:15 が一番人手が要る時間。早番と中番の二枚重ねで厚くする。",
    15, NAVY, True, align=PP_ALIGN.CENTER, anchor=MSO_ANCHOR.MIDDLE, fill=LBLUE, line=LINE)
box(s, 0.60, 6.80, 12.1, 0.22, "※ シフト時刻は提案値（実線表の2波に合わせて設定）。各シフトの必要人数は支店が線表×SLAで確定。", 9, GREY2)
footer(s, 20)

# =====================================================================
# 割付② キャリア×時間帯 マトリクス
# =====================================================================
s = new_slide()
eyebrow(s, "現場での活用｜配置の割付（人事部案）")
title(s, "どのキャリアを、どの時間帯に")
heads = ["キャリア", "早番A (05:45–10:15)", "中番B (07:30–12:00)", "主な便・備考"]
widths = [2.0, 3.2, 3.2, 3.8]
rows = [
    ("AM", "◯", "",  "AM058/057（着06:30・発09:35）ワイド"),
    ("VN", "◯", "◯", "早番：VN318/310/306　中番：VN316/317"),
    ("VJ", "◯", "",  "VJ822・VJ932（夕方便はWinter条件付）"),
    ("5J", "◯", "◯", "早番：5J5062　中番：5J5068（昼・夕は別枠）"),
    ("UL", "",  "◯", "UL454/455（08:15–11:15）※火木土・ワイド"),
    ("RF", "",  "◯", "RF392/391（08:10–10:40）"),
]
tl, tt = 0.60, 2.25
tbl = s.shapes.add_table(len(rows)+1, len(heads), Inches(tl), Inches(tt),
                         Inches(sum(widths)), Inches(3.9)).table
tbl.first_row = False; tbl.horz_banding = False
for j, w in enumerate(widths): tbl.columns[j].width = Inches(w)
def cell(cl, txt, size=12, color=TEXT, bold=False, fill=WHITE, align=PP_ALIGN.CENTER):
    cl.fill.solid(); cl.fill.fore_color.rgb = fill
    cl.margin_left = Inches(0.06); cl.margin_right = Inches(0.06)
    cl.margin_top = Inches(0.03); cl.margin_bottom = Inches(0.03)
    cl.vertical_anchor = MSO_ANCHOR.MIDDLE
    p = cl.text_frame.paragraphs[0]; p.alignment = align
    r = p.add_run(); r.text = txt; r.font.name = FONT; r.font.size = Pt(size)
    r.font.bold = bold; r.font.color.rgb = color
for j, h in enumerate(heads): cell(tbl.cell(0, j), h, 12, WHITE, True, NAVY)
for i, (car, a, b, memo) in enumerate(rows):
    cell(tbl.cell(i+1, 0), car, 13, NAVY, True, CARD)
    cell(tbl.cell(i+1, 1), a, 15, BLUE, True, LBLUE if a else WHITE)
    cell(tbl.cell(i+1, 2), b, 15, GREENTX, True, GREEN if b else WHITE)
    cell(tbl.cell(i+1, 3), memo, 11, TEXT, False, WHITE, PP_ALIGN.LEFT)
box(s, 0.60, 6.35, 12.1, 0.55, "早番A＝AM・VN・VJ・5J（前半便）／ 中番B＝UL・RF・VN(316)・5J(5068)。6社が2本のシフトに収まる。",
    14, WHITE, True, align=PP_ALIGN.CENTER, anchor=MSO_ANCHOR.MIDDLE, fill=NAVY, shape=MSO_SHAPE.ROUNDED_RECTANGLE)
box(s, 0.60, 6.96, 12.1, 0.22, "◯＝配置候補。時刻は線表の実データ。必要人数は各支店で確定。", 9, GREY2)
footer(s, 21)

# ---- スピーカーノート（セリフ）----
def notes(idx, paras):
    tf = prs.slides[idx].notes_slide.notes_text_frame; tf.clear()
    for i, p in enumerate(paras):
        (tf.paragraphs[0] if i == 0 else tf.add_paragraph()).text = p

n_design = [
    "ここからは、では実際に“どのキャリアの、どの時間帯に”短時間勤務者を入れていくか、という具体の話に入ります。対象はAM・VN・RF・5J・UL・VJの6社、すべて朝の便です。",
    "線表を見ると、朝の山は実は一つではありません。まず6時前後に、出発カウンターの開設が一斉に立ち上がる“開設の波”。そして7時半から9時半にかけて、到着と折り返しが重なる“到着の波”。ピークは8時半で、同時に11便が動いています。",
    "そこで短時間のシフトを二本立てで考えています。一つが早番、5時45分から10時15分。開設と到着の前半をカバー。もう一つが中番、7時半から12時。到着ピークと、10時半以降に残る後半の出発をカバーします。いずれも4時間半で、週20時間未満に収まります。",
    "重なる7時半から10時過ぎが一番人手が要る時間なので、早番と中番の二枚重ねでここを厚くする設計です。",
]
n_matrix = [
    "この二本に、6社を割り付けます。早番には、AM、そしてVN・VJ・5Jの前半便。だいたい9時半までに一巡します。中番には、UL、RF、それからVNの遅い便と5Jの遅い便。",
    "一点はっきりさせておきたいのは、私たちが今日お見せしているのは“何時に、どのキャリアが動くか”までだということです。“では各シフトに何人か”は、ここで人事が勝手に決めません。実際の必要人数は、皆さんの支店で、この線表とSLA上の必要配置を突き合わせて出していただきます。人事は、その受け皿として募集・時給・教育・受入を用意します。",
    "例外はVJです。VJには夕方の便もありますが、Winterで運休する場合は朝だけが対象。運航が残る場合は夕方帯を別途検討します。",
]
# 追加した2枚は末尾に付くので、そのインデックスを取得
total = len(prs.slides._sldIdLst)
notes(len(prs.slides)-2, n_design)
notes(len(prs.slides)-1, n_matrix)

# =====================================================================
# 既存フッターのページ番号 +2（20以降）※新2枚を追加する前に実施
# =====================================================================
# 新スライドは末尾。元スライドはインデックス 0..(total-3)
for i in range(total-2):
    sl = prs.slides[i]
    for sh in sl.shapes:
        if sh.has_text_frame:
            tf = sh.text_frame; txt = tf.text.strip()
            if txt.isdigit() and Emu(sh.left).inches > 11 and int(txt) >= 20:
                for p in tf.paragraphs:
                    for r in p.runs:
                        if r.text.strip().isdigit():
                            r.text = str(int(txt)+2)

# =====================================================================
# 並べ替え：新2枚を「現在整理している課題」(元idx20)の直後・「お願い」(元idx21)の前へ
# =====================================================================
sldIdLst = prs.slides._sldIdLst
ids = list(sldIdLst)
new_ids = ids[-2:]
for e in new_ids: sldIdLst.remove(e)
ref = ids[21]  # お願い（支店長・所属長の皆さんへ）
for e in new_ids: ref.addprevious(e)

prs.save("out_deck.pptx")
print("saved; slides=", len(prs.slides._sldIdLst))
