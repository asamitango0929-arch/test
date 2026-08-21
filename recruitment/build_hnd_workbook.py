# -*- coding: utf-8 -*-
"""HND MMチーム 人員推移資料（稟議添付用）を整形して出力する。"""
from openpyxl import Workbook
from openpyxl.styles import Font, PatternFill, Alignment, Border, Side
from openpyxl.utils import get_column_letter
from openpyxl.formatting.rule import CellIsRule
from openpyxl.chart import LineChart, BarChart, Reference
from openpyxl.comments import Comment

OUT = "/tmp/claude-0/-home-user-test/159a94d1-fe56-5981-a3d6-2b440c9c3e90/scratchpad/HND_人員推移_稟議添付資料.xlsx"

JP = "游ゴシック"
NAVY    = "1F3864"
NAVY_L  = "2F5597"
BAND    = "D9E2F3"
LABEL   = "F2F5FA"
TOTAL   = "E7EDF7"
RED     = "C00000"
REDFILL = "FCE0E0"
GREEN   = "1F7A3D"
GRAY    = "808080"
LINE    = "B4C6E7"

thin  = Side(style="thin",   color=LINE)
med   = Side(style="medium", color=NAVY)

def F(sz=10, b=False, color="000000"):
    return Font(name=JP, size=sz, bold=b, color=color)

def fill(c):
    return PatternFill("solid", start_color=c, end_color=c)

CENTER = Alignment(horizontal="center", vertical="center")
LEFTV  = Alignment(horizontal="left",   vertical="center")
LEFTI  = Alignment(horizontal="left",   vertical="center", indent=1)
RIGHTV = Alignment(horizontal="right",  vertical="center")
WRAP   = Alignment(horizontal="left",   vertical="top", wrap_text=True)

NUM   = '#,##0;[Red]"▲"#,##0;"−"'      # 人数（負は▲赤）
PLUS  = '"+"#,##0;[Red]"▲"#,##0;"±0"'  # 差異・過不足
PCT   = '0.0%'
MEI   = '#,##0.0"名"'

MONTHS = ["9月","10月","11月","12月","1月","2月","3月","4月","5月","6月","7月","8月"]
COLS   = [get_column_letter(i) for i in range(2, 14)]   # B..M
NCOL   = "N"

wb = Workbook()

# ============================================================
# シート2（先に作成）：月次推移
# ============================================================
ws = wb.active
ws.title = "月次推移"
ws.sheet_view.showGridLines = False

ws.column_dimensions["A"].width = 27
for c in COLS:
    ws.column_dimensions[c].width = 6.6
ws.column_dimensions[NCOL].width = 10.5

# --- タイトル ---
ws.merge_cells("A1:N1")
ws["A1"] = "HND MMチーム　人員推移（2025年9月〜2026年8月）"
ws["A1"].font = F(15, True, NAVY)
ws["A1"].alignment = Alignment(horizontal="left", vertical="center")
ws.row_dimensions[1].height = 30

ws.merge_cells("A2:N2")
ws["A2"] = "採用媒体の活用に関する稟議　添付資料　2/3　（月次明細）　／　単位：名"
ws["A2"].font = F(9, False, GRAY)
ws["A2"].alignment = Alignment(horizontal="left", vertical="center")
ws.row_dimensions[2].height = 16

for col in range(1, 15):                     # タイトル下のアクセントライン
    ws.cell(row=3, column=col).border = Border(bottom=Side(style="medium", color=NAVY))
ws.row_dimensions[3].height = 6

# --- 表ヘッダー（4-5行目） ---
ws.merge_cells("A4:A5"); ws["A4"] = "項　目"
ws.merge_cells("B4:F4"); ws["B4"] = "2025年"
ws.merge_cells("G4:M4"); ws["G4"] = "2026年"
ws.merge_cells("N4:N5"); ws["N4"] = "年間計"
for i, m in enumerate(MONTHS):
    ws.cell(row=5, column=2 + i, value=m)
for r in (4, 5):
    ws.row_dimensions[r].height = 20
    for col in range(1, 15):
        c = ws.cell(row=r, column=col)
        c.font = F(10, True, "FFFFFF")
        c.fill = fill(NAVY if r == 4 else NAVY_L)
        c.alignment = CENTER
        c.border = Border(left=thin, right=thin, top=med, bottom=med)

# --- 行定義 ---
# (種別, ラベル, 行番号)  band=区分見出し / in=入力 / fx=計算
ROWS = [
    ("band", "体　制", 6),
    ("in",   "必要数 (a)", 7),
    ("fxs",  "月初稼働数 (b)", 8),
    ("fx",   "差異 (b − a)", 9),
    ("band", "入（増員）", 10),
    ("in",   "採用数（正社員）", 11),
    ("in",   "採用数（派遣）", 12),
    ("in",   "異動（入）", 13),
    ("tot",  "入 合計", 14),
    ("band", "出（減員）", 15),
    ("in",   "退職数（正社員）", 16),
    ("in",   "退職数（派遣）", 17),
    ("in",   "長期欠勤等による離脱者数", 18),
    ("in",   "異動（出）", 19),
    ("tot",  "出 合計", 20),
    ("band", "結　果", 21),
    ("fxs",  "月末稼働数 (c)", 22),
    ("fx",   "過不足 (c − a)", 23),
]

DATA = {
    7:  [20, 25, 25, 25, 25, 25, 25, 25, 25, 25, 25, 25],
    11: [0, 4, 6, 4, 1, 3, 0, 3, 0, 0, 0, 0],
    12: [1, 0, 1, 0, 0, 0, 0, 0, 0, 0, 0, 0],
    13: [0] * 12,
    16: [2, 1, 1, 1, 3, 3, 1, 1, 0, 1, 2, 1],
    17: [0, 0, 1, 1, 0, 0, 0, 0, 0, 0, 0, 0],
    18: [0, 0, 0, 0, 0, 0, 0, 2, 0, 0, 0, 2],
    19: [0] * 12,
}

for kind, label, r in ROWS:
    ws.row_dimensions[r].height = 19 if kind != "band" else 18
    if kind == "band":
        ws.merge_cells(f"A{r}:N{r}")
        c = ws[f"A{r}"]
        c.value = label
        c.font = F(10, True, NAVY)
        c.alignment = LEFTI
        for col in range(1, 15):
            cc = ws.cell(row=r, column=col)
            cc.fill = fill(BAND)
            cc.border = Border(left=(med if col == 1 else None),
                               right=(med if col == 14 else None),
                               top=thin, bottom=thin)
        continue

    a = ws.cell(row=r, column=1, value=label)
    a.font = F(10, True if kind == "tot" else False)
    a.alignment = LEFTI
    a.fill = fill(TOTAL if kind == "tot" else LABEL)
    a.border = Border(left=med, right=thin, top=thin, bottom=thin)

    for i, col in enumerate(COLS):
        c = ws[f"{col}{r}"]
        c.font = F(10, kind == "tot")
        c.alignment = CENTER
        c.border = Border(left=thin, right=thin, top=thin, bottom=thin)
        if kind == "tot":
            c.fill = fill(TOTAL)
        if r in DATA:
            c.value = DATA[r][i]
            c.number_format = NUM
            c.font = F(10, False, "0000FF")          # 入力値は青字
        elif r == 8:                                  # 月初稼働数
            if i == 0:
                c.value = 20
                c.number_format = NUM
                c.font = F(10, False, "0000FF")
            else:
                c.value = f"={COLS[i-1]}22"
                c.number_format = NUM
        elif r == 9:
            c.value = f"={col}8-{col}7"; c.number_format = PLUS
        elif r == 14:
            c.value = f"=SUM({col}11:{col}13)"; c.number_format = NUM
        elif r == 20:
            c.value = f"=SUM({col}16:{col}19)"; c.number_format = NUM
        elif r == 22:
            c.value = f"={col}8+{col}14-{col}20"; c.number_format = NUM
        elif r == 23:
            c.value = f"={col}22-{col}7"; c.number_format = PLUS

    n = ws[f"{NCOL}{r}"]
    n.alignment = CENTER
    n.border = Border(left=thin, right=med, top=thin, bottom=thin)
    n.fill = fill(TOTAL)
    if r in (11, 12, 13, 14, 16, 17, 18, 19, 20):
        n.value = f"=SUM(B{r}:M{r})"
        n.number_format = NUM
        n.font = F(10, True, NAVY)
    else:
        n.value = "−"
        n.font = F(10, False, GRAY)

# 表の外枠（下辺）
for col in range(1, 15):
    c = ws.cell(row=23, column=col)
    b = c.border
    c.border = Border(left=b.left, right=b.right, top=b.top, bottom=med)

# 稼働数行を強調
for r in (8, 22):
    ws.cell(row=r, column=1).font = F(10, True)
    for col in COLS + [NCOL]:
        ws[f"{col}{r}"].font = Font(name=JP, size=10, bold=True,
                                    color=("0000FF" if (r == 8 and col == "B") else
                                           (GRAY if col == NCOL else "000000")))

# 差異・過不足のマイナスを赤網掛け
for r in (9, 23):
    ws.conditional_formatting.add(
        f"B{r}:M{r}",
        CellIsRule(operator="lessThan", formula=["0"],
                   fill=fill(REDFILL), font=Font(name=JP, size=10, bold=True, color=RED)))

ws.freeze_panes = "B6"

# --- 注記 ---
notes = [
    "【注記】",
    "1. 数値の単位は「名」。青字セルは実績入力値、黒字セルは計算式による自動算出。マイナスは「▲」で表示。",
    "2. 「月初稼働数(b)」は前月の「月末稼働数(c)」を引き継ぐ（2025年9月のみ実績入力）。",
    "3. 「月末稼働数(c)」＝ 月初稼働数(b) ＋ 入 合計 − 出 合計。",
    "4. 「長期欠勤等による離脱者数」は在籍のままラインに入れない人数を指し、退職数には含まない。",
    "5. 出典：HND MMチーム 人員推移管理表（2025年9月〜2026年8月実績）。",
]
for i, t in enumerate(notes):
    r = 25 + i
    ws.merge_cells(f"A{r}:N{r}")
    c = ws[f"A{r}"]
    c.value = t
    c.font = F(9, i == 0, NAVY if i == 0 else "404040")
    c.alignment = LEFTV
    ws.row_dimensions[r].height = 15

# --- 印刷設定 ---
ws.page_setup.orientation = "landscape"
ws.page_setup.paperSize = ws.PAPERSIZE_A4
ws.page_setup.fitToWidth = 1
ws.page_setup.fitToHeight = 1
ws.sheet_properties.pageSetUpPr.fitToPage = True
ws.print_options.horizontalCentered = True
ws.page_margins.left = ws.page_margins.right = 0.4
ws.page_margins.top = ws.page_margins.bottom = 0.5

# ============================================================
# シート1：サマリー
# ============================================================
s = wb.create_sheet("サマリー", 0)
s.sheet_view.showGridLines = False
widths = {"A": 2.5, "B": 20, "C": 13, "D": 2.5, "E": 20, "F": 13,
          "G": 2.5, "H": 20, "I": 13, "J": 2.5}
for k, v in widths.items():
    s.column_dimensions[k].width = v

s.merge_cells("B2:I2")
s["B2"] = "HND MMチーム　人員体制の現状と採用課題"
s["B2"].font = F(17, True, NAVY)
s["B2"].alignment = Alignment(horizontal="left", vertical="center")
s.row_dimensions[2].height = 34

s.merge_cells("B3:I3")
s["B3"] = "対象期間：2025年9月〜2026年8月（12ヶ月実績）　／　採用媒体の活用に関する稟議　添付資料　1/3"
s["B3"].font = F(10, False, GRAY)
s["B3"].alignment = Alignment(horizontal="left", vertical="center")
s.row_dimensions[3].height = 18

for col in range(2, 10):
    s.cell(row=4, column=col).border = Border(bottom=Side(style="medium", color=NAVY))
s.row_dimensions[4].height = 6

# --- KPIカード ---
def kpi(anchor_col, row, title, formula, fmt, note, accent=NAVY):
    c1 = get_column_letter(anchor_col)
    c2 = get_column_letter(anchor_col + 1)
    s.merge_cells(f"{c1}{row}:{c2}{row}")
    t = s[f"{c1}{row}"]
    t.value = title
    t.font = F(9.5, True, "FFFFFF")
    t.fill = fill(accent)
    t.alignment = CENTER
    s.row_dimensions[row].height = 20

    s.merge_cells(f"{c1}{row+1}:{c2}{row+1}")
    v = s[f"{c1}{row+1}"]
    v.value = formula
    v.number_format = fmt
    v.font = Font(name=JP, size=22, bold=True, color=accent)
    v.alignment = CENTER
    v.fill = fill(LABEL)
    s.row_dimensions[row + 1].height = 38

    s.merge_cells(f"{c1}{row+2}:{c2}{row+2}")
    n = s[f"{c1}{row+2}"]
    n.value = note
    n.font = F(8.5, False, GRAY)
    n.alignment = CENTER
    n.fill = fill(LABEL)
    s.row_dimensions[row + 2].height = 16

    for r in range(row, row + 3):
        for col in (anchor_col, anchor_col + 1):
            cc = s.cell(row=r, column=col)
            cc.border = Border(
                left=(med if col == anchor_col else None),
                right=(med if col == anchor_col + 1 else None),
                top=(med if r == row else None),
                bottom=(med if r == row + 2 else None))

D = "'月次推移'!"
kpi(2, 6, "年間退職者数", f"={D}N16+{D}N17", NUM, "正社員17名・派遣2名", RED)
kpi(5, 6, "年間離職率",   f"=({D}N16+{D}N17)/AVERAGE({D}B8:M8)", PCT, "年間退職者数 ÷ 平均月初稼働数", RED)
kpi(8, 6, "月平均退職者数", f"=({D}N16+{D}N17)/12", MEI, "12ヶ月中11ヶ月で退職が発生", RED)

kpi(2, 10, "年間採用者数", f"={D}N11+{D}N12", NUM, "正社員21名・派遣2名", NAVY)
kpi(5, 10, "年間の純増減", f"={D}N14-{D}N20", PLUS, "採用23名 − 離脱23名", NAVY)
kpi(8, 10, "期末の過不足", f"={D}M23", PLUS, "2026年8月末　必要数25名に対して", RED)

# --- 現状の課題 ---
s.merge_cells("B14:I14")
s["B14"] = "現状の課題"
s["B14"].font = F(12, True, NAVY)
s["B14"].fill = fill(BAND)
s["B14"].alignment = LEFTI
s.row_dimensions[14].height = 24

issues = [
    ("退職が恒常的に発生し、年間離職率は76.8%に達する",
     "12ヶ月で19名（正社員17名・派遣2名）が退職。月平均1.6名のペースで欠員が発生し続けている。"),
    ("年間23名を採用しても、同数の23名が離脱し純増はゼロ",
     "入（採用・異動）23名に対し、出（退職19名＋長期欠勤等による離脱4名）23名。採用は欠員補充に費やされ、体制の積み増しに至っていない。"),
    ("必要数の増加（20名→25名）に体制が追いつかず、期末時点で▲5名の欠員",
     "必要数が20名から25名へ引き上げられた一方、稼働数は期首・期末とも20名で横ばい。2026年7月以降は必要数を下回り、8月末は稼働率80%（20名／25名）の状態にある。"),
    ("現行の採用手法では、退職ペースを上回る採用が実現できていない",
     "欠員解消（5名）に加え、今後も想定される年間19名規模の退職補充が必要。採用母集団の拡大が不可欠である。"),
]
r = 15
for i, (head, body) in enumerate(issues, 1):
    s.merge_cells(f"B{r}:I{r}")
    c = s[f"B{r}"]
    c.value = f"{i}.　{head}"
    c.font = F(11, True, "1F1F1F")
    c.alignment = LEFTI
    s.row_dimensions[r].height = 22
    s.merge_cells(f"B{r+1}:I{r+1}")
    d = s[f"B{r+1}"]
    d.value = "　　" + body
    d.font = F(9.5, False, "404040")
    d.alignment = Alignment(horizontal="left", vertical="top", wrap_text=True, indent=1)
    s.row_dimensions[r + 1].height = 34
    r += 2

# --- 参考試算 ---
calc_top = r + 1
s.merge_cells(f"B{calc_top}:I{calc_top}")
s[f"B{calc_top}"] = "参考試算：今後1年間に必要となる採用数"
s[f"B{calc_top}"].font = F(12, True, NAVY)
s[f"B{calc_top}"].fill = fill(BAND)
s[f"B{calc_top}"].alignment = LEFTI
s.row_dimensions[calc_top].height = 24

rows = [
    ("① 現時点の欠員（2026年8月末）", f"=-{D}M23", NUM, "必要数25名 − 稼働数20名"),
    ("② 今後1年間の想定退職者数",      f"={D}N16+{D}N17", NUM, "直近12ヶ月の実績と同水準で推移すると仮定"),
    ("③ 必要採用数（① ＋ ②）",        None, NUM, "欠員解消および退職補充に要する年間採用数"),
]
rr = calc_top + 1
for i, (lab, f_, fmt, note) in enumerate(rows):
    s.merge_cells(f"B{rr}:C{rr}")
    lc = s[f"B{rr}"]
    lc.value = lab
    lc.font = F(10, i == 2)
    lc.alignment = LEFTI
    lc.fill = fill(TOTAL if i == 2 else LABEL)

    vc = s.cell(row=rr, column=5)
    vc.value = f_ if f_ else f"=E{calc_top+1}+E{calc_top+2}"
    vc.number_format = fmt
    vc.font = Font(name=JP, size=12 if i == 2 else 11, bold=True,
                   color=(NAVY if i == 2 else "000000"))
    vc.alignment = CENTER
    vc.fill = fill(TOTAL if i == 2 else LABEL)

    s.merge_cells(f"F{rr}:I{rr}")
    nc = s[f"F{rr}"]
    nc.value = note
    nc.font = F(9, False, GRAY)
    nc.alignment = LEFTI
    nc.fill = fill(TOTAL if i == 2 else LABEL)

    for col in range(2, 10):
        cc = s.cell(row=rr, column=col)
        cc.fill = fill(TOTAL if i == 2 else LABEL)
        cc.border = Border(
            left=(med if col == 2 else thin), right=(med if col == 9 else thin),
            top=thin, bottom=(med if i == 2 else thin))
    s.row_dimensions[rr].height = 21
    rr += 1

s.cell(row=calc_top + 2, column=5).comment = Comment(
    "前提：今後1年間の退職者数は直近12ヶ月の実績（19名）と同水準で推移すると仮定した試算値。", "作成者")

s.merge_cells(f"B{rr}:I{rr}")
s[f"B{rr}"] = "※ 上記②は直近12ヶ月の退職実績（19名）を前提とした試算値であり、実績値ではない。"
s[f"B{rr}"].font = F(9, False, GRAY)
s[f"B{rr}"].alignment = LEFTV
s.row_dimensions[rr].height = 16

foot = rr + 2
s.merge_cells(f"B{foot}:I{foot}")
s[f"B{foot}"] = "出典：HND MMチーム 人員推移管理表（2025年9月〜2026年8月実績）　／　明細は「月次推移」シート参照"
s[f"B{foot}"].font = F(9, False, GRAY)
s[f"B{foot}"].alignment = LEFTV

s.page_setup.orientation = "portrait"
s.page_setup.paperSize = s.PAPERSIZE_A4
s.page_setup.fitToWidth = 1
s.page_setup.fitToHeight = 1
s.sheet_properties.pageSetUpPr.fitToPage = True
s.page_margins.left = s.page_margins.right = 0.4

# ============================================================
# シート3：グラフ
# ============================================================
g = wb.create_sheet("グラフ")
g.sheet_view.showGridLines = False
g.column_dimensions["A"].width = 2.5
g.merge_cells("B2:L2")
g["B2"] = "HND MMチーム　人員推移グラフ"
g["B2"].font = F(15, True, NAVY)
g["B2"].alignment = Alignment(horizontal="left", vertical="center")
g.row_dimensions[2].height = 30
g.merge_cells("B3:L3")
g["B3"] = "対象期間：2025年9月〜2026年8月　／　採用媒体の活用に関する稟議　添付資料　3/3"
g["B3"].font = F(9, False, GRAY)
for col in range(2, 13):
    g.cell(row=4, column=col).border = Border(bottom=Side(style="medium", color=NAVY))
g.row_dimensions[4].height = 6

cats = Reference(ws, min_col=2, max_col=13, min_row=5, max_row=5)

ch1 = LineChart()
ch1.title = "必要数と稼働数の推移（名）"
ch1.style = 2
ch1.height = 8.6
ch1.width = 21
ch1.y_axis.title = "人数（名）"
ch1.y_axis.scaling.min = 0
for row, color, dash in ((7, "C00000", "dash"), (22, "1F3864", None)):
    ref = Reference(ws, min_col=1, max_col=13, min_row=row, max_row=row)
    ch1.add_data(ref, titles_from_data=True, from_rows=True)
ch1.set_categories(cats)
ch1.series[0].graphicalProperties.line.solidFill = "C00000"
ch1.series[0].graphicalProperties.line.width = 22000
ch1.series[0].graphicalProperties.line.dashStyle = "dash"
ch1.series[1].graphicalProperties.line.solidFill = "1F3864"
ch1.series[1].graphicalProperties.line.width = 28000
ch1.series[1].smooth = False
g.add_chart(ch1, "B6")

ch2 = BarChart()
ch2.type = "col"
ch2.title = "入（増員）と出（減員）の推移（名）"
ch2.style = 2
ch2.height = 8.6
ch2.width = 21
ch2.y_axis.title = "人数（名）"
ch2.gapWidth = 60
for row in (14, 20):
    ref = Reference(ws, min_col=1, max_col=13, min_row=row, max_row=row)
    ch2.add_data(ref, titles_from_data=True, from_rows=True)
ch2.set_categories(cats)
ch2.series[0].graphicalProperties.solidFill = "4472C4"
ch2.series[1].graphicalProperties.solidFill = "C00000"
g.add_chart(ch2, "B24")

g.page_setup.orientation = "portrait"
g.page_setup.paperSize = g.PAPERSIZE_A4
g.page_setup.fitToWidth = 1
g.sheet_properties.pageSetUpPr.fitToPage = True

wb.calculation.fullCalcOnLoad = True
s.sheet_properties.tabColor = NAVY
ws.sheet_properties.tabColor = NAVY_L
g.sheet_properties.tabColor = NAVY_L
s.sheet_view.tabSelected = True
ws.sheet_view.tabSelected = False
g.sheet_view.tabSelected = False
wb.active = 0
wb.properties.title = "HND MMチーム 人員推移（稟議添付資料）"
wb.properties.creator = "HND MMチーム"
wb.save(OUT)
print("saved:", OUT)
