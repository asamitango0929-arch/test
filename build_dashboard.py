# -*- coding: utf-8 -*-
"""短時間勤務者 採用KPIダッシュボード ビルダー"""
import openpyxl
from openpyxl.styles import Font, PatternFill, Alignment, Border, Side, NamedStyle
from openpyxl.utils import get_column_letter
from openpyxl.worksheet.datavalidation import DataValidation
from openpyxl.formatting.rule import CellIsRule, FormulaRule
from openpyxl.chart import BarChart, LineChart, Reference, Series
from openpyxl.chart.label import DataLabelList

OUT = "/home/user/test/短時間勤務者_採用KPIダッシュボード.xlsx"
JP = "Meiryo"

# ---------- palette ----------
NAVY   = "1F3864"   # title band
BLUE   = "2E5496"   # section header
BLUE_L = "D9E1F2"   # light band
GREY_L = "F2F2F2"
GREY_H = "808080"
WHITE  = "FFFFFF"
INPUT  = "FFF2CC"   # editable (target) cells
INK    = "1A1A1A"
GREEN  = "548235"
AMBER  = "BF8F00"
RED    = "C00000"
GREEN_F= "C6EFCE"; GREEN_T="006100"
AMBER_F= "FFEB9C"; AMBER_T="9C6500"
RED_F  = "FFC7CE"; RED_T ="9C0006"

thin = Side(style="thin", color="BFBFBF")
med  = Side(style="medium", color="808080")
BORDER = Border(left=thin,right=thin,top=thin,bottom=thin)

def fill(hexc): return PatternFill("solid", fgColor=hexc)
def font(sz=11,b=False,color=INK,it=False): return Font(name=JP,size=sz,bold=b,color=color,italic=it)
CEN = Alignment(horizontal="center", vertical="center", wrap_text=True)
LEF = Alignment(horizontal="left",   vertical="center", wrap_text=True)
RIG = Alignment(horizontal="right",  vertical="center")

wb = openpyxl.Workbook()

# =====================================================================
#  MASTER
# =====================================================================
mst = wb.active; mst.title = "マスタ"
master = {
 "媒体":       ["バイトル","マイナビバイト","ちゃんと","社員紹介","HP","その他"],
 "基地":       ["NRT","CTS"],
 "性別":       ["男","女"],
 "国籍":       ["日本","外国籍"],
 "書類選考":    ["未対応","確認中","通過","保留","不通過","辞退","連絡不通"],
 "書類NG理由":  ["年齢","通勤距離","国籍","日本語力","英語力","勤務時間不一致","その他"],
 "面接":       ["未設定","面接調整中","面接日確定","面接実施済み"],
 "面接NG理由":  ["勤務条件不一致","シフト条件不一致","経験不足","接客適性","コミュニケーション",
               "日本語力","英語力","航空保安上の要件","通勤困難","給与条件不一致",
               "入社可能時期が合わない","他社内定","本人辞退","総合判断","その他"],
 "採用":       ["選考中","採用","不採用","採用辞退","入社予定","入社辞退","保留"],
}
mst["A1"]="マスタ（プルダウン用リスト）"; mst["A1"].font=font(13,True,BLUE)
for c,(head,vals) in enumerate(master.items(),1):
    col=get_column_letter(c)
    h=mst.cell(row=2,column=c,value=head)
    h.font=font(11,True,WHITE); h.fill=fill(BLUE); h.alignment=CEN; h.border=BORDER
    for r,v in enumerate(vals,3):
        cel=mst.cell(row=r,column=c,value=v)
        cel.font=font(10); cel.border=BORDER; cel.alignment=LEF
    mst.column_dimensions[col].width = 16 if head in("面接NG理由","書類NG理由") else 12
mst.freeze_panes="A3"
# named ranges for validation
last = {h:(get_column_letter(i+1),2+len(v)) for i,(h,v) in enumerate(master.items())}

# =====================================================================
#  APPLICANT DATA  (応募者データ)
# =====================================================================
app = wb.create_sheet("応募者データ")
cols = ["No.","媒体","基地","応募日","氏名","年齢","性別","国籍",
        "書類選考","書類NG理由","面接","面接NG理由","採用","備考"]
widths=[5,13,7,11,16,6,6,8,10,12,12,14,10,26]
data = [
 [1,"バイトル","NRT",46227,"日高 秀光",73,"男","日本","不通過","年齢","未設定","","不採用",""],
 [2,"バイトル","NRT",46230,"浅利 真弓",63,"女","日本","不通過","年齢","未設定","","不採用",""],
 [3,"バイトル","NRT",46231,"浅木 洋二",76,"男","日本","不通過","年齢","未設定","","不採用",""],
 [4,"バイトル","NRT",46231,"佐伯 幸男",55,"男","日本","不通過","年齢","未設定","","不採用",""],
 [5,"バイトル","NRT",46232,"町田 侑里夏",21,"女","日本","通過","","面接調整中","","選考中",""],
 [6,"バイトル","NRT",46233,"三木 祐樹",38,"男","日本","不通過","年齢","未設定","","不採用",""],
 [7,"バイトル","NRT",46233,"金澤 拓也",19,"男","日本","通過","","面接調整中","","選考中",""],
 [8,"バイトル","NRT",46233,"NGUYEN NGOC DINH",25,"男","外国籍","不通過","日本語力","未設定","","不採用","N3"],
 [9,"バイトル","NRT",46236,"Hashini Navodya",21,"女","外国籍","確認中","","未設定","","選考中",""],
 [10,"バイトル","NRT",46236,"JARGALBAYAR ENKHBAYAN",21,"男","外国籍","確認中","","未設定","","選考中",""],
 [11,"バイトル","NRT",46237,"坂本 隼斗",24,"男","日本","確認中","","未設定","","選考中",""],
 [12,"バイトル","CTS",46236,"Hashini Navodya",21,"女","外国籍","確認中","","未設定","","選考中",""],
 [13,"マイナビバイト","CTS",46233,"竹内 愛",33,"女","日本","確認中","","未設定","","選考中",""],
 [14,"マイナビバイト","CTS",46233,"高砂 琴音",18,"女","日本","確認中","","未設定","","選考中","高校生だが外航のGATE経験者の可能性あり"],
]
app.merge_cells("A1:N1")
t=app["A1"]; t.value="応募者データ（明細）　― 毎日ここを更新してください。ダッシュボードは自動集計されます ―"
t.font=font(12,True,WHITE); t.fill=fill(NAVY); t.alignment=LEF
app.row_dimensions[1].height=24
for c,(h,w) in enumerate(zip(cols,widths),1):
    cell=app.cell(row=2,column=c,value=h)
    cell.font=font(10,True,WHITE); cell.fill=fill(BLUE); cell.alignment=CEN; cell.border=BORDER
    app.column_dimensions[get_column_letter(c)].width=w
DATA_START=3
for i,row in enumerate(data):
    r=DATA_START+i
    for c,v in enumerate(row,1):
        cell=app.cell(row=r,column=c,value=(v if v!="" else None))
        cell.font=font(10); cell.border=BORDER
        cell.alignment=CEN if c in (1,2,3,4,6,7,8,9,11,13) else LEF
        if c==4 and v!="":
            cell.number_format="yyyy/mm/dd"
app.freeze_panes="A3"
DATA_END=502   # allow future rows
# validations
dv_map={"媒体":"B","基地":"C","性別":"G","国籍":"H","書類選考":"I",
        "書類NG理由":"J","面接":"K","面接NG理由":"L","採用":"M"}
for key,col in dv_map.items():
    l,end=last[key]
    dv=DataValidation(type="list",formula1=f"=マスタ!${l}$3:${l}${end}",allow_blank=True)
    app.add_data_validation(dv)
    dv.add(f"{col}{DATA_START}:{col}{DATA_END}")

# =====================================================================
#  DAILY REPORT LOG (日次レポート)
# =====================================================================
dl = wb.create_sheet("日次レポート")
dcols=["日付","応募数","書類通過","面接設定","面接実施","内定","承諾","担当者","備考"]
dwidths=[13,9,9,9,9,7,7,12,30]
dl.merge_cells("A1:I1")
t=dl["A1"]; t.value="日次レポート（毎日 EOD に担当者が1行追記）"
t.font=font(12,True,WHITE); t.fill=fill(NAVY); t.alignment=LEF; dl.row_dimensions[1].height=24
dl.merge_cells("A2:I2")
n=dl["A2"]
n.value=("使い方：応募者データを更新後、当日の累計値を下表に1行追記します（右の「現在の累計値」を転記でOK）。"
         "ダッシュボードの前日比・推移グラフに自動反映されます。灰色の行はサンプルです。実データ入力時に削除してください。")
n.font=font(9,False,GREY_H,it=True); n.alignment=LEF; dl.row_dimensions[2].height=30
for c,(h,w) in enumerate(zip(dcols,dwidths),1):
    cell=dl.cell(row=3,column=c,value=h)
    cell.font=font(10,True,WHITE); cell.fill=fill(BLUE); cell.alignment=CEN; cell.border=BORDER
    dl.column_dimensions[get_column_letter(c)].width=w
# sample rows (grey)
samples=[
 [46234,12,1,1,0,0,0,"（サンプル）","← サンプル：削除して実データを入力"],
 [46235,13,2,2,0,0,0,"（サンプル）",""],
 [46236,14,2,2,0,0,0,"（サンプル）",""],
]
DL_START=4
for i,row in enumerate(samples):
    r=DL_START+i
    for c,v in enumerate(row,1):
        cell=dl.cell(row=r,column=c,value=v)
        cell.font=font(10,color=GREY_H,it=True); cell.border=BORDER
        cell.alignment=CEN if c<=7 else LEF
        cell.fill=fill(GREY_L)
        if c==1: cell.number_format="yyyy/mm/dd"
# empty input rows styled
for r in range(DL_START+len(samples), DL_START+len(samples)+20):
    for c in range(1,10):
        cell=dl.cell(row=r,column=c)
        cell.border=BORDER
        if c==1: cell.number_format="yyyy/mm/dd"
DL_LAST=200
dl.freeze_panes="A4"
# "現在の累計値" helper block (live from 応募者データ) to the right (K:L)
dl["K3"]="現在の累計値（自動）"; dl["K3"].font=font(10,True,WHITE); dl["K3"].fill=fill(GREEN)
dl["K3"].alignment=CEN; dl.merge_cells("K3:L3")
dl.column_dimensions["K"].width=14; dl.column_dimensions["L"].width=9
A="応募者データ"
live_formulas=[
 ("応募数",  f'=COUNTA({A}!$E$3:$E${DATA_END})'),
 ("書類通過", f'=COUNTIF({A}!$I$3:$I${DATA_END},"通過")'),
 ("面接設定", f'=COUNTIF({A}!$K$3:$K${DATA_END},"面接調整中")+COUNTIF({A}!$K$3:$K${DATA_END},"面接日確定")+COUNTIF({A}!$K$3:$K${DATA_END},"面接実施済み")'),
 ("面接実施", f'=COUNTIF({A}!$K$3:$K${DATA_END},"面接実施済み")'),
 ("内定",    f'=COUNTIF({A}!$M$3:$M${DATA_END},"採用")+COUNTIF({A}!$M$3:$M${DATA_END},"入社予定")+COUNTIF({A}!$M$3:$M${DATA_END},"採用辞退")+COUNTIF({A}!$M$3:$M${DATA_END},"入社辞退")'),
 ("承諾",    f'=COUNTIF({A}!$M$3:$M${DATA_END},"採用")+COUNTIF({A}!$M$3:$M${DATA_END},"入社予定")'),
]
for i,(lab,fm) in enumerate(live_formulas):
    r=4+i
    dl.cell(row=r,column=11,value=lab).font=font(10,True); dl.cell(row=r,column=11).border=BORDER
    dl.cell(row=r,column=11).alignment=LEF; dl.cell(row=r,column=11).fill=fill(BLUE_L)
    cc=dl.cell(row=r,column=12,value=fm); cc.font=font(10); cc.border=BORDER; cc.alignment=CEN

# =====================================================================
#  DASHBOARD (KPIダッシュボード)
# =====================================================================
dash = wb.create_sheet("KPIダッシュボード")
wb.move_sheet("KPIダッシュボード", -(wb.sheetnames.index("KPIダッシュボード")))  # move to front
dash.sheet_view.showGridLines=False
dash.column_dimensions["A"].width=2
for col,w in zip("BCDEFGH",[18,11,12,10,11,11,11]):
    dash.column_dimensions[col].width=w
dash.column_dimensions["I"].width=2

# Title band
dash.merge_cells("B1:H1")
t=dash["B1"]; t.value="📊 短時間勤務者　採用KPIダッシュボード"
t.font=font(18,True,WHITE); t.fill=fill(NAVY); t.alignment=Alignment(horizontal="left",vertical="center")
dash.row_dimensions[1].height=34
dash.merge_cells("B2:H2")
st=dash["B2"]
st.value='=TEXT(MAX(日次レポート!$A$4:$A$'+str(DL_LAST)+'),"yyyy/mm/dd")&" 時点　｜　成田(NRT)・千歳(CTS) パート採用進捗　｜　数値は応募者データより自動集計"'
st.font=font(10,color=WHITE); st.fill=fill(BLUE); st.alignment=Alignment(horizontal="left",vertical="center")
dash.row_dimensions[2].height=20

D="KPIダッシュボード"
def sec(row, text):
    dash.merge_cells(f"B{row}:H{row}")
    c=dash[f"B{row}"]; c.value=text; c.font=font(12,True,WHITE); c.fill=fill(BLUE)
    c.alignment=LEF; dash.row_dimensions[row].height=22

def hdr(row, labels, start=2):
    for i,lab in enumerate(labels):
        c=dash.cell(row=row,column=start+i,value=lab)
        c.font=font(10,True,WHITE); c.fill=fill(GREY_H); c.alignment=CEN; c.border=BORDER

# ---- Section 1: funnel target vs actual ----
sec(4, "①  採用ファネル：目標 vs 実績（累計）")
hdr(5, ["指標","目標","実績(累計)","達成率","前日","前日比"])
funnel_rows=["応募数","書類通過","面接設定","面接実施","内定","承諾"]
targets=[50,25,15,12,8,6]
# live formulas per metric (col D = 実績)
live=[
 f'=COUNTA({A}!$E$3:$E${DATA_END})',
 f'=COUNTIF({A}!$I$3:$I${DATA_END},"通過")',
 f'=COUNTIF({A}!$K$3:$K${DATA_END},"面接調整中")+COUNTIF({A}!$K$3:$K${DATA_END},"面接日確定")+COUNTIF({A}!$K$3:$K${DATA_END},"面接実施済み")',
 f'=COUNTIF({A}!$K$3:$K${DATA_END},"面接実施済み")',
 f'=COUNTIF({A}!$M$3:$M${DATA_END},"採用")+COUNTIF({A}!$M$3:$M${DATA_END},"入社予定")+COUNTIF({A}!$M$3:$M${DATA_END},"採用辞退")+COUNTIF({A}!$M$3:$M${DATA_END},"入社辞退")',
 f'=COUNTIF({A}!$M$3:$M${DATA_END},"採用")+COUNTIF({A}!$M$3:$M${DATA_END},"入社予定")',
]
dlcol=["B","C","D","E","F","G"]  # daily log cols for 前日 lookup: 応募数=B ...承諾=G
F1=6
for i,name in enumerate(funnel_rows):
    r=F1+i
    dash.cell(row=r,column=2,value=name).font=font(11,True)
    dash.cell(row=r,column=2).alignment=LEF
    # target (input, yellow)
    tc=dash.cell(row=r,column=3,value=targets[i])
    tc.font=font(11,color="0000FF"); tc.fill=fill(INPUT); tc.alignment=CEN
    # actual live
    ac=dash.cell(row=r,column=4,value=live[i]); ac.font=font(11,True); ac.alignment=CEN
    # achievement
    dash.cell(row=r,column=5,value=f'=IFERROR(D{r}/C{r},"")').number_format="0.0%"
    dash.cell(row=r,column=5).alignment=CEN
    # 前日 (last logged value)
    dcol=dlcol[i]
    dash.cell(row=r,column=6,value=f'=IFERROR(LOOKUP(1E+100,日次レポート!{dcol}$4:{dcol}${DL_LAST}),0)').alignment=CEN
    # 前日比
    dc=dash.cell(row=r,column=7,value=f'=D{r}-F{r}'); dc.alignment=CEN
    dc.number_format='+#,##0;-#,##0;0'
    for c in range(2,8):
        cell=dash.cell(row=r,column=c); cell.border=BORDER
        if c in(4,6): cell.font=font(11)
        if r%2==0 and cell.fill.fgColor.rgb in (None,"00000000"):
            pass
# conditional format: 達成率
dash.conditional_formatting.add(f"E{F1}:E{F1+5}",
    CellIsRule(operator="greaterThanOrEqual",formula=["1"],fill=fill(GREEN_F),font=Font(name=JP,color=GREEN_T,bold=True)))
dash.conditional_formatting.add(f"E{F1}:E{F1+5}",
    CellIsRule(operator="between",formula=["0.7","0.9999"],fill=fill(AMBER_F),font=Font(name=JP,color=AMBER_T)))
dash.conditional_formatting.add(f"E{F1}:E{F1+5}",
    CellIsRule(operator="lessThan",formula=["0.7"],fill=fill(RED_F),font=Font(name=JP,color=RED_T)))
# conditional format: 前日比
dash.conditional_formatting.add(f"G{F1}:G{F1+5}",
    CellIsRule(operator="greaterThan",formula=["0"],fill=fill(GREEN_F),font=Font(name=JP,color=GREEN_T,bold=True)))
dash.conditional_formatting.add(f"G{F1}:G{F1+5}",
    CellIsRule(operator="lessThan",formula=["0"],fill=fill(RED_F),font=Font(name=JP,color=RED_T,bold=True)))

# ---- Section 2: conversion rates ----
CV=13
sec(CV, "②  転換率（歩留まり）")
hdr(CV+1, ["指標","値","計算式・定義"])
conv=[
 ("書類通過率", f"=IFERROR(D{F1+1}/D{F1},\"\")", "書類通過 ÷ 応募数"),
 ("面接設定率", f"=IFERROR(D{F1+2}/D{F1+1},\"\")", "面接設定 ÷ 書類通過"),
 ("面接実施率", f"=IFERROR(D{F1+3}/D{F1+2},\"\")", "面接実施 ÷ 面接設定"),
 ("内定率",    f"=IFERROR(D{F1+4}/D{F1+3},\"\")", "内定 ÷ 面接実施"),
 ("承諾率",    f"=IFERROR(D{F1+5}/D{F1+4},\"\")", "承諾 ÷ 内定"),
]
for i,(name,fm,note) in enumerate(conv):
    r=CV+2+i
    dash.cell(row=r,column=2,value=name).font=font(11,True); dash.cell(row=r,column=2).alignment=LEF
    vc=dash.cell(row=r,column=3,value=fm); vc.number_format="0.0%"; vc.alignment=CEN; vc.font=font(11,True,BLUE)
    dash.merge_cells(start_row=r,start_column=4,end_row=r,end_column=7)
    nc=dash.cell(row=r,column=4,value=note); nc.font=font(9,color=GREY_H); nc.alignment=LEF
    for c in range(2,8): dash.cell(row=r,column=c).border=BORDER

# ---- Section 3: by media ----
MB=20
sec(MB, "③  媒体別 実績")
hdr(MB+1, ["媒体","応募数","書類通過","面接設定","面接実施","内定","承諾"])
media=master["媒体"]
def media_formulas(mref):
    return [
      f'=COUNTIF({A}!$B$3:$B${DATA_END},{mref})',
      f'=COUNTIFS({A}!$B$3:$B${DATA_END},{mref},{A}!$I$3:$I${DATA_END},"通過")',
      f'=COUNTIFS({A}!$B$3:$B${DATA_END},{mref},{A}!$K$3:$K${DATA_END},"面接調整中")+COUNTIFS({A}!$B$3:$B${DATA_END},{mref},{A}!$K$3:$K${DATA_END},"面接日確定")+COUNTIFS({A}!$B$3:$B${DATA_END},{mref},{A}!$K$3:$K${DATA_END},"面接実施済み")',
      f'=COUNTIFS({A}!$B$3:$B${DATA_END},{mref},{A}!$K$3:$K${DATA_END},"面接実施済み")',
      f'=COUNTIFS({A}!$B$3:$B${DATA_END},{mref},{A}!$M$3:$M${DATA_END},"採用")+COUNTIFS({A}!$B$3:$B${DATA_END},{mref},{A}!$M$3:$M${DATA_END},"入社予定")+COUNTIFS({A}!$B$3:$B${DATA_END},{mref},{A}!$M$3:$M${DATA_END},"採用辞退")+COUNTIFS({A}!$B$3:$B${DATA_END},{mref},{A}!$M$3:$M${DATA_END},"入社辞退")',
      f'=COUNTIFS({A}!$B$3:$B${DATA_END},{mref},{A}!$M$3:$M${DATA_END},"採用")+COUNTIFS({A}!$B$3:$B${DATA_END},{mref},{A}!$M$3:$M${DATA_END},"入社予定")',
    ]
for i,m in enumerate(media):
    r=MB+2+i
    nc=dash.cell(row=r,column=2,value=m); nc.font=font(10,True); nc.alignment=LEF
    for j,fm in enumerate(media_formulas(f'"{m}"')):
        cc=dash.cell(row=r,column=3+j,value=fm); cc.alignment=CEN; cc.font=font(10)
    for c in range(2,9): dash.cell(row=r,column=c).border=BORDER
    if i%2==1:
        for c in range(2,9): dash.cell(row=r,column=c).fill=fill(GREY_L)
# total row
tr=MB+2+len(media)
dash.cell(row=tr,column=2,value="合計").font=font(10,True,WHITE)
dash.cell(row=tr,column=2).fill=fill(BLUE); dash.cell(row=tr,column=2).alignment=LEF
for j in range(6):
    col=get_column_letter(3+j)
    cc=dash.cell(row=tr,column=3+j,value=f'=SUM({col}{MB+2}:{col}{MB+1+len(media)})')
    cc.font=font(10,True,WHITE); cc.fill=fill(BLUE); cc.alignment=CEN
for c in range(2,9): dash.cell(row=tr,column=c).border=BORDER

# ---- Section 4: by base ----
BB=tr+2
sec(BB, "④  基地別 実績（NRT=成田 / CTS=千歳）")
hdr(BB+1, ["基地","応募数","書類通過","面接設定","面接実施","内定","承諾"])
def base_formulas(bref):
    return [
      f'=COUNTIF({A}!$C$3:$C${DATA_END},{bref})',
      f'=COUNTIFS({A}!$C$3:$C${DATA_END},{bref},{A}!$I$3:$I${DATA_END},"通過")',
      f'=COUNTIFS({A}!$C$3:$C${DATA_END},{bref},{A}!$K$3:$K${DATA_END},"面接調整中")+COUNTIFS({A}!$C$3:$C${DATA_END},{bref},{A}!$K$3:$K${DATA_END},"面接日確定")+COUNTIFS({A}!$C$3:$C${DATA_END},{bref},{A}!$K$3:$K${DATA_END},"面接実施済み")',
      f'=COUNTIFS({A}!$C$3:$C${DATA_END},{bref},{A}!$K$3:$K${DATA_END},"面接実施済み")',
      f'=COUNTIFS({A}!$C$3:$C${DATA_END},{bref},{A}!$M$3:$M${DATA_END},"採用")+COUNTIFS({A}!$C$3:$C${DATA_END},{bref},{A}!$M$3:$M${DATA_END},"入社予定")+COUNTIFS({A}!$C$3:$C${DATA_END},{bref},{A}!$M$3:$M${DATA_END},"採用辞退")+COUNTIFS({A}!$C$3:$C${DATA_END},{bref},{A}!$M$3:$M${DATA_END},"入社辞退")',
      f'=COUNTIFS({A}!$C$3:$C${DATA_END},{bref},{A}!$M$3:$M${DATA_END},"採用")+COUNTIFS({A}!$C$3:$C${DATA_END},{bref},{A}!$M$3:$M${DATA_END},"入社予定")',
    ]
bases=[("NRT","成田"),("CTS","千歳")]
for i,(code,label) in enumerate(bases):
    r=BB+2+i
    dash.cell(row=r,column=2,value=f"{code}（{label}）").font=font(10,True)
    dash.cell(row=r,column=2).alignment=LEF
    for j,fm in enumerate(base_formulas(f'"{code}"')):
        cc=dash.cell(row=r,column=3+j,value=fm); cc.alignment=CEN; cc.font=font(10)
    for c in range(2,9): dash.cell(row=r,column=c).border=BORDER

# ---- legend / notes ----
LG=BB+5
dash.merge_cells(f"B{LG}:H{LG}")
dash.cell(row=LG,column=2,value="◆ 凡例・定義").font=font(10,True,BLUE)
notes=[
 "・黄色セル＝入力（目標値）。運用開始時に目標を設定してください（現在は仮値）。",
 "・実績は［応募者データ］シートから自動集計。担当者は毎日そこを更新するだけです。",
 "・前日／前日比は［日次レポート］の最終記録行との差分。毎日EODに1行追記してください。",
 "・面接設定＝面接ステータスが「面接調整中／面接日確定／面接実施済み」。面接実施＝「面接実施済み」。",
 "・内定＝採用が「採用／入社予定／採用辞退／入社辞退」。承諾＝「採用／入社予定」。",
]
for i,tx in enumerate(notes):
    r=LG+1+i; dash.merge_cells(f"B{r}:H{r}")
    dash.cell(row=r,column=2,value=tx).font=font(9,color=GREY_H)
    dash.cell(row=r,column=2).alignment=LEF

# =====================================================================
#  CHARTS
# =====================================================================
# 1) Funnel bar (実績)
ch1=BarChart(); ch1.type="bar"; ch1.title="採用ファネル（実績）"; ch1.height=7; ch1.width=12
d=Reference(dash,min_col=4,min_row=F1,max_row=F1+5)
cats=Reference(dash,min_col=2,min_row=F1,max_row=F1+5)
ch1.add_data(d,titles_from_data=False); ch1.set_categories(cats)
ch1.legend=None; ch1.dataLabels=DataLabelList(); ch1.dataLabels.showVal=True
dash.add_chart(ch1,"J4")

# 2) media clustered bar (応募数/面接実施/承諾)
ch2=BarChart(); ch2.type="col"; ch2.title="媒体別 応募数"; ch2.height=7; ch2.width=12
md=Reference(dash,min_col=3,min_row=MB+1,max_row=MB+1+len(media))  # 応募数 col with header
mcats=Reference(dash,min_col=2,min_row=MB+2,max_row=MB+1+len(media))
ch2.add_data(md,titles_from_data=True); ch2.set_categories(mcats)
ch2.legend=None; ch2.dataLabels=DataLabelList(); ch2.dataLabels.showVal=True
dash.add_chart(ch2,"J20")

# 3) trend line (from daily log)
ch3=LineChart(); ch3.title="推移（日次）"; ch3.height=7; ch3.width=12
tdata=Reference(dl,min_col=2,max_col=7,min_row=3,max_row=DL_START+len(samples)+9)
tcats=Reference(dl,min_col=1,min_row=DL_START,max_row=DL_START+len(samples)+9)
ch3.add_data(tdata,titles_from_data=True); ch3.set_categories(tcats)
ch3.markers=True
dash.add_chart(ch3,"J36")

dash.sheet_view.zoomScale=100
dash.sheet_properties.tabColor=NAVY

# force full recalculation when the file is opened (Excel/LibreOffice/Sheets)
from openpyxl.workbook.properties import CalcProperties
wb.calculation = CalcProperties(calcId=0, fullCalcOnLoad=True)

wb.save(OUT)
print("saved", OUT)
print("sheets:", wb.sheetnames)
