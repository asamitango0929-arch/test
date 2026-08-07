# -*- coding: utf-8 -*-
"""社長報告シート（社長報告08_05 = sheet1）の下に、コンサル標準の
Issue / Action Tracker を追記する。

方針：
- openpyxl で保存するとグラフ（KPIダッシュボード）や入力規則（応募者データの
  x14 dataValidation）が失われるため、XML を直接編集して既存要素を完全保持する。
- 追記対象は sheet1.xml（社長報告）と styles.xml（スタイル追加）のみ。
- スタイルは既存のもの（青セクション見出し=s55、グレー表ヘッダ=s2）を流用し、
  データ行用の xf を数個だけ追加する。
"""
import re
import shutil
import zipfile
import os

SRC = "KPI_original.xlsx"
DST = "採用活動_日次報告書_ActionTracker付.xlsx"
WORK = "_build"

# ---- Issue / Action Tracker のデータ（提供資料を忠実に反映） ----
# (課題/Action, TODO（次にやること）, Responsible, STATUS（進捗状況）, DUE DATE)
UPDATED = "2026-08-08"
rows = [
    ("社員紹介制度の開始",
     "稟議を上程・決裁を取得し、制度を正式スタート",
     "遠藤・髙田", "稟議準備中。現場説明・協力依頼済", "8/14"),
    ("交通費ルールの整理",
     "短時間勤務者向けの交通費取扱いを確定・規程化",
     "髙田・坂本", "短時間勤務者向け取扱いを検討中", "8/14"),
    ("既存パート時給の見直し",
     "UL等 既存者との時給差の是正方針を決定",
     "髙田・坂本", "UL等、既存者との時給差を確認中", "8/14"),
    ("短時間勤務者向け教育の整理",
     "現行7日間研修に代わる短縮研修プランを策定",
     "加納", "現行7日間研修が採用上のネック。代替方法を検討", "8/14"),
    ("効果的な配置時間帯・業務の特定",
     "各所属長から時間帯別の必要出面を回収・集約",
     "各所属長・内山", "本日、各所属長へ検討依頼", "8/12"),
    ("経験者採用方法の具体化",
     "採用チャネル案を取りまとめ、実施方針を決定",
     "内山", "現場へアイデア出しを依頼。紹介制度と並行検討", "8/12"),
    ("面接予定3名の選考",
     "面接を実施し合否を判定（経験者1名含む）",
     "内山・野澤", "面接調整中（経験者1名含む）", "随時"),
]
NEAR_DUE = {"8/12"}   # 直近期日はハイライト


def esc(s):
    return (s.replace("&", "&amp;").replace("<", "&lt;").replace(">", "&gt;"))


def unzip():
    if os.path.exists(WORK):
        shutil.rmtree(WORK)
    os.makedirs(WORK)
    with zipfile.ZipFile(SRC) as z:
        names = z.namelist()
        z.extractall(WORK)
    return names


def add_styles():
    """styles.xml の cellXfs に追加スタイルを append し、開始インデックスを返す。"""
    path = os.path.join(WORK, "xl/styles.xml")
    s = open(path, encoding="utf-8").read()
    m = re.search(r'<cellXfs count="(\d+)">', s)
    base = int(m.group(1))

    def xf(font, fill, border, h, v, wrap=True):
        a = f'<alignment horizontal="{h}" vertical="{v}"' + (' wrapText="1"' if wrap else '') + '/>'
        return (f'<xf numFmtId="0" fontId="{font}" fillId="{fill}" borderId="{border}" xfId="0"'
                f' applyFont="1" applyFill="1" applyBorder="1" applyAlignment="1">{a}</xf>')

    # font3=Meiryo10 dark, font8=bold10 dark, font13=grey9, font6=italic grey9
    # fill0=none, fill4=F2F2F2(zebra), fill7=FFF2CC(highlight)
    # border1=full thin
    new = [
        xf(8, 0, 1, "left", "top"),      # +0 Action (bold) 白行
        xf(3, 0, 1, "left", "top"),      # +1 通常テキスト 白行
        xf(3, 0, 1, "center", "center"), # +2 中央 白行
        xf(8, 4, 1, "left", "top"),      # +3 Action(bold) ゼブラ
        xf(3, 4, 1, "left", "top"),      # +4 通常テキスト ゼブラ
        xf(3, 4, 1, "center", "center"), # +5 中央 ゼブラ
        xf(8, 7, 1, "center", "center"), # +6 DUE ハイライト
        xf(13, 0, 0, "left", "center"),  # +7 サブタイトル(灰italicではなく灰)
        xf(6, 0, 0, "left", "center"),   # +8 凡例(italic灰)
    ]
    s = s.replace('<cellXfs count="%d">' % base,
                  '<cellXfs count="%d">' % (base + len(new)), 1)
    s = s.replace("</cellXfs>", "".join(new) + "</cellXfs>", 1)
    open(path, "w", encoding="utf-8").write(s)
    return base


def cell(coord, style, text=None):
    if text is None:
        return f'<c r="{coord}" s="{style}"/>'
    return (f'<c r="{coord}" s="{style}" t="inlineStr">'
            f'<is><t xml:space="preserve">{esc(text)}</t></is></c>')


def build_rows(base):
    S_ACT_W, S_TXT_W, S_CTR_W = base + 0, base + 1, base + 2
    S_ACT_Z, S_TXT_Z, S_CTR_Z = base + 3, base + 4, base + 5
    S_DUE_HL = base + 6
    S_SUB, S_LEG = base + 7, base + 8
    S_SECTION = 55   # 既存の青セクション見出し
    S_HEAD = 2       # 既存のグレー表ヘッダ

    out = []
    merges = []

    # row 60: spacer
    out.append('<row r="60" spans="2:7" ht="6" customHeight="1"/>')

    # row 61: セクション見出し（B:G マージ）
    out.append(f'<row r="61" spans="2:7" ht="21" customHeight="1">'
               f'{cell("B61", S_SECTION, "■ Issue / Action Tracker（課題・アクション管理）")}</row>')
    merges.append("B61:G61")

    # row 62: サブタイトル（B:G マージ）
    sub = (f"コンサル標準の課題管理表｜Action・TODO・Responsible・STATUS・DUE DATE"
           f"｜更新日：{UPDATED}｜出典：担当ヒアリング（Issue/Action Tracker）")
    out.append(f'<row r="62" spans="2:7" ht="16" customHeight="1">'
               f'{cell("B62", S_SUB, sub)}</row>')
    merges.append("B62:G62")

    # row 63: 見出し行  B=課題/Action C=TODO D=Responsible E=STATUS F:G=DUE DATE
    heads = [("B63", "課題 / Action"), ("C63", "TODO（次にやること）"),
             ("D63", "Responsible"), ("E63", "STATUS（進捗状況）"),
             ("F63", "DUE DATE"), ("G63", None)]
    hcells = "".join(cell(c, S_HEAD, t) for c, t in heads)
    out.append(f'<row r="63" spans="2:7" ht="24" customHeight="1">{hcells}</row>')
    merges.append("F63:G63")

    # data rows 64..
    r = 64
    for i, (act, todo, resp, status, due) in enumerate(rows):
        zebra = (i % 2 == 1)
        s_act = S_ACT_Z if zebra else S_ACT_W
        s_txt = S_TXT_Z if zebra else S_TXT_W
        s_ctr = S_CTR_Z if zebra else S_CTR_W
        s_due = S_DUE_HL if due in NEAR_DUE else s_ctr
        no = f"{i+1}　"
        cells = "".join([
            cell(f"B{r}", s_act, no + act),
            cell(f"C{r}", s_txt, todo),
            cell(f"D{r}", s_ctr, resp),
            cell(f"E{r}", s_txt, status),
            cell(f"F{r}", s_due, due),
            cell(f"G{r}", s_due, None),
        ])
        out.append(f'<row r="{r}" spans="2:7" ht="46" customHeight="1">{cells}</row>')
        merges.append(f"F{r}:G{r}")
        r += 1

    # 凡例
    leg = ("凡例）本トラッカーは担当ヒアリングの Issue/Action Tracker を反映。"
           "STATUS＝進捗状況、TODO＝各進捗から導いた次アクション（要編集）。"
           "DUE DATE の直近（8/12）は黄色でハイライト。上部KPI・属性は［応募者データ］から自動集計。")
    out.append(f'<row r="{r}" spans="2:7" ht="30" customHeight="1">'
               f'{cell("B%d" % r, S_LEG, leg)}</row>')
    merges.append(f"B{r}:G{r}")
    last_row = r

    return "".join(out), merges, last_row


def patch_sheet1(base):
    path = os.path.join(WORK, "xl/worksheets/sheet1.xml")
    s = open(path, encoding="utf-8").read()

    rows_xml, merges, last_row = build_rows(base)

    # 1) sheetData の直前に行を挿入
    s = s.replace("</sheetData>", rows_xml + "</sheetData>", 1)

    # 2) mergeCells に追加
    m = re.search(r'<mergeCells count="(\d+)">', s)
    old = int(m.group(1))
    add = "".join(f'<mergeCell ref="{ref}"/>' for ref in merges)
    s = s.replace('<mergeCells count="%d">' % old,
                  '<mergeCells count="%d">' % (old + len(merges)), 1)
    s = s.replace("</mergeCells>", add + "</mergeCells>", 1)

    # 3) 列幅を Tracker 用に調整（B..G を個別指定）
    newcols = ('<cols>'
               '<col min="1" max="1" width="2" customWidth="1"/>'
               '<col min="2" max="2" width="22" customWidth="1"/>'
               '<col min="3" max="3" width="24" customWidth="1"/>'
               '<col min="4" max="4" width="14" customWidth="1"/>'
               '<col min="5" max="5" width="22" customWidth="1"/>'
               '<col min="6" max="6" width="12" customWidth="1"/>'
               '<col min="7" max="7" width="13" customWidth="1"/>'
               '</cols>')
    s = re.sub(r"<cols>.*?</cols>", newcols, s, count=1, flags=re.S)

    # 4) dimension を更新
    s = re.sub(r'<dimension ref="[^"]*"/>',
               f'<dimension ref="B1:H{last_row}"/>', s, count=1)

    open(path, "w", encoding="utf-8").write(s)
    return last_row


def rezip(names):
    if os.path.exists(DST):
        os.remove(DST)
    with zipfile.ZipFile(DST, "w", zipfile.ZIP_DEFLATED) as z:
        for name in names:
            z.write(os.path.join(WORK, name), name)


def main():
    names = unzip()
    base = add_styles()
    last = patch_sheet1(base)
    rezip(names)
    print(f"OK -> {DST}  (tracker rows 61-{last}, styles base xf={base})")


if __name__ == "__main__":
    main()
