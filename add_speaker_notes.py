# -*- coding: utf-8 -*-
"""追加スライドに docx のセリフ案をスピーカーノートとして埋め込む。
発表者本人の語り（一人称）として原文どおり挿入。"""
import json
from pptx import Presentation

nar = json.load(open("/tmp/narration.json"))
prs = Presentation("成田_短時間勤務者採用_説明会.pptx")

def set_notes(idx, paragraphs):
    tf = prs.slides[idx].notes_slide.notes_text_frame
    tf.clear()
    for i, para in enumerate(paragraphs):
        p = tf.paragraphs[0] if i == 0 else tf.add_paragraph()
        p.text = para

# slide9 (idx8) 需要カーブ ← 追加①のセリフ
set_notes(8, nar["1"])
# slide10 (idx9) ガント ← 補足用の短いノート
set_notes(9, [
    "補足スライドです。便ごとの地上作業帯（カウンター開設から出発まで）を時間軸で並べています。",
    "青が朝便、橙が昼・夕便。灰色の線が12時です。",
    "個々の便で見ても、ほとんどの作業が12時までに終わり、午後はほぼ空くことを確認いただけます。",
])
# slide11 (idx10) 配置候補便 ← 事実ベースの短いノート
set_notes(10, [
    "現時点で人事として、朝のピーク帯に短時間勤務者の配置を想定している便の一覧です。",
    "対象はVN・AM・RF・5J・UL・VJ。時刻はすべて実際の線表から取っており、推測値はありません。",
    "VJの夕方便（VJ934/935）は、Winterで運休する場合は対象外、運航する場合は別途検討します。ULは火・木・土の運航です。",
])
# slide12 (idx11) 満席のレストラン ← 追加②のセリフ
set_notes(11, nar["2"])
# slide13 (idx12) 良いサービス≠良い経営 ← 追加③のセリフ
set_notes(12, nar["3"])
# slide14 (idx13) 売上を増やすだけ ← 追加④のセリフ ＋ 既存スライドへのつなぎ
bridge = nar["bridge"][0].split("\n\n") if nar.get("bridge") else []
set_notes(13, nar["4"] + ["", "―― 次のスライドへのつなぎ ――"] + bridge)

# p.9(idx10) の見出し「（社長案）」→「（人事部案）」に修正
sl = prs.slides[10]
for sh in sl.shapes:
    if sh.has_text_frame:
        for p in sh.text_frame.paragraphs:
            for r in p.runs:
                if "社長案" in r.text:
                    r.text = r.text.replace("社長案", "人事部案")

prs.save("out_deck.pptx")
print("saved")
