const fs = require("fs");
const {
  Document, Packer, Paragraph, TextRun, AlignmentType,
  Table, TableRow, TableCell, WidthType, BorderStyle, ShadingType, HeightRule, VerticalAlign,
} = require("docx");

const JP = "Yu Gothic";
const CW = 10106; // A4 (11906) minus 900 margins each side
const SH = ShadingType.CLEAR;

// ---- primitives ----
const run = (text, o = {}) =>
  new TextRun({ text, font: { name: JP, eastAsia: JP }, size: o.size || 17, bold: !!o.bold, color: o.color, underline: o.underline });

const p = (children, o = {}) =>
  new Paragraph({
    children: Array.isArray(children) ? children : [children],
    alignment: o.align,
    spacing: { before: o.before || 0, after: o.after == null ? 20 : o.after, line: o.line || 216 },
    border: o.border,
    shading: o.fill ? { type: SH, fill: o.fill, color: "auto" } : undefined,
  });

const sep = () => new Paragraph({ children: [run("", { size: 2 })], spacing: { before: 0, after: 0, line: 20 } });

// section header bar
const section = (text) =>
  new Paragraph({
    children: [run(text, { bold: true, color: "FFFFFF", size: 19 })],
    spacing: { before: 90, after: 20, line: 216 },
    shading: { type: SH, fill: "1F4E79", color: "auto" },
    border: { top: { style: BorderStyle.SINGLE, size: 2, color: "1F4E79" }, bottom: { style: BorderStyle.SINGLE, size: 2, color: "1F4E79" }, left: { style: BorderStyle.SINGLE, size: 8, color: "1F4E79", space: 4 }, right: { style: BorderStyle.SINGLE, size: 2, color: "1F4E79" } },
  });

// sub label (①②…)
const sub = (text) =>
  new Paragraph({ children: [run(text, { bold: true, color: "1F4E79", size: 17 })], spacing: { before: 50, after: 16, line: 200 } });

// ---- table builder ----
const bar = { style: BorderStyle.SINGLE, size: 4, color: "9BA7B4" };
const B = { top: bar, bottom: bar, left: bar, right: bar };

// cell: {t, span, label, align, bold, color}
function cellWidth(colWidths, start, span) {
  let w = 0;
  for (let i = start; i < start + (span || 1); i++) w += colWidths[i];
  return w;
}
function mkCell(c, w) {
  return new TableCell({
    width: { size: w, type: WidthType.DXA },
    columnSpan: c.span || 1,
    borders: B,
    verticalAlign: VerticalAlign.CENTER,
    margins: { top: 30, bottom: 30, left: 90, right: 70 },
    shading: c.label ? { type: SH, fill: "E9EFF6", color: "auto" } : undefined,
    children: [p(run(c.t || "", { bold: c.label || c.bold, size: c.size || 17, color: c.color }), { after: 0, align: c.align, line: 200 })],
  });
}
function table(colWidths, rows) {
  const trs = rows.map((r) => {
    let idx = 0;
    const cells = r.cells.map((c) => {
      const w = cellWidth(colWidths, idx, c.span);
      idx += c.span || 1;
      return mkCell(c, w);
    });
    return new TableRow({ children: cells, height: r.h ? { value: r.h, rule: HeightRule.ATLEAST } : undefined, tableHeader: r.header });
  });
  return new Table({ columnWidths: colWidths, width: { size: CW, type: WidthType.DXA }, rows: trs });
}

// convenience label/value
const L = (t, span) => ({ t, label: true, span, align: AlignmentType.LEFT });
const V = (t, span, align) => ({ t: t || "", span, align });

const body = [];

// ===== Title =====
body.push(new Paragraph({
  alignment: AlignmentType.CENTER,
  spacing: { after: 10, line: 240 },
  children: [run("採用リクエストフォーム", { bold: true, size: 30, color: "1F4E79" })],
}));
body.push(new Paragraph({
  alignment: AlignmentType.CENTER,
  spacing: { after: 60, line: 200 },
  border: { bottom: { style: BorderStyle.SINGLE, size: 12, color: "1F4E79", space: 3 } },
  children: [run("ADD（車いす補助業務）／WEC委託業務の内製化", { bold: true, size: 17, color: "2E5C8A" })],
}));
body.push(p(run("本フォームは、ADDのWEC外注費削減・自社運営への切替を目的に、採用条件を事前に整理・確定するためのものです。ADD責任者が必要人数・時間帯・勤務条件・WEC切替計画を作成、専務の確認後に人事部門へ採用を依頼。人事部門が市場調査のうえ、競争可能な募集単価・方法を提案し採用活動を実施します。", { size: 15, color: "444444" }), { after: 40, line: 190 }));

// ===== 1. 申請情報 =====
body.push(section("1．申請情報"));
body.push(table([1900, 3153, 1900, 3153], [
  { cells: [L("対象業務"), V("ADD（車いす補助業務）", 3)] },
  { cells: [L("作成責任者（ADD責任者）"), V("", 1), L("作成日"), V("　　　年　　月　　日")] },
  { cells: [L("部門責任者（専務）"), V("", 1), L("採用希望時期"), V("　　　年　　月頃")] },
]));

// ===== 2. ADD責任者 記入事項 =====
body.push(section("2．ADD責任者 記入事項　（業務量・配置状況・線表等に基づき記載）"));

body.push(sub("① 現在のWEC利用状況"));
body.push(table([2100, 1200, 1600, 1200, 1900, 2106], [
  { cells: [L("WEC配置人数"), V("　名／日", 1, AlignmentType.CENTER), L("委託単価"), V("　円／時", 1, AlignmentType.CENTER), L("月間委託費"), V("約　　円")] },
  { cells: [L("委託業務内容", 1), V("", 5)], h: 340 },
]));

body.push(sub("② 自社採用で必要な人数【必須】"));
body.push(table([1900, 1400, 1900, 4906], [
  { cells: [L("採用希望人数"), V("　　名", 1, AlignmentType.CENTER), L("人数算出の根拠"), V("", 1)] },
]));

body.push(sub("③ 必要な勤務時間帯【必須】　※求人には具体的な勤務時間の設定が必要"));
body.push(table([2800, 1500, 2900, 2906], [
  { cells: [{ t: "必要時間帯", label: true, align: AlignmentType.CENTER }, { t: "必要人数／日", label: true, align: AlignmentType.CENTER }, { t: "必要曜日・頻度", label: true, align: AlignmentType.CENTER }, { t: "業務内容", label: true, align: AlignmentType.CENTER }], header: true },
  { cells: [V("例：06:00～10:00"), V("5名", 1, AlignmentType.CENTER), V("毎日（Daily）", 1, AlignmentType.CENTER), V("ADD", 1, AlignmentType.CENTER)] },
  { cells: [V(""), V("", 1), V("", 1), V("", 1)], h: 300 },
  { cells: [V(""), V("", 1), V("", 1), V("", 1)], h: 300 },
]));
body.push(sep());
body.push(table([2900, 2200, 2400, 2606], [
  { cells: [L("最も人員が必要な時間帯"), V("　：　～　：", 1, AlignmentType.CENTER), L("同時間帯の最低必要人数"), V("　　名", 1, AlignmentType.CENTER)] },
]));

body.push(sub("④ 勤務条件【必須】"));
body.push(table([1700, 1300, 1400, 1500, 1200, 2006], [
  { cells: [L("1日の勤務時間"), V("　時間程度", 1, AlignmentType.CENTER), L("勤務日数"), V("週　　日以上", 1, AlignmentType.CENTER), L("土日祝勤務"), V("必須／いずれか必須／応相談／不要", 1, AlignmentType.CENTER)] },
  { cells: [L("必要な経験・資格", 1), V("", 5)], h: 300 },
  { cells: [L("必要な語学力"), V("", 2), L("その他必須条件"), V("", 2)], h: 300 },
]));

body.push(sub("⑤ WECから自社スタッフへの切替計画【必須】"));
body.push(table([1200, 1700, 1200, 3800, 2206], [
  { cells: [{ t: "段階", label: true, align: AlignmentType.CENTER }, { t: "自社スタッフ", label: true, align: AlignmentType.CENTER }, { t: "WEC", label: true, align: AlignmentType.CENTER }, { t: "切替内容", label: true, align: AlignmentType.CENTER }, { t: "目標時期", label: true, align: AlignmentType.CENTER }], header: true },
  { cells: [V("現在", 1, AlignmentType.CENTER), V("0名", 1, AlignmentType.CENTER), V("　名", 1, AlignmentType.CENTER), V("現行運用"), V("現在", 1, AlignmentType.CENTER)] },
  { cells: [V("STEP 1", 1, AlignmentType.CENTER), V("　名", 1, AlignmentType.CENTER), V("　名", 1, AlignmentType.CENTER), V(""), V("", 1, AlignmentType.CENTER)] },
  { cells: [V("STEP 2", 1, AlignmentType.CENTER), V("　名", 1, AlignmentType.CENTER), V("　名", 1, AlignmentType.CENTER), V(""), V("", 1, AlignmentType.CENTER)] },
  { cells: [V("最終", 1, AlignmentType.CENTER), V("　名", 1, AlignmentType.CENTER), V("　名", 1, AlignmentType.CENTER), V(""), V("", 1, AlignmentType.CENTER)] },
]));
body.push(sep());
body.push(table([1900, 8206], [
  { cells: [L("削減目標"), V("自社採用　　名達成 → WEC　　名削減　／　自社採用　　名達成 → WEC　　名削減")] },
]));

// ===== 3. 専務 確認事項 =====
body.push(section("3．専務 確認事項　（事業部門としての判断）"));
body.push(table([1800, 1400, 1500, 1500, 1500, 2406], [
  { cells: [L("承認する採用人数"), V("　　名", 1, AlignmentType.CENTER), L("採用期限"), V("　年　月　日", 3, AlignmentType.CENTER)] },
  { cells: [L("WEC委託単価"), V("　円／時", 1, AlignmentType.CENTER), L("許容できる時給目安"), V("　円", 1, AlignmentType.CENTER), L("採算上の上限時給"), V("　円", 1, AlignmentType.CENTER)] },
  { cells: [L("WEC削減方針", 1), V("☐ 完全内製化　　☐ 一部内製化　　☐ 段階的にWEC比率を縮小　　☐ その他", 5)] },
  { cells: [L("最終目標WEC配置人数"), V("　　名", 1, AlignmentType.CENTER), L("目標時期"), V("　年　月", 1, AlignmentType.CENTER), L("部門責任者確認"), V("専務　　　　　　／　　年　月　日", 1)] },
]));

// ===== 4. 人事部門 記入事項 =====
body.push(section("4．人事部門 記入事項　（採用市場の確認）"));
body.push(table([1800, 3253, 1800, 3253], [
  { cells: [L("市場の求人時給相場"), V("　　円 ～ 　　円", 1, AlignmentType.CENTER), L("採用難易度"), V("低／中／高／現条件では困難", 1, AlignmentType.CENTER)] },
  { cells: [L("人事推奨募集時給"), V("　　円", 1, AlignmentType.CENTER), L("想定採用期間"), V("", 1)] },
  { cells: [L("競合求人の状況", 1), V("", 3)], h: 300 },
  { cells: [L("推奨募集媒体"), V("", 1), L("推奨募集方法"), V("", 1)], h: 300 },
]));

// ===== 5. 最終募集要項 =====
body.push(section("5．最終募集要項　（採算条件＋市場調査を踏まえ確定）"));
body.push(table([1400, 1600, 1400, 1600, 1400, 1106, 1200], [
  { cells: [L("募集人数"), V("　名", 1, AlignmentType.CENTER), L("勤務時間"), V("　：　～　：", 1, AlignmentType.CENTER), L("勤務日数"), V("週　日以上", 2, AlignmentType.CENTER)] },
  { cells: [L("募集時給"), V("　円", 1, AlignmentType.CENTER), L("募集開始日"), V("　年　月　日", 1, AlignmentType.CENTER), L("採用目標日"), V("　年　月　日", 2, AlignmentType.CENTER)] },
]));

// ===== 流れ・役割 =====
body.push(section("採用開始までの流れ ／ 役割分担"));
body.push(p(run("① ADD責任者（必要人数・時間帯・勤務条件・WEC切替計画）→ ② 専務（採用人数・採用期限・人件費上限・WEC削減方針）→ ③ 人事部門（市場調査・時給/媒体/方法の提案）→ ④ 募集要項確定 → ⑤ 人事部門が採用活動開始", { size: 15, color: "1F4E79", bold: true }), { after: 24, line: 190 }));
body.push(p([run("ADD責任者：", { size: 14, bold: true }), run("「何人を・何時から何時まで必要か」「採用後にWECをどう減らすか」を具体化。", { size: 14 })], { after: 8, line: 185 }));
body.push(p([run("専務（部門責任者）：", { size: 14, bold: true }), run("「何人・いつまでに・採算上いくらまで・WECを最終的にどこまで削減するか」を事業判断として確認。", { size: 14 })], { after: 8, line: 185 }));
body.push(p([run("人事部門：", { size: 14, bold: true }), run("提示条件で採用可能かを市場調査で検証し、競争可能な時給・媒体・方法を提案、確定要項に基づき採用活動を実施。", { size: 14 })], { after: 8, line: 185 }));
body.push(p(run("※ 募集に必要な「人数・時間帯・採算条件」が未確定の場合は、当該条件の確定後に採用活動を開始します。", { size: 14, bold: true, color: "9C4A00" }), { after: 0, line: 185, fill: "FDF2E9", border: { top: bar, bottom: bar, left: bar, right: bar } }));

// ===== assemble =====
const doc = new Document({
  styles: { default: { document: { run: { font: { name: JP, eastAsia: JP }, size: 17 } } } },
  sections: [{
    properties: { page: { size: { width: 11906, height: 16838 }, margin: { top: 720, bottom: 620, left: 900, right: 900 } } },
    children: body,
  }],
});

Packer.toBuffer(doc).then((buf) => {
  fs.writeFileSync("採用リクエストフォーム.docx", buf);
  console.log("wrote", buf.length, "bytes");
});
