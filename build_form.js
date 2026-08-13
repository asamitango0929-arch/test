const fs = require("fs");
const {
  Document, Packer, Paragraph, TextRun, HeadingLevel, AlignmentType,
  Table, TableRow, TableCell, WidthType, BorderStyle, ShadingType, ShadingType: SH,
} = require("docx");

const JP = "Yu Gothic";
const CONTENT_W = 9026; // A4 (11906) minus 1440 margins each side

// ---------- helpers ----------
const run = (text, opts = {}) =>
  new TextRun({ text, font: { name: JP, eastAsia: JP }, size: opts.size || 21, bold: !!opts.bold, color: opts.color, ...opts });

const para = (children, opts = {}) =>
  new Paragraph({
    children: Array.isArray(children) ? children : [children],
    spacing: { after: opts.after == null ? 120 : opts.after, before: opts.before || 0, line: 276 },
    alignment: opts.alignment,
    ...opts.p,
  });

// blank fill-in line
const blank = (n = 12) => run("　".repeat(n), { underline: {} });

// label (bold) + blank + trailing unit
const field = (label, opts = {}) => {
  const kids = [run(label, { bold: true })];
  if (opts.blank !== false) kids.push(blank(opts.n || 10));
  if (opts.suffix) kids.push(run(opts.suffix));
  return para(kids, { after: opts.after == null ? 100 : opts.after });
};

const spacer = (after = 60) => new Paragraph({ children: [run("")], spacing: { after } });

const h1 = (text) =>
  new Paragraph({
    heading: HeadingLevel.HEADING_1,
    spacing: { before: 320, after: 160 },
    border: { bottom: { style: BorderStyle.SINGLE, size: 12, color: "1F4E79", space: 4 } },
    children: [run(text, { bold: true, size: 30, color: "1F4E79" })],
  });

const h2 = (text) =>
  new Paragraph({
    heading: HeadingLevel.HEADING_2,
    spacing: { before: 220, after: 120 },
    children: [run(text, { bold: true, size: 24, color: "2E5C8A" })],
  });

const bar = { style: BorderStyle.SINGLE, size: 4, color: "BFBFBF" };
const cellBorders = { top: bar, bottom: bar, left: bar, right: bar };

const th = (text, w) =>
  new TableCell({
    width: { size: w, type: WidthType.DXA },
    shading: { type: SH.CLEAR, fill: "1F4E79", color: "auto" },
    margins: { top: 60, bottom: 60, left: 90, right: 90 },
    borders: cellBorders,
    children: [para(run(text, { bold: true, color: "FFFFFF", size: 20 }), { after: 0, alignment: AlignmentType.CENTER })],
  });

const td = (text, w, opts = {}) =>
  new TableCell({
    width: { size: w, type: WidthType.DXA },
    margins: { top: 60, bottom: 60, left: 90, right: 90 },
    borders: cellBorders,
    shading: opts.fill ? { type: SH.CLEAR, fill: opts.fill, color: "auto" } : undefined,
    children: [para(run(text || "", { size: 20, color: opts.color }), { after: 0, alignment: opts.align || AlignmentType.LEFT })],
  });

const makeTable = (widths, rows) =>
  new Table({
    columnWidths: widths,
    width: { size: CONTENT_W, type: WidthType.DXA },
    rows,
  });

// ---------- document body ----------
const body = [];

// Title block
body.push(
  new Paragraph({
    alignment: AlignmentType.CENTER,
    spacing: { after: 60 },
    children: [run("採用リクエストフォーム", { bold: true, size: 40, color: "1F4E79" })],
  }),
  new Paragraph({
    alignment: AlignmentType.CENTER,
    spacing: { after: 40 },
    children: [run("ADD（車いす補助業務）／WEC委託業務の内製化", { bold: true, size: 24, color: "2E5C8A" })],
  }),
  new Paragraph({
    alignment: AlignmentType.CENTER,
    spacing: { after: 200 },
    border: { bottom: { style: BorderStyle.SINGLE, size: 18, color: "1F4E79", space: 6 } },
    children: [run("")],
  }),
);

// 目的
body.push(h2("本フォームの目的"));
body.push(para(run("本フォームは、ADD（車いす補助業務）について、WECへの外注費削減および自社運営への切替を目的として、採用活動に必要な条件を事前に整理・確定するためのものです。")));
body.push(para([
  run("ADD責任者が、現場運営に必要な人数・時間帯・勤務条件・WECからの切替計画等を作成し、部門責任者である専務の確認後、人事部門へ採用を依頼します。", { bold: true }),
]));
body.push(para(run("人事部門は、提示された採用要件および採算条件をもとに求人市場・競合求人を調査し、採用市場で競争可能な募集単価・募集方法を提案した上で、採用活動を実施します。"), { after: 160 }));

// 1. 申請情報
body.push(h1("1．申請情報"));
body.push(para([run("対象業務：", { bold: true }), run("ADD（車いす補助業務）")]));
body.push(field("作成責任者：ADD責任者　", { n: 14 }));
body.push(field("部門責任者：専務　", { n: 14 }));
body.push(para([run("作成日：", { bold: true }), blank(4), run("年"), blank(2), run("月"), blank(2), run("日")]));
body.push(para([run("採用希望時期：", { bold: true }), blank(4), run("年"), blank(2), run("月頃")], { after: 120 }));

// 2. ADD責任者 記入事項
body.push(h1("2．ADD責任者 記入事項"));
body.push(para(run("以下について、実際の業務量・配置状況・線表等に基づき記載してください。"), { after: 140 }));

// ①
body.push(h2("① 現在のWEC利用状況"));
body.push(para([run("現在のWEC配置人数：", { bold: true }), blank(6), run("名／日")]));
body.push(para([run("WEC委託単価：", { bold: true }), blank(6), run("円／時間")]));
body.push(para([run("月間委託費：約", { bold: true }), blank(8), run("円")]));
body.push(para(run("WECへ委託している業務内容：", { bold: true }), { after: 30 }));
body.push(spacer()); body.push(spacer(160));

// ②
body.push(h2("② 自社採用で必要な人数【必須】"));
body.push(para([run("採用希望人数：", { bold: true }), blank(6), run("名", { bold: true })]));
body.push(para(run("人数算出の根拠：", { bold: true }), { after: 30 }));
body.push(spacer()); body.push(spacer(160));

// ③
body.push(h2("③ 必要な勤務時間帯【必須】"));
body.push(para(run("※求人募集を行うためには、具体的な勤務時間の設定が必要です。", { size: 19, color: "C00000" }), { after: 100 }));
{
  const w = [2700, 1600, 2626, 2100];
  const rows = [
    new TableRow({ tableHeader: true, children: [th("必要時間帯", w[0]), th("必要人数／日", w[1]), th("必要曜日・頻度", w[2]), th("業務内容", w[3])] }),
    new TableRow({ children: [td("例：06:00～10:00", w[0], { fill: "F2F6FB" }), td("5名", w[1], { align: AlignmentType.CENTER, fill: "F2F6FB" }), td("毎日（Daily）", w[2], { fill: "F2F6FB" }), td("ADD", w[3], { fill: "F2F6FB" })] }),
    new TableRow({ children: [td("　", w[0]), td("", w[1]), td("", w[2]), td("", w[3])] }),
    new TableRow({ children: [td("　", w[0]), td("", w[1]), td("", w[2]), td("", w[3])] }),
    new TableRow({ children: [td("　", w[0]), td("", w[1]), td("", w[2]), td("", w[3])] }),
  ];
  body.push(makeTable(w, rows));
}
body.push(spacer(80));
body.push(para([run("最も人員が必要な時間帯：", { bold: true }), blank(3), run("："), blank(2), run("　～　"), blank(3), run("："), blank(2)]));
body.push(para([run("同時間帯の最低必要人数：", { bold: true }), blank(6), run("名", { bold: true })], { after: 160 }));

// ④
body.push(h2("④ 勤務条件【必須】"));
body.push(para([run("1日の勤務時間：", { bold: true }), blank(6), run("時間程度")]));
body.push(para([run("勤務日数：週", { bold: true }), blank(4), run("日以上")]));
body.push(para([run("土日祝勤務：", { bold: true }), run("必須　／　土日のいずれか必須　／　応相談　／　不要")]));
body.push(para(run("必要な経験・資格：", { bold: true }), { after: 30 }));
body.push(spacer(80));
body.push(para(run("必要な語学力：", { bold: true }), { after: 30 }));
body.push(spacer(80));
body.push(para(run("その他必須条件：", { bold: true }), { after: 30 }));
body.push(spacer()); body.push(spacer(160));

// ⑤
body.push(h2("⑤ WECから自社スタッフへの切替計画【必須】"));
body.push(para(run("自社採用の進捗に応じて、WECへの委託をどのように縮小するのか記載してください。"), { after: 100 }));
{
  const w = [1400, 1700, 1200, 3026, 1700];
  const rows = [
    new TableRow({ tableHeader: true, children: [th("段階", w[0]), th("自社スタッフ", w[1]), th("WEC", w[2]), th("切替内容", w[3]), th("目標時期", w[4])] }),
    new TableRow({ children: [td("現在", w[0], { align: AlignmentType.CENTER, fill: "F2F6FB" }), td("0名", w[1], { align: AlignmentType.CENTER, fill: "F2F6FB" }), td("　名", w[2], { align: AlignmentType.CENTER, fill: "F2F6FB" }), td("現行運用", w[3], { fill: "F2F6FB" }), td("現在", w[4], { align: AlignmentType.CENTER, fill: "F2F6FB" })] }),
    new TableRow({ children: [td("STEP 1", w[0], { align: AlignmentType.CENTER }), td("　名", w[1], { align: AlignmentType.CENTER }), td("　名", w[2], { align: AlignmentType.CENTER }), td("", w[3]), td("", w[4])] }),
    new TableRow({ children: [td("STEP 2", w[0], { align: AlignmentType.CENTER }), td("　名", w[1], { align: AlignmentType.CENTER }), td("　名", w[2], { align: AlignmentType.CENTER }), td("", w[3]), td("", w[4])] }),
    new TableRow({ children: [td("最終", w[0], { align: AlignmentType.CENTER }), td("　名", w[1], { align: AlignmentType.CENTER }), td("　名", w[2], { align: AlignmentType.CENTER }), td("", w[3]), td("", w[4])] }),
  ];
  body.push(makeTable(w, rows));
}
body.push(spacer(80));
body.push(para([run("自社採用", { bold: true }), blank(4), run("名達成 → WEC", { bold: true }), blank(4), run("名削減", { bold: true })]));
body.push(para([run("自社採用", { bold: true }), blank(4), run("名達成 → WEC", { bold: true }), blank(4), run("名削減", { bold: true })], { after: 160 }));

// 3. 専務 確認事項
body.push(h1("3．専務 確認事項"));
body.push(para(run("ADD責任者が作成した運用計画をもとに、事業部門として以下をご確認ください。"), { after: 140 }));

body.push(h2("① 採用人数"));
body.push(para([run("承認する採用人数：", { bold: true }), blank(6), run("名", { bold: true })], { after: 140 }));

body.push(h2("② 採用期限"));
body.push(para([run("いつまでに必要か：", { bold: true }), blank(4), run("年"), blank(2), run("月"), blank(2), run("日")], { after: 140 }));

body.push(h2("③ 採算上の人件費"));
body.push(para([run("現在のWEC委託単価：", { bold: true }), blank(6), run("円／時間")]));
body.push(para([run("自社採用で許容できる時給目安：", { bold: true }), blank(6), run("円")]));
body.push(para([run("採算上の上限時給：", { bold: true }), blank(6), run("円")], { after: 140 }));

body.push(h2("④ WEC削減方針"));
body.push(para([run("☐　", { size: 22 }), run("完全内製化を目指す")], { after: 60 }));
body.push(para([run("☐　", { size: 22 }), run("一部内製化")], { after: 60 }));
body.push(para([run("☐　", { size: 22 }), run("段階的にWEC比率を縮小")], { after: 60 }));
body.push(para([run("☐　", { size: 22 }), run("その他")], { after: 100 }));
body.push(para([run("最終的に目指すWEC配置人数：", { bold: true }), blank(6), run("名")]));
body.push(para([run("目標時期：", { bold: true }), blank(4), run("年"), blank(2), run("月")], { after: 140 }));

body.push(h2("⑤ 部門責任者確認"));
body.push(para(run("上記の内容に基づき、人事部門へ採用を依頼します。")));
body.push(field("部門責任者：専務　", { n: 14 }));
body.push(para([run("確認日：", { bold: true }), blank(4), run("年"), blank(2), run("月"), blank(2), run("日")], { after: 120 }));

// 4. 人事部門 記入事項
body.push(h1("4．人事部門 記入事項"));
body.push(para(run("ADD責任者および専務から提示された条件をもとに、人事部門にて採用市場を確認します。"), { after: 140 }));
body.push(para([run("市場の求人時給相場：", { bold: true }), blank(6), run("円 ～ "), blank(6), run("円")]));
body.push(para(run("競合求人の状況：", { bold: true }), { after: 30 }));
body.push(spacer(80));
body.push(para([run("採用難易度：", { bold: true }), run("低　／　中　／　高　／　現条件では困難")]));
body.push(para([run("人事推奨募集時給：", { bold: true }), blank(6), run("円")]));
body.push(para(run("推奨募集媒体：", { bold: true }), { after: 30 }));
body.push(spacer(80));
body.push(para(run("推奨募集方法：", { bold: true }), { after: 30 }));
body.push(spacer(80));
body.push(para(run("想定採用期間：", { bold: true }), { after: 30 }));
body.push(spacer()); body.push(spacer(160));

// 5. 最終募集要項
body.push(h1("5．最終募集要項"));
body.push(para(run("事業部門の採算条件と人事部門による市場調査を踏まえ、最終的な募集要項を確定します。"), { after: 140 }));
body.push(para([run("募集人数：", { bold: true }), blank(6), run("名")]));
body.push(para([run("勤務時間：", { bold: true }), blank(3), run("："), blank(2), run("　～　"), blank(3), run("："), blank(2)]));
body.push(para([run("勤務日数：週", { bold: true }), blank(4), run("日以上")]));
body.push(para([run("募集時給：", { bold: true }), blank(6), run("円")]));
body.push(para([run("募集開始日：", { bold: true }), blank(4), run("年"), blank(2), run("月"), blank(2), run("日")]));
body.push(para([run("採用目標日：", { bold: true }), blank(4), run("年"), blank(2), run("月"), blank(2), run("日")], { after: 160 }));

// 採用開始までの流れ
body.push(h1("採用開始までの流れ"));
const flow = [
  ["① ADD責任者", "必要人数・必要時間帯・勤務条件・WEC切替計画を作成"],
  ["② 専務（部門責任者）", "採用人数・採用期限・採算上の人件費上限・WEC削減方針を確認"],
  ["③ 人事部門", "求人市場・競合求人を調査し、採用市場で競争できる時給・募集方法を提案"],
  ["④ 募集要項確定", ""],
  ["⑤ 人事部門が採用活動開始", ""],
];
flow.forEach((step, i) => {
  const kids = [run(step[0], { bold: true, color: "1F4E79" })];
  if (step[1]) kids.push(run("　—　" + step[1]));
  body.push(new Paragraph({
    spacing: { after: 40, before: 40 },
    shading: { type: SH.CLEAR, fill: "EAF1FA", color: "auto" },
    border: { top: bar, bottom: bar, left: bar, right: bar },
    children: kids,
  }));
  if (i < flow.length - 1) {
    body.push(new Paragraph({ alignment: AlignmentType.CENTER, spacing: { after: 20, before: 20 }, children: [run("↓", { bold: true, size: 24, color: "2E5C8A" })] }));
  }
});
body.push(spacer(160));

// 役割分担
body.push(h1("役割分担"));
const roles = [
  ["ADD責任者", "「何人を、何時から何時まで必要とするのか」「採用後にWECをどのように減らすのか」を具体化します。"],
  ["専務（部門責任者）", "「何人採用するのか」「いつまでに必要なのか」「採算上いくらまで人件費を投入できるのか」「WECを最終的にどこまで削減するのか」を事業判断として確認します。"],
  ["人事部門", "提示された条件に対して市場調査を行い、「その条件で採用できるのか」を検証します。その上で、競争可能な時給・求人媒体・募集方法を提案し、確定した募集要項に基づき採用活動を実施します。"],
];
roles.forEach(([role, desc]) => {
  body.push(para(run(role, { bold: true, size: 22, color: "2E5C8A" }), { after: 30 }));
  body.push(para(run(desc), { after: 140 }));
});
body.push(new Paragraph({
  spacing: { before: 100, after: 100 },
  shading: { type: SH.CLEAR, fill: "FDF2E9", color: "auto" },
  border: { top: bar, bottom: bar, left: bar, right: bar },
  children: [run("募集に必要な「人数・時間帯・採算条件」が確定していない場合は、当該条件の確定後に採用活動を開始します。", { bold: true, color: "9C4A00" })],
}));

// ---------- assemble ----------
const doc = new Document({
  styles: {
    default: {
      document: { run: { font: { name: JP, eastAsia: JP }, size: 21 } },
    },
  },
  sections: [{
    properties: {
      page: {
        size: { width: 11906, height: 16838 }, // A4 portrait
        margin: { top: 1440, bottom: 1440, left: 1440, right: 1440 },
      },
    },
    children: body,
  }],
});

Packer.toBuffer(doc).then((buf) => {
  fs.writeFileSync("採用リクエストフォーム.docx", buf);
  console.log("wrote 採用リクエストフォーム.docx", buf.length, "bytes");
});
