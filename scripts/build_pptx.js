const pptxgen = require("pptxgenjs");
const p = new pptxgen();
p.defineLayout({ name: "W", width: 13.333, height: 7.5 });
p.layout = "W";

// ---- Palette ----
const BG_DARK = "142645";   // deep navy
const NAVY    = "1E3A66";
const BLUE    = "2E5A9C";
const ICE     = "DCE7F5";
const GOLD    = "E8A33D";
const TEAL    = "1FA67A";
const PLUM    = "6D5B9C";
const WHITE   = "FFFFFF";
const INK     = "1A2233";
const MUTE    = "6B7A90";
const LIGHT   = "F4F7FB";
const CARD    = "FFFFFF";
const LINE    = "D5DEEC";

const F = "Meiryo"; // Japanese font; LibreOffice substitutes IPAGothic for QA

const W = 13.333, H = 7.5;

function bg(slide, c){ slide.background = { color: c }; }

// Section header block (light content slides)
function header(slide, kicker, title, opts={}){
  const col = opts.dark ? WHITE : INK;
  const kcol = opts.dark ? GOLD : BLUE;
  slide.addText(kicker, { x:0.7, y:0.55, w:8, h:0.3, fontFace:F, fontSize:12, bold:true, color:kcol, charSpacing:2, margin:0 });
  slide.addText(title, { x:0.7, y:0.85, w:11.9, h:0.75, fontFace:F, fontSize:32, bold:true, color:col, margin:0 });
}

function pageNum(slide, n){
  slide.addText(String(n).padStart(2,"0"), { x:12.5, y:6.95, w:0.6, h:0.35, fontFace:F, fontSize:11, color:MUTE, align:"right", margin:0 });
  slide.addText("短時間勤務者採用プロジェクト", { x:0.7, y:6.95, w:6, h:0.35, fontFace:F, fontSize:10, color:MUTE, margin:0 });
}

function chip(slide, x, y, w, h, txt, fill, txtcol, size){
  slide.addShape(p.ShapeType.roundRect, { x, y, w, h, rectRadius:0.06, fill:{color:fill}, line:{type:"none"} });
  slide.addText(txt, { x, y, w, h, fontFace:F, fontSize:size||11, bold:true, color:txtcol, align:"center", valign:"middle", margin:0 });
}

// small solid dot accent
function dot(slide, x, y, c, d){ slide.addShape(p.ShapeType.ellipse, { x, y, w:d||0.12, h:d||0.12, fill:{color:c}, line:{type:"none"} }); }

// ================= Slide 1 : Title =================
let s = p.addSlide(); bg(s, BG_DARK);
// subtle geometric motif : concentric rings top-right
s.addShape(p.ShapeType.ellipse, { x:9.6, y:-2.2, w:6.2, h:6.2, fill:{type:"none"}, line:{color:"22406E", width:1.5} });
s.addShape(p.ShapeType.ellipse, { x:10.8, y:-1.0, w:4.4, h:4.4, fill:{type:"none"}, line:{color:"2A4C82", width:1.5} });
s.addShape(p.ShapeType.ellipse, { x:12.0, y:0.2, w:2.2, h:2.2, fill:{color:GOLD}, line:{type:"none"}, shadow:{type:"outer", color:"000000", opacity:0.35, blur:12, offset:4, angle:90} });

s.addText("会社 原価改善 重点プロジェクト", { x:0.9, y:2.35, w:9, h:0.4, fontFace:F, fontSize:14, bold:true, color:GOLD, charSpacing:3, margin:0 });
s.addText("短時間勤務者採用\nプロジェクト", { x:0.85, y:2.75, w:10.5, h:1.9, fontFace:F, fontSize:48, bold:true, color:WHITE, lineSpacing:52, margin:0 });
s.addText("ピーク時間帯の人員最適化による人件費原価の改善", { x:0.9, y:4.75, w:10, h:0.5, fontFace:F, fontSize:18, color:ICE, margin:0 });
// bottom meta row
s.addShape(p.ShapeType.line, { x:0.9, y:5.7, w:5.2, h:0, line:{color:"3A5488", width:1} });
s.addText([
  { text:"対象拠点   ", options:{ color:MUTE, bold:true } },
  { text:"成田空港 / 新千歳空港", options:{ color:WHITE } },
], { x:0.9, y:5.85, w:6, h:0.35, fontFace:F, fontSize:13, margin:0 });
s.addText([
  { text:"採用目標   ", options:{ color:MUTE, bold:true } },
  { text:"合計 20名", options:{ color:GOLD, bold:true } },
], { x:0.9, y:6.25, w:6, h:0.35, fontFace:F, fontSize:13, margin:0 });
s.addNotes("会社の原価改善施策として位置付ける重点プロジェクト。短時間勤務者を採用し、ピーク時間帯の人員を最適化することで人件費原価の改善を図る。対象は成田・新千歳の2空港、採用目標は合計20名。");

// ================= Slide 2 : 目的 =================
s = p.addSlide(); bg(s, LIGHT);
header(s, "PURPOSE  ─  目的", "ピーク人員の最適化で人件費原価を改善する");
// left statement card
s.addShape(p.ShapeType.roundRect, { x:0.7, y:1.85, w:6.9, h:4.75, rectRadius:0.08, fill:{color:NAVY}, line:{type:"none"}, shadow:{type:"outer", color:"9AA7BC", opacity:0.4, blur:10, offset:3, angle:90} });
dot(s, 1.1, 2.3, GOLD, 0.16);
s.addText("プロジェクトの狙い", { x:1.4, y:2.2, w:5.8, h:0.4, fontFace:F, fontSize:15, bold:true, color:GOLD, margin:0 });
s.addText("短時間勤務者を採用し、繁忙時間帯に人員を集中配置。\n過剰・固定的な人件費を見直し、原価構造そのものを改善する。", { x:1.1, y:2.75, w:6.1, h:1.4, fontFace:F, fontSize:16, color:WHITE, lineSpacing:26, margin:0 });
s.addShape(p.ShapeType.line, { x:1.1, y:4.35, w:6.1, h:0, line:{color:"3A5488", width:1} });
s.addText("本プロジェクトは、会社の原価改善施策として位置付ける重点プロジェクトとする。", { x:1.1, y:4.6, w:6.1, h:1.2, fontFace:F, fontSize:15, italic:true, color:ICE, lineSpacing:24, margin:0 });

// right : 3 pillar rows
const pillars = [
  ["人件費原価の改善", "固定的な人員配置を見直し、コスト構造を最適化", GOLD],
  ["ピーク時間帯の最適化", "繁忙帯に合わせた柔軟な短時間シフトを構築", TEAL],
  ["重点プロジェクト化", "経営層報告のもと全社施策として推進", BLUE],
];
let py = 1.95;
pillars.forEach((r,i)=>{
  s.addShape(p.ShapeType.roundRect, { x:7.95, y:py, w:4.65, h:1.4, rectRadius:0.06, fill:{color:CARD}, line:{color:LINE, width:1} });
  s.addShape(p.ShapeType.ellipse, { x:8.2, y:py+0.32, w:0.75, h:0.75, fill:{color:r[2]}, line:{type:"none"} });
  s.addText(String(i+1), { x:8.2, y:py+0.32, w:0.75, h:0.75, fontFace:F, fontSize:22, bold:true, color:WHITE, align:"center", valign:"middle", margin:0 });
  s.addText(r[0], { x:9.15, y:py+0.24, w:3.35, h:0.4, fontFace:F, fontSize:15, bold:true, color:INK, margin:0 });
  s.addText(r[1], { x:9.15, y:py+0.66, w:3.35, h:0.6, fontFace:F, fontSize:11.5, color:MUTE, lineSpacing:15, margin:0 });
  py += 1.6;
});
pageNum(s, 2);
s.addNotes("目的は人件費原価の改善。短時間勤務者の採用でピーク帯人員を最適化する。全社の原価改善における重点プロジェクトとして推進する。");

// ================= Slide 3 : 採用目標 =================
s = p.addSlide(); bg(s, BG_DARK);
s.addText("TARGET  ─  採用目標", { x:0.7, y:0.55, w:8, h:0.3, fontFace:F, fontSize:12, bold:true, color:GOLD, charSpacing:2, margin:0 });
s.addText("2空港で合計20名を採用する", { x:0.7, y:0.85, w:11.9, h:0.75, fontFace:F, fontSize:32, bold:true, color:WHITE, margin:0 });

// two big stat cards + total
function statCard(x, label, num, unit, accent){
  s.addShape(p.ShapeType.roundRect, { x, y:2.15, w:4.05, h:3.7, rectRadius:0.08, fill:{color:NAVY}, line:{color:"2E4A78", width:1} });
  s.addShape(p.ShapeType.rect, { x:x+0.35, y:2.6, w:0.55, h:0.14, fill:{color:accent}, line:{type:"none"} });
  s.addText(label, { x:x+0.35, y:2.85, w:3.35, h:0.5, fontFace:F, fontSize:19, bold:true, color:WHITE, margin:0 });
  s.addText([
    { text:num, options:{ fontSize:86, bold:true, color:accent } },
    { text:unit, options:{ fontSize:26, bold:true, color:WHITE } },
  ], { x:x+0.3, y:3.5, w:3.45, h:1.8, fontFace:F, align:"center", valign:"middle", margin:0 });
}
statCard(0.9, "成田空港", "10", "名", GOLD);
statCard(5.2, "新千歳空港", "10", "名", TEAL);
// total card
s.addShape(p.ShapeType.roundRect, { x:9.5, y:2.15, w:3.0, h:3.7, rectRadius:0.08, fill:{color:GOLD}, line:{type:"none"}, shadow:{type:"outer", color:"000000", opacity:0.3, blur:12, offset:4, angle:90} });
s.addText("採用目標 合計", { x:9.65, y:2.85, w:2.7, h:0.5, fontFace:F, fontSize:16, bold:true, color:"5A3A00", margin:0 });
s.addText([
  { text:"20", options:{ fontSize:90, bold:true, color:INK } },
  { text:"名", options:{ fontSize:26, bold:true, color:INK } },
], { x:9.5, y:3.5, w:3.0, h:1.8, fontFace:F, align:"center", valign:"middle", margin:0 });
// footnote
s.addText("繁忙時間帯の人員を短時間勤務で補完し、拠点ごとに10名ずつ採用する。", { x:0.9, y:6.2, w:11.5, h:0.5, fontFace:F, fontSize:14, color:ICE, margin:0 });
pageNum(s, 3);
s.addNotes("採用目標は成田空港10名、新千歳空港10名の合計20名。");

// ================= Slide 4 : プロジェクト体制 =================
s = p.addSlide(); bg(s, LIGHT);
header(s, "TEAM  ─  プロジェクト体制", "5つの役割で推進体制を構築");
const team = [
  ["統括", "遠藤", "全体統括・経営報告", NAVY],
  ["PMリーダー", "内山", "プロジェクト運営管理", BLUE],
  ["実務担当", "野澤", "採用実務の中心", TEAL],
  ["制度・労務", "髙田 部長", "制度設計・労務確認", GOLD],
  ["現場調整窓口", "坂本", "入社後フォロー・調整", PLUM],
];
let tx = 0.51;
const cw = 2.31, gap = 0.19;
team.forEach((m,i)=>{
  s.addShape(p.ShapeType.roundRect, { x:tx, y:2.0, w:cw, h:4.35, rectRadius:0.08, fill:{color:CARD}, line:{color:LINE, width:1}, shadow:{type:"outer", color:"9AA7BC", opacity:0.35, blur:8, offset:2, angle:90} });
  // top color band via rounded rect header
  s.addShape(p.ShapeType.roundRect, { x:tx, y:2.0, w:cw, h:0.95, rectRadius:0.08, fill:{color:m[3]}, line:{type:"none"} });
  s.addShape(p.ShapeType.rect, { x:tx, y:2.55, w:cw, h:0.4, fill:{color:m[3]}, line:{type:"none"} });
  s.addText(m[0], { x:tx+0.15, y:2.05, w:cw-0.3, h:0.85, fontFace:F, fontSize:14, bold:true, color:WHITE, align:"center", valign:"middle", margin:0 });
  // avatar circle
  s.addShape(p.ShapeType.ellipse, { x:tx+cw/2-0.6, y:3.2, w:1.2, h:1.2, fill:{color:LIGHT}, line:{color:m[3], width:2} });
  s.addText(m[1].charAt(0), { x:tx+cw/2-0.6, y:3.2, w:1.2, h:1.2, fontFace:F, fontSize:34, bold:true, color:m[3], align:"center", valign:"middle", margin:0 });
  s.addText(m[1], { x:tx+0.1, y:4.55, w:cw-0.2, h:0.5, fontFace:F, fontSize:20, bold:true, color:INK, align:"center", margin:0 });
  s.addShape(p.ShapeType.line, { x:tx+0.6, y:5.15, w:cw-1.2, h:0, line:{color:LINE, width:1} });
  s.addText(m[2], { x:tx+0.15, y:5.3, w:cw-0.3, h:0.9, fontFace:F, fontSize:12.5, color:MUTE, align:"center", lineSpacing:17, margin:0 });
  tx += cw + gap;
});
pageNum(s, 4);
s.addNotes("統括：遠藤、プロジェクトリーダー：内山、実務担当：野澤、制度・労務担当：髙田部長、現場調整窓口：坂本。坂本は入社〜入社後のフォローとトラブル対応・現場調整の窓口を担う。実務は野澤を中心に運営し、内山が運営を管理する。");

// ================= Slide 5 : 募集媒体・掲載期間 =================
s = p.addSlide(); bg(s, LIGHT);
header(s, "MEDIA  ─  募集媒体・掲載期間", "空港・媒体別の掲載スケジュール");
// timeline chart (gantt-style)
const X0 = 3.55, TW = 8.85; // chart x-range
const range = 47; // 7/20 -> 9/5
function dayIdx(m,d){ if(m===7) return d-20; if(m===8) return 11+d; return 42+d; }
function xd(m,d){ return X0 + dayIdx(m,d)/range*TW; }
const chartTop = 2.35, rowH = 0.95, barH = 0.5;
// month gridlines
const marks = [[8,1,"8/1"],[8,15,"8/15"],[9,1,"9/1"]];
marks.forEach(mk=>{
  const x = xd(mk[0],mk[1]);
  s.addShape(p.ShapeType.line, { x, y:chartTop-0.15, w:0, h:3.2, line:{color:LINE, width:1, dashType:"dash"} });
  s.addText(mk[2], { x:x-0.4, y:chartTop-0.55, w:0.8, h:0.3, fontFace:F, fontSize:11, color:MUTE, align:"center", margin:0 });
});
const bars = [
  ["成田空港", "バイトル", 7,24, 8,30, GOLD],
  ["新千歳空港", "バイトル", 7,24, 8,30, TEAL],
  ["新千歳空港", "マイナビバイト", 7,30, 9,2, BLUE],
];
bars.forEach((b,i)=>{
  const y = chartTop + i*rowH;
  // left labels
  s.addText(b[0], { x:0.7, y:y-0.02, w:1.7, h:0.32, fontFace:F, fontSize:12.5, bold:true, color:INK, margin:0 });
  s.addText(b[1], { x:0.7, y:y+0.28, w:2.7, h:0.3, fontFace:F, fontSize:11, color:MUTE, margin:0 });
  // track
  s.addShape(p.ShapeType.roundRect, { x:X0, y:y+0.05, w:TW, h:barH, rectRadius:0.05, fill:{color:"E7EDF6"}, line:{type:"none"} });
  const bx = xd(b[2],b[3]), bx2 = xd(b[4],b[5]);
  s.addShape(p.ShapeType.roundRect, { x:bx, y:y+0.05, w:bx2-bx, h:barH, rectRadius:0.05, fill:{color:b[6]}, line:{type:"none"} });
  s.addText(`${b[2]}/${b[3]}〜${b[4]}/${b[5]}`, { x:bx, y:y+0.05, w:bx2-bx, h:barH, fontFace:F, fontSize:11.5, bold:true, color:WHITE, align:"center", valign:"middle", margin:0 });
});
// legend / summary chips
s.addShape(p.ShapeType.line, { x:0.7, y:5.75, w:11.9, h:0, line:{color:LINE, width:1} });
chip(s, 0.7, 6.0, 2.6, 0.55, "成田：1媒体", NAVY, WHITE, 12);
chip(s, 3.45, 6.0, 2.6, 0.55, "新千歳：2媒体", NAVY, WHITE, 12);
s.addText("マイナビバイト（新千歳）は 9/2 まで掲載し、後半の応募獲得を強化。", { x:6.3, y:6.0, w:6.3, h:0.55, fontFace:F, fontSize:12.5, color:MUTE, valign:"middle", margin:0 });
pageNum(s, 5);
s.addNotes("成田はバイトル（7/24〜8/30）。新千歳はバイトル（7/24〜8/30）とマイナビバイト（7/30〜9/2）の2媒体を併用する。");

// ================= Slide 6 : 担当業務 =================
s = p.addSlide(); bg(s, LIGHT);
header(s, "ROLES  ─  担当業務", "メンバー別の主な担当業務");
const duties = [
  ["遠藤", "統括", NAVY, ["プロジェクト全体統括","採用目標管理・進捗確認","課題抽出、改善指示","経営層への報告"]],
  ["内山", "PMリーダー", BLUE, ["プロジェクト運営","現場との調整","面接対応","進捗管理"]],
  ["野澤", "実務担当", TEAL, ["ネオキャリアとの窓口・調整","求人媒体管理・原稿修正","応募者管理・面接日程調整","面接対応・採用進捗管理"]],
  ["髙田 部長", "制度・労務", GOLD, ["制度設計","労務確認","雇用条件確認","社会保険等の制度確認"]],
  ["坂本", "現場調整窓口", PLUM, ["入社〜入社後のフォロー","現場調整の窓口対応","トラブル・相談対応","定着支援・早期離職の防止"]],
];
let dx = 0.51;
const dw = 2.31, dgap = 0.19;
duties.forEach(d=>{
  s.addShape(p.ShapeType.roundRect, { x:dx, y:1.95, w:dw, h:4.55, rectRadius:0.08, fill:{color:CARD}, line:{color:LINE, width:1}, shadow:{type:"outer", color:"9AA7BC", opacity:0.3, blur:7, offset:2, angle:90} });
  s.addShape(p.ShapeType.ellipse, { x:dx+0.28, y:2.2, w:0.5, h:0.5, fill:{color:d[2]}, line:{type:"none"} });
  s.addText(d[0], { x:dx+0.9, y:2.18, w:dw-1.0, h:0.35, fontFace:F, fontSize:16, bold:true, color:INK, margin:0 });
  s.addText(d[1], { x:dx+0.9, y:2.52, w:dw-1.0, h:0.28, fontFace:F, fontSize:11, bold:true, color:d[2], margin:0 });
  s.addShape(p.ShapeType.line, { x:dx+0.28, y:3.0, w:dw-0.56, h:0, line:{color:LINE, width:1} });
  const items = d[3].map((t,i)=>({ text:t, options:{ bullet:{ code:"2022", indent:11 }, color:INK, breakLine:true, paraSpaceAfter:7 } }));
  s.addText(items, { x:dx+0.26, y:3.15, w:dw-0.46, h:3.2, fontFace:F, fontSize:11, color:INK, lineSpacing:14, valign:"top", margin:0 });
  dx += dw + dgap;
});
pageNum(s, 6);
s.addNotes("遠藤は全体統括・経営報告。内山は運営・調整・進捗管理。野澤はネオキャリア窓口や媒体・応募者・面接の実務全般。髙田部長は制度設計・労務・雇用条件・社会保険の確認。坂本は入社〜入社後のフォロー、現場調整の窓口、トラブル・相談対応、定着支援を担う。");

// ================= Slide 7 : 面接体制 =================
s = p.addSlide(); bg(s, LIGHT);
header(s, "INTERVIEW  ─  面接体制", "人事2名＋現場管理職1名で面接を実施");
// common structure banner
s.addShape(p.ShapeType.roundRect, { x:0.7, y:1.95, w:11.9, h:0.85, rectRadius:0.07, fill:{color:NAVY}, line:{type:"none"} });
s.addText([
  { text:"面接官の基本構成    ", options:{ color:GOLD, bold:true, fontSize:14 } },
  { text:"人事（内山・野澤） ＋ 現場管理職 1名（各拠点の候補者から必ず1名が同席）", options:{ color:WHITE, fontSize:14 } },
], { x:1.05, y:1.95, w:11.3, h:0.85, fontFace:F, valign:"middle", margin:0 });

function interviewCol(x, title, accent, people){
  s.addShape(p.ShapeType.roundRect, { x, y:3.05, w:5.85, h:3.4, rectRadius:0.08, fill:{color:CARD}, line:{color:LINE, width:1}, shadow:{type:"outer", color:"9AA7BC", opacity:0.3, blur:7, offset:2, angle:90} });
  s.addShape(p.ShapeType.roundRect, { x, y:3.05, w:5.85, h:0.7, rectRadius:0.08, fill:{color:accent}, line:{type:"none"} });
  s.addShape(p.ShapeType.rect, { x, y:3.4, w:5.85, h:0.35, fill:{color:accent}, line:{type:"none"} });
  s.addText(title, { x:x+0.3, y:3.05, w:5.3, h:0.7, fontFace:F, fontSize:16, bold:true, color:WHITE, valign:"middle", margin:0 });
  s.addText("現場管理職 候補（いずれか1名が同席）", { x:x+0.3, y:3.9, w:5.3, h:0.3, fontFace:F, fontSize:11.5, bold:true, color:MUTE, margin:0 });
  // people chips grid 2 columns
  let cy = 4.3;
  people.forEach((pp,i)=>{
    const col = i%2, rowi = Math.floor(i/2);
    const cx = x + 0.3 + col*2.75;
    const yy = 4.3 + rowi*0.62;
    s.addShape(p.ShapeType.roundRect, { x:cx, y:yy, w:2.55, h:0.5, rectRadius:0.05, fill:{color:LIGHT}, line:{color:LINE, width:1} });
    s.addShape(p.ShapeType.ellipse, { x:cx+0.12, y:yy+0.15, w:0.2, h:0.2, fill:{color:accent}, line:{type:"none"} });
    s.addText(pp, { x:cx+0.42, y:yy, w:2.05, h:0.5, fontFace:F, fontSize:12.5, bold:true, color:INK, valign:"middle", margin:0 });
  });
}
interviewCol(0.7, "成田空港", GOLD, ["ロシャン","古永 MGR","佐藤 いずみ","村上 MGR"]);
interviewCol(6.75, "新千歳空港", TEAL, ["佐藤 部長","山下 さん"]);
pageNum(s, 7);
s.addNotes("面接は人事（内山・野澤）に現場管理職1名が必ず同席。成田はロシャン・古永MGR・佐藤いずみ・村上MGRのいずれか。新千歳は佐藤部長・山下さんのいずれか。");

// ================= Slide 8 : KPI =================
s = p.addSlide(); bg(s, BG_DARK);
s.addText("KPI  ─  進捗管理項目", { x:0.7, y:0.55, w:8, h:0.3, fontFace:F, fontSize:12, bold:true, color:GOLD, charSpacing:2, margin:0 });
s.addText("採用ファネルで進捗を可視化・管理する", { x:0.7, y:0.85, w:11.9, h:0.75, fontFace:F, fontSize:32, bold:true, color:WHITE, margin:0 });
// funnel steps (bars kept left of the right-hand panel at x=8.5)
const funnel = [
  ["応募数","エントリー獲得", 7.1],
  ["面接設定数","日程確定", 6.2],
  ["面接実施数","面接完了", 5.3],
  ["採用数","内定・採用決定", 4.4],
  ["入社数","勤務開始", 3.5],
];
const fY = 2.15, fH = 0.72, fGap = 0.19, fX = 0.9;
const fcols = [BLUE,"2F6FA8",TEAL,"1FA67A",GOLD];
funnel.forEach((fn,i)=>{
  const y = fY + i*(fH+fGap);
  const w = fn[2];
  s.addShape(p.ShapeType.roundRect, { x:fX, y, w, h:fH, rectRadius:0.05, fill:{color:fcols[i]}, line:{type:"none"} });
  s.addText(`${i+1}`, { x:fX+0.15, y, w:0.55, h:fH, fontFace:F, fontSize:20, bold:true, color:"FFFFFF", align:"center", valign:"middle", margin:0 });
  s.addText(fn[0], { x:fX+0.75, y, w:w-2.45, h:fH, fontFace:F, fontSize:16, bold:true, color:WHITE, valign:"middle", margin:0 });
  // descriptor inside the bar, right-aligned
  s.addText(fn[1], { x:fX+w-1.7, y, w:1.55, h:fH, fontFace:F, fontSize:11, color:"EAF1FB", align:"right", valign:"middle", margin:0 });
});
// right side : cross-cut KPIs
s.addShape(p.ShapeType.roundRect, { x:8.5, y:2.15, w:4.0, h:4.35, rectRadius:0.08, fill:{color:NAVY}, line:{color:"2E4A78", width:1} });
s.addText("集計の切り口", { x:8.8, y:2.4, w:3.4, h:0.4, fontFace:F, fontSize:15, bold:true, color:GOLD, margin:0 });
const cuts = [["媒体別応募数","バイトル / マイナビバイト"],["空港別進捗","成田 / 新千歳"]];
let cyy = 3.05;
cuts.forEach(c=>{
  dot(s, 8.85, cyy+0.08, GOLD, 0.14);
  s.addText(c[0], { x:9.1, y:cyy-0.02, w:3.2, h:0.35, fontFace:F, fontSize:14, bold:true, color:WHITE, margin:0 });
  s.addText(c[1], { x:9.1, y:cyy+0.32, w:3.2, h:0.35, fontFace:F, fontSize:11.5, color:ICE, margin:0 });
  cyy += 1.0;
});
s.addShape(p.ShapeType.line, { x:8.8, y:5.15, w:3.4, h:0, line:{color:"3A5488", width:1} });
s.addText("各段階の歩留まりを追い、応募〜入社の転換率を改善する。", { x:8.8, y:5.35, w:3.5, h:1.0, fontFace:F, fontSize:12, color:ICE, lineSpacing:17, margin:0 });
pageNum(s, 8);
s.addNotes("KPIは応募数・面接設定数・面接実施数・採用数・入社数のファネルで管理。加えて媒体別応募数と空港別進捗（成田・千歳）を集計する。");

// ================= Slide 9 : 基本方針 =================
s = p.addSlide(); bg(s, BG_DARK);
s.addShape(p.ShapeType.ellipse, { x:-2.0, y:4.5, w:6.0, h:6.0, fill:{type:"none"}, line:{color:"22406E", width:1.5} });
s.addText("PRINCIPLES  ─  基本方針", { x:0.9, y:0.75, w:8, h:0.3, fontFace:F, fontSize:12, bold:true, color:GOLD, charSpacing:2, margin:0 });
s.addText("役割分担を徹底し、連携して推進する", { x:0.9, y:1.1, w:11.5, h:0.75, fontFace:F, fontSize:32, bold:true, color:WHITE, margin:0 });
const princ = [
  ["実務の中心", "実務は野澤を中心に運営する", TEAL],
  ["運営管理", "プロジェクト運営は内山が管理する", BLUE],
  ["全体統括", "遠藤は全体統括・進捗管理・経営報告および課題解決に注力", GOLD],
  ["制度・労務連携", "制度・労務面は髙田部長と連携しながら進める", ICE],
];
let ry = 2.35;
princ.forEach((r,i)=>{
  s.addShape(p.ShapeType.roundRect, { x:0.9, y:ry, w:11.5, h:0.95, rectRadius:0.07, fill:{color:NAVY}, line:{color:"2E4A78", width:1} });
  s.addShape(p.ShapeType.ellipse, { x:1.2, y:ry+0.22, w:0.5, h:0.5, fill:{color:r[2]}, line:{type:"none"} });
  s.addText(String(i+1), { x:1.2, y:ry+0.22, w:0.5, h:0.5, fontFace:F, fontSize:18, bold:true, color:i===3?INK:WHITE, align:"center", valign:"middle", margin:0 });
  s.addText(r[0], { x:1.95, y:ry, w:2.7, h:0.95, fontFace:F, fontSize:16, bold:true, color:r[2], valign:"middle", margin:0 });
  s.addText(r[1], { x:4.7, y:ry, w:7.4, h:0.95, fontFace:F, fontSize:15, color:WHITE, valign:"middle", margin:0 });
  ry += 1.08;
});
pageNum(s, 9);
s.addNotes("基本方針：実務は野澤が中心、運営管理は内山、全体統括・経営報告・課題解決は遠藤、制度・労務は髙田部長と連携。");

const path = require("path");
const out = path.join(__dirname, "..", "採用プロジェクト_資料.pptx");
p.writeFile({ fileName: out }).then(f=>console.log("WROTE", f));
