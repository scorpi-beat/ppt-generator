/**
 * pptxgen_builder.js — Phase 2 PPTX 빌더 (pptxgenjs 기반)
 *
 * 사용법:
 *   node src/build/pptxgen_builder.js \
 *     --draft outputs/draft_{type}_{slug}.json \
 *     [--style outputs/style_{type}.json] \
 *     [--out   outputs/{type}_{slug}.pptx]
 *
 * ─── 규격 변경 방법 ──────────────────────────────────────────────────
 *   CFG.emuW / CFG.emuH 만 수정 → 모든 좌표·크기 자동 재계산
 *   폰트: FONT 객체 (style JSON 로드 시 자동 덮어씀)
 *         폰트는 절대 pt값이므로 슬라이드 크기와 무관하게 유지됨
 * ─────────────────────────────────────────────────────────────────────
 */

"use strict";
const PptxGenJS = require("pptxgenjs");
const fs        = require("fs");
const path      = require("path");

// ═══════════════════════════════════════════════════════════════════════
// [1] SLIDE CONFIG  ← 여기만 변경하면 전체 레이아웃이 자동으로 재계산됨
// ═══════════════════════════════════════════════════════════════════════
const CFG = {
  emuW: 12192000,  // 960pt × 12700 — HTML 프리뷰 기준, component_template_report 동일
  emuH:  6858000,  // 540pt × 12700

  // ─ H 대비 분율 ─────────────────────────────────────────────────────
  headerHf:      0.052,   // 헤더바 높이
  bannerHf:      0.048,   // head_message 배너 높이
  footerHf:      0.028,   // 각주/note 높이
  contentTopPad: 0.012,   // 배너 아래 패딩
  contentBotPad: 0.008,   // 콘텐츠 하단 패딩
  rowHf:         0.040,   // 기본 표 행 높이

  // ─ W 대비 분율 ─────────────────────────────────────────────────────
  marginXf: 0.015,        // 좌우 마진
  gapf:     0.008,        // 존 간 간격
};

// 파생 상수 ─ CFG에서 자동 계산 (직접 수정 금지)
const W         = CFG.emuW / 914400;
const H         = CFG.emuH / 914400;
const MX        = W * CFG.marginXf;
const GAP       = W * CFG.gapf;
const HEADER_H  = H * CFG.headerHf;
const BANNER_H  = H * CFG.bannerHf;
const FOOTER_H  = H * CFG.footerHf;
const CONTENT_Y = HEADER_H + BANNER_H + H * CFG.contentTopPad;
const CONTENT_H = H - CONTENT_Y - FOOTER_H - H * CFG.contentBotPad;
const CONTENT_W = W - MX * 2;
const ROW_H     = H * CFG.rowHf;

// ═══════════════════════════════════════════════════════════════════════
// [2] FONT — 절대 pt값, style JSON 로드 시 덮어씌워짐
// ═══════════════════════════════════════════════════════════════════════
let FONT = {
  face:       "Pretendard",
  headMsg:    14,
  header:     13,
  body:        9,
  tableH:      9,
  tableB:      9,
  footnote:    6,
  kpiVal:     26,
  kpiLbl:      9,
  sectionN:   22,
  sectionT:   16,
  titleMain:  30,
  titleSub:   16,
  tocItem:    10,
};

// ═══════════════════════════════════════════════════════════════════════
// [3] COLOR — style JSON 로드 시 덮어씌워짐
// ═══════════════════════════════════════════════════════════════════════
let C = {
  primary:  "627365",
  accent:   "2D3734",
  salmon:   "D98F76",
  warm:     "A09567",
  blue:     "406D92",
  text:     "232323",
  white:    "FFFFFF",
  neutral:  "F2F2F2",
  negative: "C0392B",
};

// ─── style JSON 로드 ───────────────────────────────────────────────────
function loadStyle(stylePath) {
  if (!stylePath || !fs.existsSync(stylePath)) return;
  try {
    const s = JSON.parse(fs.readFileSync(stylePath, "utf-8"));
    const col = s.colors || {};
    const fnt = s.fonts  || {};
    if (col.primary)       C.primary  = col.primary.replace("#", "");
    if (col.accent_dark)   C.accent   = col.accent_dark.replace("#", "");
    if (col.accent_salmon) C.salmon   = col.accent_salmon.replace("#", "");
    if (col.accent_warm)   C.warm     = col.accent_warm.replace("#", "");
    if (col.accent_blue)   C.blue     = col.accent_blue.replace("#", "");
    if (col.text_main)     C.text     = col.text_main.replace("#", "");
    if (fnt.primary_typeface) FONT.face    = fnt.primary_typeface;
    if (fnt.head_message?.size_pt)     FONT.headMsg = fnt.head_message.size_pt;
    if (fnt.header_bar_title?.size_pt) FONT.header  = fnt.header_bar_title.size_pt;
    if (fnt.body?.size_pt)             FONT.body    = Math.max(fnt.body.size_pt, 8);
    if (fnt.table_body?.size_pt)       FONT.tableB  = Math.max(fnt.table_body.size_pt, 8);
    if (fnt.table_header?.size_pt)     FONT.tableH  = Math.max(fnt.table_header.size_pt, 8);
    if (fnt.footnote_source?.size_pt)  FONT.footnote = Math.max(fnt.footnote_source.size_pt, 6);
    if (fnt.head_message?.color) {
      const hc = fnt.head_message.color.replace("#","");
      if (/^[0-9A-Fa-f]{6}$/.test(hc)) C.salmon = hc;
    }
  } catch (e) { /* style 로드 실패 시 기본값 유지 */ }
}

// ═══════════════════════════════════════════════════════════════════════
// [4] HELPERS
// ═══════════════════════════════════════════════════════════════════════

function addCommonHeader(slide, s, pageNum) {
  slide.addShape("rect", { x:0, y:0, w:W, h:HEADER_H,
    fill:{color:C.primary}, line:{color:C.primary} });

  const hasSection = !!s.section_number;
  if (hasSection) {
    slide.addText(`Section ${s.section_number}`, {
      x:MX, y:HEADER_H*0.05, w:W-1.5, h:HEADER_H*0.32,
      fontSize:7, color:"FFFFFF", fontFace:FONT.face,
      transparency:35, valign:"top",
    });
  }
  slide.addText(s.title || "", {
    x:MX, y: hasSection ? HEADER_H*0.34 : HEADER_H*0.12,
    w:W-1.5, h:HEADER_H * (hasSection ? 0.60 : 0.76),
    fontSize:FONT.header, color:"FFFFFF", bold:true,
    fontFace:FONT.face, valign:"middle",
  });
  slide.addText(String(pageNum), {
    x:W-0.9, y:0, w:0.75, h:HEADER_H,
    fontSize:8, color:"FFFFFF", fontFace:FONT.face,
    align:"right", valign:"middle", transparency:40,
  });

  if (s.head_message) {
    slide.addShape("rect", { x:0, y:HEADER_H, w:W, h:BANNER_H,
      fill:{color:"FFFFFF"}, line:{color:"FFFFFF"} });
    slide.addShape("rect", {
      x:MX, y:HEADER_H + BANNER_H*0.1, w:0.04, h:BANNER_H*0.8,
      fill:{color:C.salmon}, line:{color:C.salmon},
    });
    slide.addText(s.head_message, {
      x:MX+0.10, y:HEADER_H+BANNER_H*0.06,
      w:W-MX*2-0.10, h:BANNER_H*0.88,
      fontSize:FONT.headMsg, color:C.salmon, fontFace:FONT.face,
      valign:"middle", wrap:true,
    });
  }
}

function addNoteBox(slide, noteBox) {
  if (!noteBox?.content) return;
  const ny = H - FOOTER_H + H*0.004;
  slide.addText(noteBox.content, {
    x:MX, y:ny, w:W-MX*2, h:FOOTER_H*0.8,
    fontSize:FONT.footnote, color:"808080", fontFace:FONT.face,
    italic: noteBox.type === "source", valign:"middle",
  });
}

function makeTableRows(tbl, opts = {}) {
  const { hdrBg = C.primary, stripe = true } = opts;
  const hdr = (tbl.headers || []).map((h, i) => ({
    text: String(h),
    options: {
      bold:true, fontSize:FONT.tableH, color:C.white,
      fontFace:FONT.face, align: i===0 ? "left" : "center",
      fill:{color:hdrBg}, margin:[2,4,2,4],
    },
  }));
  const rows = (tbl.rows || []).map((row, ri) =>
    row.map((cell, ci) => ({
      text: String(cell ?? ""),
      options: {
        fontSize:FONT.tableB, color:C.text, fontFace:FONT.face,
        align: ci===0 ? "left" : "center",
        bold: ci===0,
        fill:{color: stripe && ri%2===0 ? C.neutral : C.white},
        margin:[2,4,2,4],
      },
    }))
  );
  return [hdr, ...rows];
}

function normalizeChart(chart) {
  if (!chart) return { labels:[], seriesList:[] };
  const labels=[], values=[];
  if (Array.isArray(chart.data)) {
    chart.data.forEach(d => { labels.push(String(d.label)); values.push(Number(d.value)); });
    return { labels, seriesList:[{ name:chart.title||"값", values }] };
  }
  if (Array.isArray(chart.series)) {
    const first = chart.series[0] || {};
    if ("label" in first || "value" in first) {
      chart.series.forEach(d => { labels.push(String(d.label)); values.push(Number(d.value)); });
      return { labels, seriesList:[{ name:chart.title||"값", values }] };
    }
    if ("name" in first) {
      const cats = chart.categories || first.labels || [];
      return { labels:cats.map(String), seriesList:chart.series.map(s=>({ name:s.name, values:(s.values||[]).map(Number) })) };
    }
  }
  return { labels:[], seriesList:[{ name:"값", values:[] }] };
}

function buildWaterfallData(chart) {
  const { labels, seriesList } = normalizeChart(chart);
  const vals = seriesList[0]?.values || [];
  const src   = chart.series || chart.data || [];
  const base=[], pos=[], neg=[];
  let run = 0;
  vals.forEach((v, i) => {
    const type = src[i]?.type;
    if (type === "absolute" || type === "total") {
      base.push(0); pos.push(v>0?v:0); neg.push(v<0?-v:0); run=v;
    } else {
      base.push(v>=0 ? run : run+v);
      pos.push(v>=0 ? v : 0);
      neg.push(v<0  ? -v: 0);
      run += v;
    }
  });
  return { labels, seriesList:[
    { name:"_base", values:base },
    { name:"증가",   values:pos  },
    { name:"감소",   values:neg  },
  ]};
}

function addChartToSlide(prs, slide, chart, x, y, w, h) {
  if (!chart) return;
  const type = (chart.chart_type||"bar").toLowerCase();
  let chartData, ctype, extra={};

  if (type === "waterfall") {
    const wf = buildWaterfallData(chart);
    chartData = wf.seriesList.map(s=>({ name:s.name, labels:wf.labels, values:s.values }));
    ctype = prs.ChartType.bar;
    extra = { barGrouping:"stacked", chartColors:[C.white, C.primary, C.negative],
              showLegend:true, legendPos:"b", legendFontSize:7 };
  } else {
    const { labels, seriesList } = normalizeChart(chart);
    chartData = seriesList.map(s=>({ name:s.name, labels, values:s.values }));
    ctype = prs.ChartType[type] || prs.ChartType.bar;
    extra = { chartColors: seriesList.length>1
      ? [C.primary, C.blue, C.warm, C.accent, C.salmon]
      : [C.primary] };
  }

  const CHART_COMMON = {
    x, y, w, h,
    showTitle: !!chart.title,
    title: chart.title || "",
    titleFontSize: FONT.footnote+1,
    titleColor: C.accent,
    titleFontFace: FONT.face,
    showLegend: chartData.length > 1,
    legendFontSize: 7,
    legendPos: "b",
    catAxisLabelFontSize: FONT.footnote+1,
    valAxisLabelFontSize: FONT.footnote+1,
    catAxisLabelColor: C.text,
    valAxisLabelColor: C.text,
    valGridLine: { style:"none" },
    catGridLine: { style:"none" },
    showValue: type==="pie"||type==="doughnut",
    dataLabelFontSize: FONT.footnote,
  };
  slide.addChart(ctype, chartData, { ...CHART_COMMON, ...extra });
}

function addBulletsToSlide(slide, bullets, x, y, w, h) {
  if (!bullets?.length) return;
  const items = bullets.map(b => ({
    text: String(b),
    options: {
      fontSize:FONT.body, color:C.text, fontFace:FONT.face,
      bullet:{ type:"bullet", code:"25AA", indent:10 },
      paraSpaceAfter:3,
    },
  }));
  slide.addText(items, { x, y, w, h, valign:"top", wrap:true });
}

function addKeyPoints(slide, keyPoints, y) {
  if (!keyPoints?.length) return;
  const kpY = y + H*0.01;
  slide.addShape("line", { x:MX, y:kpY-H*0.005, w:CONTENT_W, h:0,
    line:{ color:"DDDDDD", width:0.5 } });
  const items = keyPoints.map(k => ({
    text: String(k),
    options:{ fontSize:FONT.body, color:C.text, fontFace:FONT.face,
              bullet:{ type:"bullet", code:"25CF", indent:10 }, paraSpaceAfter:3 },
  }));
  const kpH = H - kpY - FOOTER_H - H*CFG.contentBotPad;
  slide.addText(items, { x:MX, y:kpY, w:CONTENT_W, h:kpH, valign:"top", wrap:true });
}

function renderZoneContent(prs, slide, zone, x, y, w, h) {
  if (!zone) return;
  const ct = (zone.content_type||"").toLowerCase();
  if (ct==="bullets"||ct==="bullet") {
    addBulletsToSlide(slide, zone.bullets||zone.body||[], x, y, w, h);
  } else if (ct==="chart") {
    addChartToSlide(prs, slide, zone.chart, x, y, w, h);
  } else if (ct==="table") {
    if (zone.table) {
      const rows = makeTableRows(zone.table);
      const rh = Math.min(ROW_H, h/(zone.table.rows.length+1.5));
      slide.addTable(rows, { x, y, w, rowH:rh, border:{type:"solid",color:"DDDDDD",pt:0.4} });
    }
  } else if (ct==="text") {
    const txt = Array.isArray(zone.body)?zone.body.join("\n"):(zone.text||zone.description||"");
    slide.addText(txt, { x, y, w, h, fontSize:FONT.body, color:C.text, fontFace:FONT.face, wrap:true, valign:"top" });
  } else if (ct==="callout") {
    slide.addShape("rect", { x, y, w, h, fill:{color:C.neutral}, line:{color:"DDDDDD",pt:0.5} });
    slide.addText(zone.value||"", { x:x+0.1, y:y+h*0.15, w:w-0.2, h:h*0.4,
      fontSize:FONT.kpiVal, color:C.primary, fontFace:FONT.face, align:"center", bold:true });
    slide.addText(zone.description||zone.label||"", { x:x+0.1, y:y+h*0.58, w:w-0.2, h:h*0.35,
      fontSize:FONT.body, color:C.text, fontFace:FONT.face, align:"center", wrap:true });
  } else if (ct==="process"||ct==="diagram") {
    slide.addShape("rect", { x, y, w, h, fill:{color:C.neutral}, line:{color:"DDDDDD",pt:0.5} });
    const desc = zone.description||zone.text||"다이어그램";
    slide.addText(desc, { x:x+0.12, y:y+0.1, w:w-0.24, h:h-0.2,
      fontSize:FONT.body, color:C.text, fontFace:FONT.face, wrap:true, valign:"top" });
  }
}

// ═══════════════════════════════════════════════════════════════════════
// [5] LAYOUT RENDERERS
// ═══════════════════════════════════════════════════════════════════════

// ── title_slide ──────────────────────────────────────────────────────
function renderTitleSlide(prs, slide, s) {
  slide.addShape("rect", { x:0, y:0, w:W, h:H, fill:{color:C.accent}, line:{color:C.accent} });
  slide.addShape("rect", { x:0, y:H*0.72, w:W, h:H*0.28, fill:{color:C.primary}, line:{color:C.primary} });
  slide.addShape("rect", { x:MX, y:H*0.68, w:W*0.12, h:H*0.008, fill:{color:C.salmon}, line:{color:C.salmon} });

  slide.addText(s.title||"", { x:MX, y:H*0.20, w:W-MX*2, h:H*0.38,
    fontSize:FONT.titleMain, color:C.white, bold:true, fontFace:FONT.face, valign:"bottom", wrap:true });
  slide.addText(s.subtitle||s.sub_title||"", { x:MX, y:H*0.60, w:W-MX*2, h:H*0.08,
    fontSize:FONT.titleSub, color:C.white, fontFace:FONT.face, valign:"middle" });

  const meta = [s.date, s.organization||s.author].filter(Boolean).join("   |   ");
  slide.addText(meta, { x:MX, y:H*0.78, w:W-MX*2, h:H*0.12,
    fontSize:FONT.body, color:C.white, fontFace:FONT.face, valign:"middle", transparency:20 });
}

// ── toc_slide ────────────────────────────────────────────────────────
function renderTocSlide(prs, slide, s) {
  addCommonHeader(slide, s, "");
  const sections = s.sections || [];
  const colW  = CONTENT_W / Math.min(sections.length, 3) - GAP;
  sections.forEach((sec, i) => {
    const col  = i % 3;
    const row  = Math.floor(i / 3);
    const sx   = MX + col*(colW+GAP);
    const sy   = CONTENT_Y + row*(CONTENT_H/2);
    const sh   = CONTENT_H/2 - GAP;
    slide.addShape("rect", { x:sx, y:sy, w:colW, h:sh, fill:{color:C.neutral}, line:{color:"DDDDDD",pt:0.5} });
    slide.addText(sec.number||String(i+1), { x:sx+0.12, y:sy+0.10, w:colW-0.2, h:sh*0.30,
      fontSize:FONT.sectionN, color:C.primary, bold:true, fontFace:FONT.face });
    slide.addText(sec.title||"", { x:sx+0.12, y:sy+sh*0.35, w:colW-0.2, h:sh*0.25,
      fontSize:FONT.tocItem+1, color:C.accent, bold:true, fontFace:FONT.face, wrap:true });
    if (sec.subsections?.length) {
      const subs = sec.subsections.map(sb=>({ text:sb,
        options:{ fontSize:FONT.footnote+1, color:C.text, fontFace:FONT.face,
                  bullet:{type:"bullet",code:"25CF",indent:8}, paraSpaceAfter:2 } }));
      slide.addText(subs, { x:sx+0.12, y:sy+sh*0.62, w:colW-0.2, h:sh*0.32, valign:"top", wrap:true });
    }
  });
}

// ── section_divider ──────────────────────────────────────────────────
function renderSectionDivider(prs, slide, s) {
  slide.addShape("rect", { x:0, y:0, w:W, h:H, fill:{color:C.accent}, line:{color:C.accent} });
  slide.addShape("rect", { x:0, y:H*0.70, w:W, h:H*0.30, fill:{color:C.primary}, line:{color:C.primary} });
  slide.addShape("rect", { x:MX, y:H*0.66, w:W*0.10, h:H*0.008, fill:{color:C.salmon}, line:{color:C.salmon} });
  slide.addText(s.section_number||"", { x:MX, y:H*0.12, w:W-MX*2, h:H*0.38,
    fontSize:FONT.sectionN*2.5, color:C.white, bold:true, fontFace:FONT.face, transparency:25 });
  slide.addText(s.title||"", { x:MX, y:H*0.42, w:W-MX*2, h:H*0.24,
    fontSize:FONT.sectionT*1.4, color:C.white, bold:true, fontFace:FONT.face, wrap:true });
  if (s.subtitle) slide.addText(s.subtitle, { x:MX, y:H*0.64, w:W-MX*2, h:H*0.08,
    fontSize:FONT.titleSub, color:C.white, fontFace:FONT.face, transparency:20 });
}

// ── content_text ─────────────────────────────────────────────────────
function renderContentText(prs, slide, s, n) {
  addCommonHeader(slide, s, n);
  const bullets = s.body || s.bullets || [];
  addBulletsToSlide(slide, bullets, MX, CONTENT_Y, CONTENT_W, CONTENT_H);
  addNoteBox(slide, s.note_box);
}

// ── content_chart ────────────────────────────────────────────────────
function renderContentChart(prs, slide, s, n) {
  addCommonHeader(slide, s, n);
  const chart = s.chart;
  const kp    = s.key_points || [];
  const chartH = kp.length ? CONTENT_H*0.72 : CONTENT_H*0.95;
  addChartToSlide(prs, slide, chart, MX, CONTENT_Y, CONTENT_W, chartH);
  if (kp.length) addKeyPoints(slide, kp, CONTENT_Y+chartH);
  addNoteBox(slide, s.note_box);
}

// ── table_slide ──────────────────────────────────────────────────────
function renderTableSlide(prs, slide, s, n) {
  addCommonHeader(slide, s, n);
  const tbl = s.table;
  const kp  = s.key_points || [];
  if (tbl) {
    const tableH = kp.length ? CONTENT_H*0.62 : CONTENT_H*0.95;
    const rows   = makeTableRows(tbl);
    const rh     = Math.min(ROW_H, tableH/(tbl.rows.length+1.5));
    slide.addTable(rows, { x:MX, y:CONTENT_Y, w:CONTENT_W, rowH:rh,
      border:{type:"solid",color:"DDDDDD",pt:0.4} });
    if (kp.length) addKeyPoints(slide, kp, CONTENT_Y+tableH);
  }
  addNoteBox(slide, s.note_box);
}

// ── wide_table ───────────────────────────────────────────────────────
function renderWideTable(prs, slide, s, n) {
  addCommonHeader(slide, s, n);
  const tbl = s.table;
  const kp  = s.key_points || [];
  if (tbl) {
    const nCols  = (tbl.headers||[]).length;
    const fstW   = CONTENT_W*0.18;
    const restW  = (CONTENT_W-fstW)/(nCols-1||1);
    const colW   = [fstW, ...Array(Math.max(nCols-1,0)).fill(restW)];
    const tableH = kp.length ? CONTENT_H*0.60 : CONTENT_H*0.92;
    const rows   = makeTableRows(tbl);
    const rh     = Math.min(ROW_H*0.90, tableH/(tbl.rows.length+1.5));
    slide.addTable(rows, { x:MX, y:CONTENT_Y, w:CONTENT_W, rowH:rh, colW,
      border:{type:"solid",color:"DDDDDD",pt:0.4} });
    if (kp.length) addKeyPoints(slide, kp, CONTENT_Y+tableH);
  }
  addNoteBox(slide, s.note_box);
}

// ── kpi_metrics ──────────────────────────────────────────────────────
function renderKpiMetrics(prs, slide, s, n) {
  addCommonHeader(slide, s, n);
  const kpis  = s.kpis || s.metrics || [];
  const count = Math.min(kpis.length, 4);
  if (!count) return;
  const cardW = (CONTENT_W - GAP*(count-1)) / count;
  const cardH = CONTENT_H*0.88;
  kpis.slice(0,count).forEach((kpi, i) => {
    const cx = MX + i*(cardW+GAP);
    slide.addShape("rect", { x:cx, y:CONTENT_Y, w:cardW, h:cardH,
      fill:{color:C.neutral}, line:{color:"DDDDDD",pt:0.5} });
    slide.addShape("rect", { x:cx, y:CONTENT_Y, w:cardW, h:cardH*0.04,
      fill:{color:C.primary}, line:{color:C.primary} });
    slide.addText(kpi.value||"", { x:cx+0.1, y:CONTENT_Y+cardH*0.12, w:cardW-0.2, h:cardH*0.38,
      fontSize:FONT.kpiVal, color:C.primary, bold:true, fontFace:FONT.face, align:"center" });
    if (kpi.unit) slide.addText(kpi.unit, { x:cx+0.1, y:CONTENT_Y+cardH*0.13, w:cardW-0.2, h:cardH*0.38,
      fontSize:FONT.body, color:C.warm, fontFace:FONT.face, align:"center", valign:"bottom" });
    slide.addText(kpi.label||kpi.title||"", { x:cx+0.1, y:CONTENT_Y+cardH*0.52, w:cardW-0.2, h:cardH*0.22,
      fontSize:FONT.kpiLbl, color:C.text, fontFace:FONT.face, align:"center", wrap:true });
    if (kpi.delta) slide.addText(kpi.delta, { x:cx+0.1, y:CONTENT_Y+cardH*0.74, w:cardW-0.2, h:cardH*0.18,
      fontSize:FONT.body, color:C.blue, fontFace:FONT.face, align:"center" });
  });
  addNoteBox(slide, s.note_box);
}

// ── two_col_text_table ───────────────────────────────────────────────
function renderTwoColTextTable(prs, slide, s, n) {
  addCommonHeader(slide, s, n);
  const lw  = CONTENT_W*0.44;
  const rw  = CONTENT_W - lw - GAP;
  const rx  = MX + lw + GAP;
  const bul = s.bullets || s.body || [];
  addBulletsToSlide(slide, bul, MX, CONTENT_Y, lw, CONTENT_H);
  if (s.table) {
    const rows = makeTableRows(s.table);
    const rh   = Math.min(ROW_H, CONTENT_H/(s.table.rows.length+1.5));
    slide.addTable(rows, { x:rx, y:CONTENT_Y, w:rw, rowH:rh,
      border:{type:"solid",color:"DDDDDD",pt:0.4} });
  }
  addNoteBox(slide, s.note_box);
}

// ── two_col_text_chart ───────────────────────────────────────────────
function renderTwoColTextChart(prs, slide, s, n) {
  addCommonHeader(slide, s, n);
  const lw = CONTENT_W*0.40;
  const rw = CONTENT_W - lw - GAP;
  const rx = MX + lw + GAP;
  addBulletsToSlide(slide, s.bullets||s.body||[], MX, CONTENT_Y, lw, CONTENT_H);
  addChartToSlide(prs, slide, s.chart, rx, CONTENT_Y, rw, CONTENT_H*0.95);
  addNoteBox(slide, s.note_box);
}

// ── two_col_chart_text ───────────────────────────────────────────────
function renderTwoColChartText(prs, slide, s, n) {
  addCommonHeader(slide, s, n);
  const lw = CONTENT_W*0.56;
  const rw = CONTENT_W - lw - GAP;
  const rx = MX + lw + GAP;
  addChartToSlide(prs, slide, s.chart, MX, CONTENT_Y, lw, CONTENT_H*0.95);
  addBulletsToSlide(slide, s.bullets||s.body||[], rx, CONTENT_Y, rw, CONTENT_H);
  addNoteBox(slide, s.note_box);
}

// ── two_column_compare ───────────────────────────────────────────────
function renderTwoColumnCompare(prs, slide, s, n) {
  addCommonHeader(slide, s, n);
  const hw = (CONTENT_W - GAP) / 2;
  const rx = MX + hw + GAP;
  const left  = s.left  || s.columns?.[0] || {};
  const right = s.right || s.columns?.[1] || {};

  [[left, MX], [right, rx]].forEach(([col, x]) => {
    if (col.title) {
      slide.addShape("rect", { x, y:CONTENT_Y, w:hw, h:H*0.036,
        fill:{color:C.primary}, line:{color:C.primary} });
      slide.addText(col.title, { x:x+0.1, y:CONTENT_Y, w:hw-0.2, h:H*0.036,
        fontSize:FONT.body, color:C.white, fontFace:FONT.face, valign:"middle" });
    }
    const ty = col.title ? CONTENT_Y+H*0.046 : CONTENT_Y;
    const th = CONTENT_H - (col.title ? H*0.046 : 0);
    addBulletsToSlide(slide, col.bullets||col.body||[], x, ty, hw, th);
  });
  addNoteBox(slide, s.note_box);
}

// ── table_chart_combo ────────────────────────────────────────────────
function renderTableChartCombo(prs, slide, s, n) {
  addCommonHeader(slide, s, n);
  const lw = CONTENT_W*0.52;
  const rw = CONTENT_W - lw - GAP;
  const rx = MX + lw + GAP;
  if (s.table) {
    const rows = makeTableRows(s.table);
    const rh   = Math.min(ROW_H, CONTENT_H/(s.table.rows.length+1.5));
    slide.addTable(rows, { x:MX, y:CONTENT_Y, w:lw, rowH:rh,
      border:{type:"solid",color:"DDDDDD",pt:0.4} });
  }
  addChartToSlide(prs, slide, s.chart, rx, CONTENT_Y, rw, CONTENT_H*0.95);
  addNoteBox(slide, s.note_box);
}

// ── three_column_summary ─────────────────────────────────────────────
function renderThreeColumnSummary(prs, slide, s, n) {
  addCommonHeader(slide, s, n);
  const cols = s.columns || s.cards || [];
  const count = Math.min(cols.length, 3);
  if (!count) return;
  const cw = (CONTENT_W - GAP*(count-1)) / count;
  cols.slice(0,count).forEach((col, i) => {
    const cx = MX + i*(cw+GAP);
    slide.addShape("rect", { x:cx, y:CONTENT_Y, w:cw, h:CONTENT_H,
      fill:{color:C.neutral}, line:{color:"DDDDDD",pt:0.5} });
    slide.addShape("rect", { x:cx, y:CONTENT_Y, w:cw, h:H*0.04,
      fill:{color:C.primary}, line:{color:C.primary} });
    slide.addText(col.title||"", { x:cx+0.12, y:CONTENT_Y+H*0.005, w:cw-0.24, h:H*0.04,
      fontSize:FONT.body, color:C.white, bold:true, fontFace:FONT.face, valign:"middle" });
    if (col.highlight) slide.addText(col.highlight, { x:cx+0.12, y:CONTENT_Y+H*0.052, w:cw-0.24, h:H*0.07,
      fontSize:FONT.kpiLbl*1.6, color:C.primary, bold:true, fontFace:FONT.face, align:"center" });
    const byH = col.highlight ? CONTENT_Y+H*0.125 : CONTENT_Y+H*0.052;
    const bH  = CONTENT_H - (col.highlight ? H*0.128 : H*0.055);
    addBulletsToSlide(slide, col.bullets||col.body||[], cx+0.12, byH, cw-0.24, bH);
  });
  addNoteBox(slide, s.note_box);
}

// ── composite_split ──────────────────────────────────────────────────
function renderCompositeSplit(prs, slide, s, n) {
  addCommonHeader(slide, s, n);
  const mainPos = s.main_zone?.position || "left";
  const mainW   = CONTENT_W * 0.54;
  const subW    = CONTENT_W - mainW - GAP;
  const mainX   = mainPos==="left" ? MX : MX+subW+GAP;
  const subX    = mainPos==="left" ? MX+mainW+GAP : MX;
  const subGap  = CONTENT_H * 0.03;
  const topH    = (s.sub_zone_top && s.sub_zone_bottom)
                  ? CONTENT_H*0.50 - subGap/2
                  : CONTENT_H;
  const botH    = CONTENT_H - topH - subGap;

  renderZoneContent(prs, slide, s.main_zone, mainX, CONTENT_Y, mainW, CONTENT_H);
  if (s.sub_zone_top)
    renderZoneContent(prs, slide, s.sub_zone_top, subX, CONTENT_Y, subW, topH);
  if (s.sub_zone_bottom)
    renderZoneContent(prs, slide, s.sub_zone_bottom, subX, CONTENT_Y+topH+subGap, subW, botH);
  addNoteBox(slide, s.note_box);
}

// ── four_quadrant ────────────────────────────────────────────────────
function renderFourQuadrant(prs, slide, s, n) {
  addCommonHeader(slide, s, n);
  const cells = s.cells || [];
  const hw = (CONTENT_W - GAP) / 2;
  const hh = (CONTENT_H - GAP) / 2;
  const positions = {
    top_left:     [MX,       CONTENT_Y],
    top_right:    [MX+hw+GAP,CONTENT_Y],
    bottom_left:  [MX,       CONTENT_Y+hh+GAP],
    bottom_right: [MX+hw+GAP,CONTENT_Y+hh+GAP],
  };
  cells.forEach(cell => {
    const pos = positions[cell.position] || positions.top_left;
    const [cx, cy] = pos;
    slide.addShape("rect", { x:cx, y:cy, w:hw, h:hh,
      fill:{color:C.neutral}, line:{color:"DDDDDD",pt:0.5} });
    if (cell.label) {
      slide.addShape("rect", { x:cx, y:cy, w:hw, h:H*0.038,
        fill:{color:C.primary}, line:{color:C.primary} });
      slide.addText(cell.label, { x:cx+0.1, y:cy, w:hw-0.2, h:H*0.038,
        fontSize:FONT.body, color:C.white, bold:true, fontFace:FONT.face, valign:"middle" });
    }
    const innerY = cell.label ? cy+H*0.046 : cy+H*0.02;
    const innerH = hh - (cell.label ? H*0.056 : H*0.03);
    renderZoneContent(prs, slide, cell, cx+0.1, innerY, hw-0.2, innerH);
  });
  addNoteBox(slide, s.note_box);
}

// ── process_flow ─────────────────────────────────────────────────────
function renderProcessFlow(prs, slide, s, n) {
  addCommonHeader(slide, s, n);
  const steps = s.steps || [];
  const count = Math.min(steps.length, 6);
  if (!count) return;
  const arrowW = CONTENT_W * 0.02;
  const stepW  = (CONTENT_W - arrowW*(count-1)) / count;
  const stepH  = CONTENT_H * 0.80;
  const stepY  = CONTENT_Y + CONTENT_H*0.10;

  steps.slice(0,count).forEach((step, i) => {
    const sx = MX + i*(stepW+arrowW);
    const isEven = i%2===0;
    slide.addShape("rect", { x:sx, y:stepY, w:stepW, h:stepH,
      fill:{color: isEven ? C.primary : C.neutral},
      line:{color: isEven ? C.primary : "CCCCCC", pt:0.5} });
    slide.addText(String(step.number||i+1), { x:sx+0.1, y:stepY+stepH*0.04, w:stepW-0.2, h:stepH*0.18,
      fontSize:FONT.kpiLbl*1.5, color: isEven ? C.white : C.primary, bold:true, fontFace:FONT.face, align:"center" });
    slide.addText(step.title||"", { x:sx+0.1, y:stepY+stepH*0.22, w:stepW-0.2, h:stepH*0.20,
      fontSize:FONT.body, color: isEven ? C.white : C.accent, bold:true, fontFace:FONT.face, align:"center", wrap:true });
    const items = (step.items||[]).map(t=>({
      text:t, options:{ fontSize:FONT.footnote+1, color: isEven ? "FFFFFF" : C.text, fontFace:FONT.face,
        bullet:{type:"bullet",code:"25CF",indent:6}, paraSpaceAfter:2 } }));
    if (items.length) slide.addText(items, { x:sx+0.1, y:stepY+stepH*0.44, w:stepW-0.2, h:stepH*0.52, valign:"top", wrap:true });
    if (i<count-1) {
      const ax = sx+stepW;
      slide.addText("→", { x:ax, y:stepY+stepH*0.38, w:arrowW, h:stepH*0.24,
        fontSize:FONT.body*1.2, color:C.warm, fontFace:FONT.face, align:"center", valign:"middle" });
    }
  });
  addNoteBox(slide, s.note_box);
}

// ── roadmap_timeline ────────────────────────────────────────────────
function renderRoadmapTimeline(prs, slide, s, n) {
  addCommonHeader(slide, s, n);
  const nodes = s.nodes || s.timeline || [];
  const count = Math.min(nodes.length, 7);
  if (!count) return;
  const nodeSpacing = CONTENT_W / count;
  const lineY  = CONTENT_Y + CONTENT_H*0.38;
  const lineH  = H*0.006;
  slide.addShape("rect", { x:MX, y:lineY, w:CONTENT_W, h:lineH, fill:{color:C.primary}, line:{color:C.primary} });

  nodes.slice(0,count).forEach((node, i) => {
    const nx = MX + nodeSpacing*(i+0.5);
    const dotR = CONTENT_H*0.04;
    slide.addShape("ellipse", { x:nx-dotR, y:lineY-dotR+lineH/2, w:dotR*2, h:dotR*2,
      fill:{color: i===0||i===count-1 ? C.salmon : C.primary}, line:{color:C.white,pt:1} });
    const above = i%2===0;
    slide.addText(node.year||node.date||"", { x:nx-nodeSpacing*0.4, y: above?CONTENT_Y:lineY+lineH+H*0.02,
      w:nodeSpacing*0.8, h:H*0.04, fontSize:FONT.body, color:C.primary, bold:true, fontFace:FONT.face, align:"center" });
    slide.addText(node.title||node.event||"", { x:nx-nodeSpacing*0.45, y: above?CONTENT_Y+H*0.045:lineY+lineH+H*0.065,
      w:nodeSpacing*0.9, h:H*0.06, fontSize:FONT.footnote+1, color:C.text, fontFace:FONT.face, align:"center", wrap:true });
  });
  addNoteBox(slide, s.note_box);
}

// ── closing_slide ───────────────────────────────────────────────────
function renderClosingSlide(prs, slide, s) {
  slide.addShape("rect", { x:0, y:0, w:W, h:H, fill:{color:C.primary}, line:{color:C.primary} });
  slide.addShape("rect", { x:W*0.38, y:0, w:W*0.62, h:H, fill:{color:C.accent}, line:{color:C.accent} });
  slide.addShape("rect", { x:MX, y:H*0.44, w:W*0.10, h:H*0.008, fill:{color:C.salmon}, line:{color:C.salmon} });

  slide.addText(s.title||"감사합니다", { x:MX, y:H*0.24, w:W*0.55, h:H*0.20,
    fontSize:FONT.titleMain, color:C.white, bold:true, fontFace:FONT.face, wrap:true });
  if (s.message||s.summary) slide.addText(s.message||s.summary||"", { x:MX, y:H*0.50, w:W*0.55, h:H*0.18,
    fontSize:FONT.titleSub, color:C.white, fontFace:FONT.face, wrap:true, transparency:15 });
  if (s.contact) slide.addText(s.contact, { x:MX, y:H*0.72, w:W*0.55, h:H*0.12,
    fontSize:FONT.body, color:C.white, fontFace:FONT.face, transparency:25 });
}

// ═══════════════════════════════════════════════════════════════════════
// [6] DISPATCHER
// ═══════════════════════════════════════════════════════════════════════
const RENDERERS = {
  title_slide:           renderTitleSlide,
  toc_slide:             renderTocSlide,
  section_divider:       renderSectionDivider,
  content_text:          renderContentText,
  content_chart:         renderContentChart,
  table_slide:           renderTableSlide,
  wide_table:            renderWideTable,
  kpi_metrics:           renderKpiMetrics,
  two_col_text_table:    renderTwoColTextTable,
  two_col_text_chart:    renderTwoColTextChart,
  two_col_chart_text:    renderTwoColChartText,
  two_column_compare:    renderTwoColumnCompare,
  table_chart_combo:     renderTableChartCombo,
  three_column_summary:  renderThreeColumnSummary,
  composite_split:       renderCompositeSplit,
  four_quadrant:         renderFourQuadrant,
  process_flow:          renderProcessFlow,
  roadmap_timeline:      renderRoadmapTimeline,
  closing_slide:         renderClosingSlide,
};

function renderSlide(prs, slideData, pageNum) {
  const slide    = prs.addSlide();
  const layout   = (slideData.layout||"content_text").toLowerCase();
  const renderer = RENDERERS[layout];
  if (renderer) {
    renderer(prs, slide, slideData, pageNum);
  } else {
    // 알 수 없는 레이아웃 → content_text 폴백
    console.warn(`[warn] 미지원 레이아웃: ${layout} → content_text로 대체`);
    renderContentText(prs, slide, slideData, pageNum);
  }
}

// ═══════════════════════════════════════════════════════════════════════
// [7] MAIN
// ═══════════════════════════════════════════════════════════════════════
async function main() {
  const args = process.argv.slice(2);
  const get  = (flag) => { const i=args.indexOf(flag); return i>=0?args[i+1]:null; };

  const draftPath = get("--draft");
  const stylePath = get("--style");
  const outPath   = get("--out");

  if (!draftPath) { console.error("사용법: node pptxgen_builder.js --draft <path> [--style <path>] [--out <path>]"); process.exit(1); }

  const draft = JSON.parse(fs.readFileSync(draftPath, "utf-8"));
  loadStyle(stylePath);

  const prs = new PptxGenJS();
  prs.defineLayout({ name:"SRC", width:W, height:H });
  prs.layout = "SRC";
  prs.author  = draft.author  || "PPT Generator";
  prs.company = draft.company || "";
  prs.subject = draft.topic  || "";
  prs.title   = draft.topic  || draft.title || "";

  const slides = draft.slides || [];
  slides.forEach((s, i) => renderSlide(prs, s, s.slide_number ?? i+1));

  // 출력 경로 자동 생성
  const defaultOut = draftPath.replace(/draft_/, "").replace(/\.json$/, ".pptx")
                              .replace(/outputs[\\/]/, "outputs/");
  const finalOut = outPath || defaultOut;
  await prs.writeFile({ fileName: finalOut });

  // QA: 슬라이드 수 검증
  const { default: JSZip } = await import("jszip");
  const buf  = fs.readFileSync(finalOut);
  const zip  = await JSZip.loadAsync(buf);
  const actual = Object.keys(zip.files).filter(f=>/ppt\/slides\/slide\d+\.xml$/.test(f)).length;
  const expected = slides.length;
  const status = actual===expected ? "✓" : "✗ 불일치";
  console.log(`${status} 슬라이드 ${actual}/${expected}장  →  ${finalOut}`);
}

main().catch(e => { console.error("빌드 실패:", e.message); process.exit(1); });
