/**
 * pptxgenjs_pilot.js
 *
 * 파일럿: draft JSON의 슬라이드 5(composite_split)와 8(wide_table)을
 * pptxgenjs로 생성해 repair 없는 PPTX 출력 여부 검증
 *
 * 실행: node src/pilot/pptxgenjs_pilot.js
 */

const PptxGenJS = require("pptxgenjs");
const fs = require("fs");
const path = require("path");

// ─── 색상/폰트 (style_report.json 기반) ─────────────────────────────
const STYLE = {
  primary:    "627365",
  accent:     "2D3734",
  salmon:     "D98F76",
  warm:       "A09567",
  blue:       "406D92",
  text:       "232323",
  white:      "FFFFFF",
  neutral:    "F2F2F2",
  font:       "Pretendard",
};

// ─── 슬라이드 치수 (LAYOUT_WIDE 실제 크기: 13.33×7.5 inch) ───────────
const W = 13.33; // inch  ← LAYOUT_WIDE 실제 너비
const H = 7.5;   // inch  ← LAYOUT_WIDE 실제 높이

// 헤더바 높이
const HEADER_H   = 0.56;
// head_message 배너 높이
const BANNER_H   = 0.40;
// footer 높이
const FOOTER_H   = 0.28;
// 콘텐츠 영역 상단 Y
const CONTENT_Y  = HEADER_H + BANNER_H + 0.10;
// 콘텐츠 영역 높이
const CONTENT_H  = H - CONTENT_Y - FOOTER_H - 0.06;
// 좌우 여백
const MARGIN_X   = 0.28;

// ─── 공통 헬퍼 ────────────────────────────────────────────────────────

/**
 * 슬라이드 공통 헤더바 + head_message 배너 + 페이지번호 추가
 */
function addHeader(slide, slideData, pageNum) {
  // 헤더바 배경
  slide.addShape("rect", {
    x: 0, y: 0, w: W, h: HEADER_H,
    fill: { color: STYLE.primary },
    line: { color: STYLE.primary },
  });

  // section_label (섹션 번호)
  const sectionLabel = slideData.section_number
    ? `Section ${slideData.section_number}`
    : "";

  // 두 줄 헤더: section_label + slide title
  if (sectionLabel) {
    slide.addText(sectionLabel, {
      x: MARGIN_X, y: 0.03, w: W - MARGIN_X * 2, h: 0.15,
      fontSize: 7, color: "FFFFFF", bold: false,
      fontFace: STYLE.font, transparency: 35,
      valign: "top",
    });
  }
  slide.addText(slideData.title || "", {
    x: MARGIN_X, y: sectionLabel ? 0.17 : 0.07,
    w: W - 1.2, h: 0.24,
    fontSize: 13, color: "FFFFFF", bold: true,
    fontFace: STYLE.font, valign: "middle",
  });

  // 페이지 번호
  slide.addText(String(pageNum), {
    x: W - 0.55, y: 0.07, w: 0.40, h: 0.28,
    fontSize: 9, color: "FFFFFF", bold: false,
    fontFace: STYLE.font, align: "right", valign: "middle",
    transparency: 35,
  });

  // head_message 배너
  if (slideData.head_message) {
    slide.addShape("rect", {
      x: 0, y: HEADER_H, w: W, h: BANNER_H,
      fill: { color: "FFFFFF" },
      line: { color: "FFFFFF" },
    });
    // 왼쪽 강조 선
    slide.addShape("rect", {
      x: MARGIN_X, y: HEADER_H + 0.04,
      w: 0.03, h: BANNER_H - 0.08,
      fill: { color: STYLE.salmon },
      line: { color: STYLE.salmon },
    });
    slide.addText(slideData.head_message, {
      x: MARGIN_X + 0.08, y: HEADER_H + 0.02,
      w: W - MARGIN_X * 2 - 0.08, h: BANNER_H - 0.04,
      fontSize: 9, color: STYLE.salmon, bold: false,
      fontFace: STYLE.font, valign: "middle",
      wrap: true,
    });
  }
}

/**
 * note_box 추가 (슬라이드 하단)
 */
function addNoteBox(slide, noteBox) {
  if (!noteBox) return;
  const noteY = H - FOOTER_H;
  slide.addText(noteBox.content || "", {
    x: MARGIN_X, y: noteY + 0.02, w: W - MARGIN_X * 2, h: FOOTER_H - 0.04,
    fontSize: 6, color: "808080", fontFace: STYLE.font,
    italic: noteBox.type === "source",
  });
}

// ─── 슬라이드 5: composite_split ──────────────────────────────────────
function addSlide5(prs, slideData) {
  const slide = prs.addSlide();
  addHeader(slide, slideData, 5);

  // 레이아웃: 좌 55% 차트, 우 45% (상 bullets + 하 mini-table)
  const leftW  = (W - MARGIN_X * 2) * 0.54;
  const rightW = (W - MARGIN_X * 2) * 0.44;
  const leftX  = MARGIN_X;
  const rightX = MARGIN_X + leftW + 0.12;
  const splitH_top = CONTENT_H * 0.52;
  const splitH_bot = CONTENT_H * 0.45;
  const splitGap   = CONTENT_H * 0.03;

  // ── 좌: Bar 차트 ──
  const chartData = slideData.main_zone.chart;
  const barLabels = chartData.data.map(d => d.label);
  const barValues = chartData.data.map(d => d.value);

  slide.addChart(prs.ChartType.bar, [
    {
      name: "신규 설치 (GW)",
      labels: barLabels,
      values: barValues,
    }
  ], {
    x: leftX, y: CONTENT_Y, w: leftW, h: CONTENT_H - 0.05,
    chartColors: [STYLE.primary],
    showLegend: false,
    showTitle: true,
    title: chartData.title || "",
    titleFontSize: 8,
    titleColor: STYLE.accent,
    titleFontFace: STYLE.font,
    dataLabelFontSize: 7,
    catAxisLabelFontSize: 8,
    valAxisLabelFontSize: 8,
    catAxisLabelColor: STYLE.text,
    valAxisLabelColor: STYLE.text,
    dataLabelColor: STYLE.text,
    valGridLine: { style: "none" },
    catGridLine: { style: "none" },
  });

  // ── 우 상: Bullets ──
  const bullets = slideData.sub_zone_top.bullets || [];
  const bulletRows = bullets.map((b, i) => [
    { text: `${i + 1}. ${b}`, options: { fontSize: 8, color: STYLE.text, fontFace: STYLE.font, paraSpaceAfter: 3 } }
  ]);
  slide.addText(
    bullets.map((b, i) => ({
      text: `${i + 1}. ${b}\n`,
      options: { fontSize: 8, color: STYLE.text, fontFace: STYLE.font, paraSpaceAfter: 4, bullet: false }
    })),
    {
      x: rightX, y: CONTENT_Y, w: rightW, h: splitH_top,
      valign: "top", wrap: true,
    }
  );

  // ── 우 하: Mini-table ──
  const tbl = slideData.sub_zone_bottom.table;
  const tableY = CONTENT_Y + splitH_top + splitGap;

  // 헤더 행
  const headerRow = tbl.headers.map(h => ({
    text: h,
    options: {
      bold: true, fontSize: 8, color: STYLE.white,
      fontFace: STYLE.font, align: "center",
      fill: { color: STYLE.primary },
    }
  }));

  // 데이터 행
  const dataRows = tbl.rows.map((row, ri) => row.map((cell, ci) => ({
    text: cell,
    options: {
      fontSize: 8, color: STYLE.text, fontFace: STYLE.font,
      align: ci === 0 ? "left" : "center",
      fill: { color: ri % 2 === 0 ? STYLE.neutral : STYLE.white },
    }
  })));

  slide.addTable([headerRow, ...dataRows], {
    x: rightX, y: tableY, w: rightW,
    rowH: 0.22,
    border: { type: "solid", color: "DDDDDD", pt: 0.5 },
    autoPage: false,
  });

  addNoteBox(slide, slideData.note_box);
}

// ─── 슬라이드 8: wide_table ───────────────────────────────────────────
function addSlide8(prs, slideData) {
  const slide = prs.addSlide();
  addHeader(slide, slideData, 8);

  const tbl = slideData.table;
  const colCount = tbl.headers.length;

  // 첫 열 넓게, 나머지 균등
  const firstColW = 1.6;
  const restColW  = (W - MARGIN_X * 2 - firstColW) / (colCount - 1);
  const colWidths = [firstColW, ...Array(colCount - 1).fill(restColW)];

  // 헤더
  const headerRow = tbl.headers.map((h, i) => ({
    text: h,
    options: {
      bold: true, fontSize: 8, color: STYLE.white,
      fontFace: STYLE.font, align: "center",
      fill: { color: STYLE.accent },
    }
  }));
  // 첫 열 헤더는 좌정렬
  headerRow[0].options.align = "left";

  // 데이터 행
  const dataRows = tbl.rows.map((row, ri) => row.map((cell, ci) => ({
    text: cell,
    options: {
      fontSize: 7.5, color: STYLE.text, fontFace: STYLE.font,
      align: ci === 0 ? "left" : "center",
      fill: { color: ri % 2 === 0 ? STYLE.neutral : STYLE.white },
      valign: "middle",
    }
  })));

  // 테이블 높이: 콘텐츠 영역의 60% (key_points 공간 확보)
  const tableH = CONTENT_H * 0.60;

  slide.addTable([headerRow, ...dataRows], {
    x: MARGIN_X, y: CONTENT_Y, w: W - MARGIN_X * 2,
    rowH: tableH / (tbl.rows.length + 1),
    colW: colWidths,
    border: { type: "solid", color: "DDDDDD", pt: 0.5 },
    autoPage: false,
  });

  // key_points
  const keyY = CONTENT_Y + tableH + 0.10;
  const keyH  = H - keyY - FOOTER_H - 0.05;

  const keyPoints = (slideData.key_points || []).map((kp, i) => ({
    text: `• ${kp}\n`,
    options: {
      fontSize: 8, color: STYLE.text, fontFace: STYLE.font,
      paraSpaceAfter: 3, bullet: false
    }
  }));

  if (keyPoints.length > 0) {
    // 구분선
    slide.addShape("line", {
      x: MARGIN_X, y: keyY - 0.05, w: W - MARGIN_X * 2, h: 0,
      line: { color: "DDDDDD", width: 0.5 },
    });
    slide.addText(keyPoints, {
      x: MARGIN_X, y: keyY, w: W - MARGIN_X * 2, h: keyH,
      valign: "top", wrap: true,
    });
  }

  addNoteBox(slide, slideData.note_box);
}

// ─── 메인 실행 ────────────────────────────────────────────────────────
async function main() {
  const draftPath = path.join(__dirname, "../../outputs/draft_report_solar_power_invest_v2.json");
  const draft = JSON.parse(fs.readFileSync(draftPath, "utf-8"));

  const slide5 = draft.slides.find(s => s.slide_number === 5);
  const slide8 = draft.slides.find(s => s.slide_number === 8);

  const prs = new PptxGenJS();
  prs.layout = "LAYOUT_WIDE"; // 13.33 × 7.5 — pptxgenjs 기본 와이드

  // 슬라이드 추가
  addSlide5(prs, slide5);
  addSlide8(prs, slide8);

  const outPath = path.join(__dirname, "../../outputs/pilot_pptxgenjs_slides_5_8.pptx");
  await prs.writeFile({ fileName: outPath });
  console.log("✓ 생성 완료:", outPath);
}

main().catch(err => {
  console.error("오류:", err.message);
  process.exit(1);
});
