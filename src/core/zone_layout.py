"""
zone_layout.py — 존 기반 레이아웃 시스템

슬라이드 콘텐츠 영역을 최대 2×2 그리드로 분할하고,
참고 draft 파일 기반 캘리브레이션으로 콘텐츠 분량 적합성을 판단한다.

존 구성 (zone_config):
  "full"    — 1개 블록 (전체)
  "L|R"     — 좌(53%) / 우(47%)
  "T/B"     — 상(50%) / 하(50%)
  "L1|R2"   — 좌 넓음(58%) / 우 2단
  "L2|R1"   — 좌 2단 / 우 넓음(58%)
  "2x2"     — 2×2 격자 (4블록)

draft.json 슬라이드에 zone_config 사용 예시:
{
  "layout": "zone",
  "zone_config": "L|R",
  "title": "섹션명",
  "head_message": "...",
  "source": "출처",
  "zones": [
    {"id": "L", "component": "table", "title": "표 제목", "table": {"headers": [...], "rows": [...]}},
    {"id": "R", "component": "chart", "chart": {"chart_type": "bar", "series": [...]}}
  ]
}
"""

import glob
import json
import os

# ---------------------------------------------------------------------------
# 슬라이드 크기 (template_test1.pptx 기준)
# PPTX 파일은 그대로 유지하고, 파이썬 로직이 이 크기에 맞춰 동작한다.
# ---------------------------------------------------------------------------
SLIDE_W_EMU = 17_610_138   # 1386.63 pt
SLIDE_H_EMU =  9_906_000   # 780 pt

# ---------------------------------------------------------------------------
# 슬라이드 콘텐츠 영역 상수 (EMU)
# template_test1 실측값 기준:
#   헤더 breadcrumb: y=41pt, h=74pt → 하단 y=115pt
#   head_message:    y=141pt, h≈29pt
#   bullet bar:      y=174pt, h=30pt → 하단 y=204pt
#   콘텐츠 시작:     y≈210pt
#   슬라이드 번호:   y=731pt → 콘텐츠 하단 ≈ y=720pt
#   좌측 여백:       x=67pt
#   우측 끝:         x≈1317pt (폭 1250pt)
# ---------------------------------------------------------------------------
CONTENT_X = int(67  * 12700)    #  850_900 EMU
CONTENT_Y = int(210 * 12700)    # 2_667_000 EMU
CONTENT_W = int(1250 * 12700)   # 15_875_000 EMU
CONTENT_H = int(510 * 12700)    #  6_477_000 EMU  (210→720pt)
ZONE_GAP  = int(18  * 12700)    #  228_600 EMU

# 출처(source) 텍스트 y 좌표 — 슬라이드 번호 바로 위
SOURCE_Y  = int(722 * 12700)    # 9_169_400 EMU
SOURCE_CY = int(24  * 12700)    #   304_800 EMU


# ---------------------------------------------------------------------------
# 존 좌표 계산
# ---------------------------------------------------------------------------

def _build_zone_configs() -> dict:
    cx, cy = CONTENT_X, CONTENT_Y
    cw, ch = CONTENT_W, CONTENT_H
    g = ZONE_GAP

    # L|R
    lw   = int(cw * 0.53)
    rw   = cw - lw - g
    rx   = cx + lw + g

    # T/B
    th   = int(ch * 0.50)
    bh   = ch - th - g
    by   = cy + th + g

    # L1|R2
    l1w  = int(cw * 0.58)
    r2w  = cw - l1w - g
    r2x  = cx + l1w + g
    r2th = int(ch * 0.50)
    r2bh = ch - r2th - g
    r2by = cy + r2th + g

    # L2|R1
    r1w  = int(cw * 0.58)
    l2w  = cw - r1w - g
    r1x  = cx + l2w + g
    l2th = int(ch * 0.50)
    l2bh = ch - l2th - g
    l2by = cy + l2th + g

    # 2×2
    qlw  = int(cw * 0.50) - g // 2
    qrw  = cw - qlw - g
    qrx  = cx + qlw + g
    qth  = int(ch * 0.50)
    qbh  = ch - qth - g
    qby  = cy + qth + g

    return {
        "full": [
            {"id": "A",   "x": cx,  "y": cy,   "w": cw,  "h": ch,   "size": "full"},
        ],
        "L|R": [
            {"id": "L",   "x": cx,  "y": cy,   "w": lw,  "h": ch,   "size": "half_w"},
            {"id": "R",   "x": rx,  "y": cy,   "w": rw,  "h": ch,   "size": "half_w"},
        ],
        "T/B": [
            {"id": "T",   "x": cx,  "y": cy,   "w": cw,  "h": th,   "size": "half_h"},
            {"id": "B",   "x": cx,  "y": by,   "w": cw,  "h": bh,   "size": "half_h"},
        ],
        "L1|R2": [
            {"id": "L",   "x": cx,  "y": cy,   "w": l1w, "h": ch,   "size": "half_w"},
            {"id": "R_T", "x": r2x, "y": cy,   "w": r2w, "h": r2th, "size": "quarter"},
            {"id": "R_B", "x": r2x, "y": r2by, "w": r2w, "h": r2bh, "size": "quarter"},
        ],
        "L2|R1": [
            {"id": "L_T", "x": cx,  "y": cy,   "w": l2w, "h": l2th, "size": "quarter"},
            {"id": "L_B", "x": cx,  "y": l2by, "w": l2w, "h": l2bh, "size": "quarter"},
            {"id": "R",   "x": r1x, "y": cy,   "w": r1w, "h": ch,   "size": "half_w"},
        ],
        "2x2": [
            {"id": "TL",  "x": cx,  "y": cy,   "w": qlw, "h": qth,  "size": "quarter"},
            {"id": "TR",  "x": qrx, "y": cy,   "w": qrw, "h": qth,  "size": "quarter"},
            {"id": "BL",  "x": cx,  "y": qby,  "w": qlw, "h": qbh,  "size": "quarter"},
            {"id": "BR",  "x": qrx, "y": qby,  "w": qrw, "h": qbh,  "size": "quarter"},
        ],
    }


ZONE_CONFIGS = _build_zone_configs()


# ---------------------------------------------------------------------------
# 기본 캘리브레이션
# 참고 draft 파일에서 실제 분량을 측정해 덮어씀.
# "full" 기준: 참고 자료의 75 퍼센타일 값
# "half_*" / "quarter" 는 full 대비 비율로 산출
# ---------------------------------------------------------------------------

DEFAULT_CALIBRATION: dict = {
    "full": {
        "table_max_rows":   10,
        "bullet_max_items": 7,
        "text_max_chars":   500,
    },
    "half_w": {
        "table_max_rows":   6,
        "bullet_max_items": 5,
        "text_max_chars":   300,
    },
    "half_h": {
        "table_max_rows":   5,
        "bullet_max_items": 4,
        "text_max_chars":   250,
    },
    "quarter": {
        "table_max_rows":   3,
        "bullet_max_items": 3,
        "text_max_chars":   150,
    },
}

MIN_FONT_PT: dict = {
    "body_text":    11,   # 본문 최소 폰트
    "table_cell":   10,   # 표 셀 최소 폰트
    "chart_label":   9,   # 차트 레이블 최소 폰트
    "head_message": 16,   # 헤드메시지 고정 (절대 축소 불가)
}


def load_calibration(draft_dir: str = "outputs") -> dict:
    """
    outputs/draft_*.json 을 스캔해 실제 콘텐츠 분량을 측정하고
    DEFAULT_CALIBRATION 을 보정한 캘리브레이션 딕셔너리를 반환한다.
    참고 파일이 없으면 DEFAULT_CALIBRATION 그대로 반환.
    """
    cal = {k: dict(v) for k, v in DEFAULT_CALIBRATION.items()}

    draft_files = glob.glob(os.path.join(draft_dir, "draft_*.json"))
    if not draft_files:
        return cal

    table_rows:   list[int] = []
    bullet_items: list[int] = []
    text_chars:   list[int] = []

    for path in draft_files:
        try:
            with open(path, encoding="utf-8") as f:
                draft = json.load(f)
        except Exception:
            continue

        for slide in draft.get("slides", []):
            tbl = slide.get("table", {})
            if tbl and tbl.get("rows"):
                table_rows.append(len(tbl["rows"]))

            body = slide.get("body", [])
            if isinstance(body, list) and body:
                bullet_items.append(len(body))
                for item in body:
                    text_chars.append(len(str(item)))

    def _p75(lst: list) -> int:
        if not lst:
            return 0
        return int(sorted(lst)[int(len(lst) * 0.75)])

    measured_rows   = _p75(table_rows)
    measured_items  = _p75(bullet_items)
    measured_chars  = _p75(text_chars) * max(1, int(_p75(bullet_items)))

    if measured_rows > 0:
        cal["full"]["table_max_rows"]   = max(DEFAULT_CALIBRATION["full"]["table_max_rows"],   measured_rows)
    if measured_items > 0:
        cal["full"]["bullet_max_items"] = max(DEFAULT_CALIBRATION["full"]["bullet_max_items"], measured_items)
    if measured_chars > 0:
        cal["full"]["text_max_chars"]   = max(DEFAULT_CALIBRATION["full"]["text_max_chars"],   measured_chars)

    # 절반/쿼터 크기는 full 기준 비율 적용
    for size, ratio_rows, ratio_items, ratio_chars in [
        ("half_w",  0.58, 0.65, 0.60),
        ("half_h",  0.52, 0.57, 0.52),
        ("quarter", 0.35, 0.40, 0.30),
    ]:
        cal[size]["table_max_rows"]   = max(2, int(cal["full"]["table_max_rows"]   * ratio_rows))
        cal[size]["bullet_max_items"] = max(2, int(cal["full"]["bullet_max_items"] * ratio_items))
        cal[size]["text_max_chars"]   = max(80, int(cal["full"]["text_max_chars"]  * ratio_chars))

    return cal


# ---------------------------------------------------------------------------
# 분량 적합성 판단
# ---------------------------------------------------------------------------

def check_content_fits(component: str, data: dict, zone_size: str,
                        cal: dict | None = None) -> bool:
    """
    컴포넌트 데이터가 지정된 존 크기에 들어가는지 확인.
    True → 적합 / False → 분할 필요
    """
    if cal is None:
        cal = DEFAULT_CALIBRATION
    limits = cal.get(zone_size, cal["full"])

    if component == "table":
        rows = data.get("rows", [])
        return len(rows) <= limits["table_max_rows"]

    if component in ("bullet", "text"):
        items = data.get("body", data.get("items", []))
        if isinstance(items, list):
            return len(items) <= limits["bullet_max_items"]
        return len(str(items)) <= limits["text_max_chars"]

    # chart, diagram 은 존 크기에 맞게 스케일 조정하므로 항상 True
    return True


def should_split_slide(zones: list, zone_config: str,
                        cal: dict | None = None) -> bool:
    """
    zones 리스트의 각 컴포넌트가 해당 존 크기에 맞는지 확인.
    하나라도 초과이면 True (분할 권장).
    """
    rects = ZONE_CONFIGS.get(zone_config, ZONE_CONFIGS["full"])
    rect_by_id = {r["id"]: r for r in rects}

    for zone in zones:
        zid       = zone.get("id", "A")
        component = zone.get("component", "bullet")
        rect      = rect_by_id.get(zid, rects[0])
        size      = rect["size"]

        data = {}
        if component == "table":
            data = zone.get("table", {})
        elif component in ("bullet", "text"):
            data = {"body": zone.get("body", zone.get("items", []))}

        if not check_content_fits(component, data, size, cal):
            return True

    return False


# ---------------------------------------------------------------------------
# 존 좌표 조회
# ---------------------------------------------------------------------------

def get_zone_rects(zone_config: str) -> list[dict]:
    """
    zone_config 이름으로 존 좌표 리스트 반환.
    각 항목: {"id": str, "x": int, "y": int, "w": int, "h": int, "size": str}
    단위: EMU
    """
    return ZONE_CONFIGS.get(zone_config, ZONE_CONFIGS["full"])


# ---------------------------------------------------------------------------
# zone_config 자동 추천 (content-planner / logic-analyst 보조용)
# ---------------------------------------------------------------------------

def suggest_zone_config(zones: list[dict]) -> str:
    """
    zones 리스트의 컴포넌트 수·타입으로 적합한 zone_config 문자열 제안.
    zones: [{"component": "table"|"chart"|"bullet"|"text"|"diagram", ...}]
    """
    n = len(zones)
    if n <= 1:
        return "full"

    types = [z.get("component", "text") for z in zones]

    if n == 2:
        # 표 + 차트, 불릿 + 차트 → 좌우 분할
        if "chart" in types:
            return "L|R"
        # 텍스트끼리 비교 → 좌우
        if set(types) <= {"bullet", "text"}:
            t0 = zones[0].get("body", zones[0].get("items", []))
            # 분량이 적으면 T/B, 많으면 L|R
            if isinstance(t0, list) and len(t0) <= 4:
                return "T/B"
            return "L|R"
        return "L|R"

    if n == 3:
        # 왼쪽에 큰 컴포넌트(표/차트), 오른쪽 2단 → L1|R2
        if types[0] in ("table", "chart") and types[1:] != ["chart", "chart"]:
            return "L1|R2"
        # 오른쪽에 큰 컴포넌트 → L2|R1
        if types[-1] in ("table", "chart"):
            return "L2|R1"
        return "L1|R2"

    if n == 4:
        return "2x2"

    # 5개 이상이면 분할 권장 (여기서는 full 반환, should_split_slide가 처리)
    return "full"
