"""
참고 파일들을 파싱해 .cache/{filename}.json 캐시를 생성한다.
v2 — 레이아웃 존 비율, Note박스 패턴, 데이터 밀도, 헤드메시지 샘플 강화
- PDF: pdfplumber로 컬럼 구성·데이터 밀도·헤드메시지 추출
- PPTX: zipfile+xml로 EMU 존 좌표·표 구조·폰트·색상 정밀 추출
"""

import os, json, hashlib, re, zipfile, xml.etree.ElementTree as ET
from pathlib import Path
from collections import Counter, defaultdict

try:
    import pdfplumber
    HAS_PDF = True
except ImportError:
    HAS_PDF = False

BASE = Path(__file__).resolve().parents[2]

TARGETS = [
    {"path": BASE / "references/report/best_practices/[SRCIG] 1호_인프라 투자의 개요와 섹터별 소개_22.3Q.pdf",
     "folder": "best_practices", "ext": "pdf", "type": "report"},
    {"path": BASE / "references/report/best_practices/[SRCIG] 3호_23_1H_태양광발전.pdf",
     "folder": "best_practices", "ext": "pdf", "type": "report"},
    {"path": BASE / "references/report/best_practices/[SRCIG] 4호_교통인프라섹터_철도 (23.2H).pdf",
     "folder": "best_practices", "ext": "pdf", "type": "report"},
    {"path": BASE / "references/report/best_practices/[SRC 35기 유형리포트] 리테일 부동산.pdf",
     "folder": "best_practices", "ext": "pdf", "type": "report"},
    {"path": BASE / "references/report/best_practices/[34기 유형리포트] 상업용 데이터센터.pdf",
     "folder": "best_practices", "ext": "pdf", "type": "report"},
    {"path": BASE / "references/report/best_practices/36기 12월 마켓리포트-물류.pptx",
     "folder": "best_practices", "ext": "pptx", "type": "report"},
    {"path": BASE / "references/report/best_practices/오피스 자산의 호텔 컨버전.pptx",
     "folder": "best_practices", "ext": "pptx", "type": "report"},
    {"path": BASE / "references/report/narratives/[RSQUARE]_2025_서울_코리빙_하우스_마켓_리포트_파트_2_KOR.pdf",
     "folder": "narratives", "ext": "pdf", "type": "report"},
    {"path": BASE / "references/report/narratives/[SRCIG] 2호_디지털인프라 섹터의 데이터센터_22.4Q.pdf",
     "folder": "narratives", "ext": "pdf", "type": "report"},
]

TYPE_HINTS = {
    "best_practices": {
        "style_extract": True, "logic_extract": True,
        "hint": "양식+내용 모두 우수. 헤드메시지 패턴, bullet 밀도, 표/차트 배치, 섹션 논리, 좌우 분할 비율, Note박스 패턴을 최우선으로 학습하라.",
    },
    "narratives": {
        "style_extract": False, "logic_extract": True,
        "hint": "내용·논리 구조만 학습. 섹션 구성, 데이터 제시 방식, 시사점 도출 패턴, 투자 제언 구조를 추출하라. 레이아웃·색상은 무시.",
    },
    "templates": {
        "style_extract": True, "logic_extract": False,
        "hint": "디자인·레이아웃만 학습. 도형 좌표, 색상, 폰트만 추출. 내용 논리는 무시.",
    },
}


def file_hash(p: Path) -> str:
    h = hashlib.md5()
    with open(p, "rb") as f:
        for chunk in iter(lambda: f.read(8192), b""):
            h.update(chunk)
    return h.hexdigest()


# ═══════════════════════════════════════════════════════════════
# PDF 파싱 (v2)
# ═══════════════════════════════════════════════════════════════
def detect_layout_zones(page) -> dict:
    """페이지에서 텍스트 존 vs 시각화 존 비율 추정."""
    words = page.extract_words(x_tolerance=3, y_tolerance=3) or []
    if not words:
        return {"left_pct": 100, "right_pct": 0, "split_type": "single"}

    w = page.width
    left_words = [wd for wd in words if float(wd["x0"]) < w * 0.5]
    right_words = [wd for wd in words if float(wd["x0"]) >= w * 0.5]

    left_chars = sum(len(wd["text"]) for wd in left_words)
    right_chars = sum(len(wd["text"]) for wd in right_words)
    total = left_chars + right_chars or 1

    left_pct = round(left_chars / total * 100)
    right_pct = 100 - left_pct

    if left_pct > 70:
        split = "single_left"
    elif right_pct > 70:
        split = "single_right"
    elif 35 <= left_pct <= 65:
        split = "two_column"
    else:
        split = "asymmetric"

    return {"left_pct": left_pct, "right_pct": right_pct, "split_type": split}


def extract_head_messages(text: str) -> list:
    """헤드메시지 후보 추출 — 단언형 완결 문장, 수치 포함 우선."""
    lines = [l.strip() for l in text.split("\n") if l.strip()]
    candidates = []
    for l in lines:
        # 너무 짧거나 긴 것 제외
        if not (20 <= len(l) <= 90):
            continue
        # 한글 포함
        if not re.search(r'[가-힣]', l):
            continue
        # 숫자·% 포함이면 우선 후보
        has_num = bool(re.search(r'\d', l))
        # 단언형: 서술어로 끝나는 문장
        is_declarative = bool(re.search(r'[다임됨함함됨인]$|[다임]\.?$', l))
        # 목차·페이지 번호 제외
        if re.match(r'^\d+$|^(목차|contents|index)', l, re.IGNORECASE):
            continue
        score = (2 if has_num else 0) + (2 if is_declarative else 0)
        candidates.append((score, l))

    candidates.sort(key=lambda x: -x[0])
    return [c[1] for c in candidates[:10]]


def detect_note_boxes(page_text: str) -> dict:
    """Note 박스 패턴 감지."""
    note_patterns = [
        r'Note\s*[:：]', r'주\s*[:：]', r'※', r'\*\s*',
        r'출처\s*[:：]', r'Source\s*[:：]', r'\[참고\]', r'\[주\]'
    ]
    found = []
    for pat in note_patterns:
        if re.search(pat, page_text, re.IGNORECASE):
            found.append(pat.replace(r'\s*', '').replace('[:：]', '').strip())
    return {
        "has_note_box": len(found) > 0,
        "note_types": list(set(found))[:5]
    }


def extract_dense_data_points(text: str) -> list:
    """수치 데이터 포인트 추출 (단위 포함)."""
    patterns = [
        r'\d{1,3}(?:,\d{3})*(?:\.\d+)?\s*(?:GW|MW|kW|조|억|만|원|달러|USD|EUR|%|bp|배|개|명|곳|년)',
        r'(?:약|총|전년比|전년대비|YoY|CAGR)\s*\d+(?:\.\d+)?\s*%',
        r'\d{4}년\s*\d+(?:\.\d+)?',
    ]
    points = []
    for pat in patterns:
        matches = re.findall(pat, text)
        points.extend(matches[:5])
    return list(set(points))[:20]


def parse_pdf(path: Path) -> dict:
    pages_data = []
    all_text = []
    layout_counts = Counter()
    note_box_count = 0
    all_head_candidates = []
    all_data_points = []

    with pdfplumber.open(path) as pdf:
        total_pages = len(pdf.pages)

        for i, page in enumerate(pdf.pages):
            text = page.extract_text() or ""
            tables = page.extract_tables() or []

            # 레이아웃 존 분석
            zones = detect_layout_zones(page)
            layout_counts[zones["split_type"]] += 1

            # 데이터 포인트
            dp = extract_dense_data_points(text)
            all_data_points.extend(dp)

            # Note 박스
            nb = detect_note_boxes(text)
            if nb["has_note_box"]:
                note_box_count += 1

            # 헤드메시지 후보 (첫 2줄 우선)
            first_lines = [l.strip() for l in text.split("\n") if l.strip()][:3]
            for fl in first_lines:
                if 20 <= len(fl) <= 90 and re.search(r'[가-힣]', fl):
                    all_head_candidates.append(fl)

            # bullet 개수
            lines = [l.strip() for l in text.split("\n") if l.strip()]
            bullet_lines = [l for l in lines if re.match(
                r'^[▶▸•·\-\*○●◆□■▷➡→]|^\d+[\.)]|^[①②③④⑤⑥⑦⑧⑨]', l)]

            # 표 구조 상세
            table_structures = []
            for tbl in tables[:2]:
                if tbl:
                    rows = len([r for r in tbl if any(c for c in r)])
                    cols = max((len(r) for r in tbl), default=0)
                    table_structures.append({"rows": rows, "cols": cols})

            pages_data.append({
                "page": i + 1,
                "char_count": len(text),
                "line_count": len(lines),
                "bullet_count": len(bullet_lines),
                "table_count": len(tables),
                "table_structures": table_structures,
                "data_points": len(dp),
                "layout_split": zones["split_type"],
                "has_note_box": nb["has_note_box"],
            })
            all_text.append(text)

    full_text = "\n".join(all_text)

    # 헤드메시지 샘플 (품질 필터링)
    head_samples = extract_head_messages(full_text)

    # 섹션 구조
    section_headers = re.findall(
        r'^(?:[IVX]+\.|제\d+장|\d+\.)\s*.{5,40}',
        full_text, re.MULTILINE
    )[:12]

    # 통계
    avg_bullets = sum(p["bullet_count"] for p in pages_data) / max(len(pages_data), 1)
    avg_chars = sum(p["char_count"] for p in pages_data) / max(len(pages_data), 1)
    avg_data_pts = sum(p["data_points"] for p in pages_data) / max(len(pages_data), 1)
    table_pages = sum(1 for p in pages_data if p["table_count"] > 0)
    two_col_pages = sum(1 for p in pages_data if p["layout_split"] == "two_column")

    # 주요 표 구조
    all_tables = [ts for p in pages_data for ts in p["table_structures"]]
    common_table = None
    if all_tables:
        rows_list = [t["rows"] for t in all_tables]
        cols_list = [t["cols"] for t in all_tables]
        common_table = {
            "avg_rows": round(sum(rows_list) / len(rows_list), 1),
            "avg_cols": round(sum(cols_list) / len(cols_list), 1),
            "max_rows": max(rows_list),
            "max_cols": max(cols_list),
        }

    return {
        "total_pages": total_pages,
        "avg_bullets_per_page": round(avg_bullets, 1),
        "avg_chars_per_page": round(avg_chars),
        "avg_data_points_per_page": round(avg_data_pts, 1),
        "table_page_ratio": round(table_pages / max(total_pages, 1), 2),
        "two_column_page_ratio": round(two_col_pages / max(total_pages, 1), 2),
        "note_box_page_ratio": round(note_box_count / max(total_pages, 1), 2),
        "layout_distribution": dict(layout_counts.most_common()),
        "head_message_samples": head_samples,
        "section_headers": section_headers[:10],
        "sample_data_points": list(set(all_data_points))[:20],
        "common_table_structure": common_table,
        "pages_detail": pages_data[:8],
    }


# ═══════════════════════════════════════════════════════════════
# PPTX 파싱 (v2)
# ═══════════════════════════════════════════════════════════════
A_NS = "http://schemas.openxmlformats.org/drawingml/2006/main"
P_NS = "http://schemas.openxmlformats.org/presentationml/2006/main"


def get_shape_zones(root, slide_cx: int, slide_cy: int) -> dict:
    """슬라이드에서 텍스트 존과 시각화 존의 EMU 영역 분석."""
    text_zones = []
    visual_zones = []  # 표·차트·이미지

    for sp in root.iter(f"{{{P_NS}}}sp"):
        xfrm = sp.find(f".//{{{A_NS}}}xfrm")
        if xfrm is None:
            continue
        off = xfrm.find(f"{{{A_NS}}}off")
        ext = xfrm.find(f"{{{A_NS}}}ext")
        if off is None or ext is None:
            continue
        try:
            x, y = int(off.get("x", 0)), int(off.get("y", 0))
            cx, cy = int(ext.get("cx", 0)), int(ext.get("cy", 0))
        except (ValueError, TypeError):
            continue

        if cx == 0 or cy == 0:
            continue

        ph = sp.find(f".//{{{P_NS}}}ph")
        ph_type = ph.get("type", "body") if ph is not None else "body"

        if ph_type in ("title", "body", "subTitle"):
            text_zones.append({"x": x, "y": y, "cx": cx, "cy": cy, "type": ph_type})

    for tbl in root.iter(f"{{{A_NS}}}tbl"):
        sp_parent = tbl
        for parent in root.iter(f"{{{P_NS}}}sp"):
            if tbl in list(parent.iter(f"{{{A_NS}}}tbl")):
                xfrm = parent.find(f".//{{{A_NS}}}xfrm")
                if xfrm:
                    off = xfrm.find(f"{{{A_NS}}}off")
                    ext = xfrm.find(f"{{{A_NS}}}ext")
                    if off and ext:
                        visual_zones.append({
                            "x": int(off.get("x", 0)),
                            "y": int(off.get("y", 0)),
                            "cx": int(ext.get("cx", 0)),
                            "cy": int(ext.get("cy", 0)),
                            "type": "table"
                        })
                break

    # 좌우 분할 비율 계산
    mid = slide_cx // 2
    left_text = sum(1 for z in text_zones if z["x"] + z["cx"] / 2 < mid)
    right_text = sum(1 for z in text_zones if z["x"] + z["cx"] / 2 >= mid)

    split_type = "single"
    if left_text > 0 and right_text > 0:
        split_type = "two_column"
    elif visual_zones and text_zones:
        split_type = "text_plus_visual"

    return {
        "split_type": split_type,
        "text_zone_count": len(text_zones),
        "visual_zone_count": len(visual_zones),
        "text_zones_sample": text_zones[:3],
    }


def extract_table_structures(root) -> list:
    """PPTX에서 표 구조(행×열, 헤더 색상) 추출."""
    tables = []
    for tbl in root.iter(f"{{{A_NS}}}tbl"):
        rows = list(tbl.iter(f"{{{A_NS}}}tr"))
        if not rows:
            continue
        cols = list(rows[0].iter(f"{{{A_NS}}}tc")) if rows else []

        # 헤더 행 배경색
        header_color = None
        first_row = rows[0] if rows else None
        if first_row is not None:
            for srgb in first_row.iter(f"{{{A_NS}}}srgbClr"):
                val = srgb.get("val", "")
                if val:
                    header_color = f"#{val.upper()}"
                    break

        tables.append({
            "rows": len(rows),
            "cols": len(cols),
            "header_color": header_color,
        })
    return tables


def parse_pptx(path: Path) -> dict:
    slides_data = []
    fonts_found = Counter()
    colors_found = Counter()
    layout_counts = Counter()
    all_table_structures = []
    slide_cx = 12192000
    slide_cy = 6858000

    with zipfile.ZipFile(path) as z:
        names = z.namelist()

        # 슬라이드 크기 가져오기
        if "ppt/presentation.xml" in names:
            try:
                prs_xml = z.read("ppt/presentation.xml")
                prs_root = ET.fromstring(prs_xml)
                sz_el = prs_root.find(f".//{{{P_NS}}}sldSz")
                if sz_el is not None:
                    slide_cx = int(sz_el.get("cx", slide_cx))
                    slide_cy = int(sz_el.get("cy", slide_cy))
            except Exception:
                pass

        slide_files = sorted(
            [n for n in names if re.match(r'ppt/slides/slide\d+\.xml$', n)],
            key=lambda x: int(re.search(r'\d+', x.split('/')[-1]).group())
        )

        for sf in slide_files:
            raw = z.read(sf).decode("utf-8", errors="replace")
            root = ET.fromstring(raw)

            # 텍스트
            texts = [
                el.text for el in root.iter(f"{{{A_NS}}}t")
                if el.text and el.text.strip()
            ]

            # 폰트
            for rPr in root.iter(f"{{{A_NS}}}rPr"):
                latin = rPr.find(f"{{{A_NS}}}latin")
                if latin is not None:
                    tf = latin.get("typeface", "")
                    if tf and not tf.startswith("+"):
                        fonts_found[tf] += 1
                sz = rPr.get("sz")
                if sz:
                    try:
                        fonts_found[f"_sz_{int(sz)//100}pt"] += 1
                    except Exception:
                        pass

            # 색상
            for srgb in root.iter(f"{{{A_NS}}}srgbClr"):
                val = srgb.get("val", "")
                if val and val.upper() not in ("000000", "FFFFFF", "FEFEFE"):
                    colors_found[f"#{val.upper()}"] += 1

            # 존 분석
            zones = get_shape_zones(root, slide_cx, slide_cy)
            layout_counts[zones["split_type"]] += 1

            # 표 구조
            tbl_structs = extract_table_structures(root)
            all_table_structures.extend(tbl_structs)

            # bullet 수
            bullets = sum(1 for t in texts if re.match(r'^[▶▸•·\-○●]|^\d+[\.)]', t))

            slides_data.append({
                "texts_count": len(texts),
                "bullet_count": bullets,
                "table_count": len(tbl_structs),
                "split_type": zones["split_type"],
                "text_zones": zones["text_zone_count"],
                "visual_zones": zones["visual_zone_count"],
            })

        # 마스터 폰트
        for mf in [n for n in names if "slideMasters/slideMaster1.xml" in n]:
            try:
                mxml = z.read(mf).decode("utf-8", errors="replace")
                mroot = ET.fromstring(mxml)
                for latin in mroot.iter(f"{{{A_NS}}}latin"):
                    tf = latin.get("typeface", "")
                    if tf and not tf.startswith("+"):
                        fonts_found[f"master_{tf}"] += 1
            except Exception:
                pass

    # 집계
    top_fonts = [f for f, _ in fonts_found.most_common(12) if not f.startswith("_sz_")]
    font_sizes = sorted(set(
        int(f.replace("_sz_", "").replace("pt", ""))
        for f, _ in fonts_found.most_common(20) if f.startswith("_sz_")
    ))
    top_colors = [c for c, _ in colors_found.most_common(10)][:8]
    avg_bullets = sum(s["bullet_count"] for s in slides_data) / max(len(slides_data), 1)
    two_col_ratio = sum(1 for s in slides_data if s["split_type"] == "two_column") / max(len(slides_data), 1)

    common_table = None
    if all_table_structures:
        common_table = {
            "count": len(all_table_structures),
            "avg_rows": round(sum(t["rows"] for t in all_table_structures) / len(all_table_structures), 1),
            "avg_cols": round(sum(t["cols"] for t in all_table_structures) / len(all_table_structures), 1),
            "header_colors": list(set(t["header_color"] for t in all_table_structures if t["header_color"]))[:5],
        }

    return {
        "slide_dimensions_emu": {"cx": slide_cx, "cy": slide_cy},
        "slide_dimensions_pt": {"cx": round(slide_cx / 12700, 1), "cy": round(slide_cy / 12700, 1)},
        "total_slides": len(slides_data),
        "top_fonts": top_fonts,
        "font_sizes_pt": font_sizes[:8],
        "top_colors": top_colors,
        "layout_distribution": dict(layout_counts.most_common()),
        "two_column_slide_ratio": round(two_col_ratio, 2),
        "avg_bullets_per_slide": round(avg_bullets, 1),
        "table_slide_count": sum(1 for s in slides_data if s["table_count"] > 0),
        "common_table_structure": common_table,
        "slides_detail": slides_data[:8],
    }


# ═══════════════════════════════════════════════════════════════
# 캐시 생성
# ═══════════════════════════════════════════════════════════════
def build_logic_summary(path: Path) -> dict:
    """PDF에서 논리 구조 심층 추출."""
    full_text_sample = ""
    with pdfplumber.open(path) as pdf:
        # 전체 텍스트의 30% 샘플 (앞 40% + 뒤 10%)
        total = len(pdf.pages)
        sample_pages = list(range(min(int(total * 0.4) + 1, total)))
        sample_pages += list(range(max(0, total - 3), total))
        for i in set(sample_pages):
            t = pdf.pages[i].extract_text() or ""
            full_text_sample += t + "\n"

    head_samples = extract_head_messages(full_text_sample)

    # 섹션 구조
    sections = re.findall(
        r'(?:^|\n)((?:[IVX]+\.|제\d+장|\d+\.|[①②③④⑤⑥⑦])\s*.{5,40})',
        full_text_sample, re.MULTILINE
    )[:12]

    # 제언·시사점 패턴
    recommendation_patterns = []
    for line in full_text_sample.split("\n"):
        line = line.strip()
        if re.search(r'시사점|제언|권고|투자 포인트|결론|향후|전망', line) and 10 < len(line) < 100:
            recommendation_patterns.append(line)

    # 수치 기반 주장 패턴
    data_driven_claims = []
    for line in full_text_sample.split("\n"):
        line = line.strip()
        if re.search(r'\d+', line) and re.search(r'[가-힣]', line) and 20 < len(line) < 100:
            if re.search(r'%|배|조|억|GW|MW|bp', line):
                data_driven_claims.append(line)
    data_driven_claims = data_driven_claims[:10]

    narrative = "evidence_first" if any(
        kw in full_text_sample[:500] for kw in ["현황", "동향", "배경", "개요"]
    ) else "conclusion_first"

    return {
        "head_message_samples": head_samples,
        "section_structure": sections[:10],
        "recommendation_patterns": recommendation_patterns[:5],
        "data_driven_claim_samples": data_driven_claims[:8],
        "narrative_style": narrative,
    }


def build_cache(target: dict):
    path: Path = target["path"]
    folder: str = target["folder"]
    ext: str = target["ext"]

    if not path.exists():
        print(f"  [SKIP] 파일 없음: {path.name}")
        return

    cache_dir = path.parent / ".cache"
    cache_dir.mkdir(exist_ok=True)
    cache_path = cache_dir / (path.name + ".json")

    fhash = file_hash(path)
    if cache_path.exists():
        try:
            existing = json.loads(cache_path.read_text(encoding="utf-8"))
            if existing.get("file_hash") == fhash and existing.get("version") == "v2":
                print(f"  [SKIP] 캐시 최신(v2): {path.name}")
                return
        except Exception:
            pass

    print(f"  [PARSE] {path.name} ...")
    hints = TYPE_HINTS.get(folder, TYPE_HINTS["best_practices"])

    if ext == "pdf":
        if not HAS_PDF:
            print("  [ERROR] pdfplumber 없음"); return
        content = parse_pdf(path)
        logic_patterns = build_logic_summary(path) if hints["logic_extract"] else {}
    else:
        content = parse_pptx(path)
        logic_patterns = {}

    cache = {
        "version": "v2",
        "file_name": path.name,
        "file_hash": fhash,
        "file_type": ext,
        "subfolder": folder,
        "target_type": target["type"],
        "type_usage_hints": {
            "style_extract": hints["style_extract"],
            "logic_extract": hints["logic_extract"],
            "guidance": hints["hint"],
        },
        "content": content,
        "logic_patterns": logic_patterns,
    }

    cache_path.write_text(json.dumps(cache, ensure_ascii=False, indent=2), encoding="utf-8")
    size = cache_path.stat().st_size
    print(f"  [DONE] {size:,}B → {cache_path.relative_to(BASE)}")


def main():
    print(f"\n=== ref-distiller v2: {len(TARGETS)}개 파일 ===\n")
    for t in TARGETS:
        build_cache(t)
    print("\n=== 완료 ===")


if __name__ == "__main__":
    main()
