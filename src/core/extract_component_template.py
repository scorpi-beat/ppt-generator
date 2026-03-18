"""
extract_component_template.py — Plan A PPTX에서 레이아웃별 대표 슬라이드 추출

draft.json + Plan A PPTX → component_template.pptx + index.json

사용법:
  python src/core/extract_component_template.py \
    --draft outputs/draft_report_offshore_wind.json \
    --source outputs/report_offshore_wind_A.pptx \
    --output outputs/component_template_report.pptx
"""

import argparse
import json
import os
import re
import zipfile
from xml.etree import ElementTree as ET


# ---------------------------------------------------------------------------
# 슬라이드 복사 (zipfile 방식)
# ---------------------------------------------------------------------------

NS_REL = "http://schemas.openxmlformats.org/package/2006/relationships"


def _normalize_path(base_dir: str, target: str) -> str:
    """상대 경로(../ 포함)를 ppt/ 기준 절대 경로로 정규화."""
    if target.startswith("http"):
        return target
    # 슬라이드 기준 경로 병합
    if target.startswith("../"):
        combined = base_dir.rstrip("/") + "/" + target
    elif not target.startswith("ppt/"):
        combined = base_dir.rstrip("/") + "/" + target
    else:
        combined = target

    # path normalize (.. 처리)
    parts = combined.split("/")
    normalized = []
    for p in parts:
        if p == "..":
            if normalized:
                normalized.pop()
        elif p and p != ".":
            normalized.append(p)
    return "/".join(normalized)


def copy_slide(src_zf: zipfile.ZipFile, dest_files: dict,
               src_slide_num: int, dest_slide_num: int):
    """
    src_zf에서 슬라이드 src_slide_num을 dest_files에 dest_slide_num으로 복사.

    Args:
        src_zf: source PPTX의 열린 ZipFile
        dest_files: {arcname: bytes} — 최종 ZIP에 들어갈 파일 모음
        src_slide_num: source 슬라이드 번호 (1-based)
        dest_slide_num: destination 슬라이드 번호 (1-based)
    """
    slide_path = f"ppt/slides/slide{src_slide_num}.xml"
    rels_path = f"ppt/slides/_rels/slide{src_slide_num}.xml.rels"
    slide_base_dir = "ppt/slides"

    if slide_path not in src_zf.namelist():
        raise FileNotFoundError(f"슬라이드를 찾을 수 없음: {slide_path}")

    slide_xml = src_zf.read(slide_path)

    # rels 파싱
    if rels_path in src_zf.namelist():
        rels_xml = src_zf.read(rels_path)
    else:
        rels_xml = b'<?xml version="1.0" encoding="UTF-8" standalone="yes"?><Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"/>'

    rels_root = ET.fromstring(rels_xml)
    all_namelist = set(src_zf.namelist())

    for rel in rels_root.findall(f"{{{NS_REL}}}Relationship"):
        target = rel.get("Target", "")
        if not target or target.startswith("http"):
            continue

        abs_target = _normalize_path(slide_base_dir, target)

        if abs_target not in all_namelist:
            continue

        basename = abs_target.split("/")[-1]
        parent_dir = abs_target.rsplit("/", 1)[0]
        new_basename = f"s{dest_slide_num}_{basename}"
        new_abs = parent_dir + "/" + new_basename

        # 파일 복사
        dest_files[new_abs] = src_zf.read(abs_target)

        # 차트 파일이면 해당 차트의 rels도 복사
        if "charts/chart" in abs_target and abs_target.endswith(".xml"):
            chart_rels_src = parent_dir + "/_rels/" + basename + ".rels"
            if chart_rels_src in all_namelist:
                chart_rels_dst = parent_dir + "/_rels/" + new_basename + ".rels"
                # 차트 rels 내용도 복사 (embedding 참조 등)
                chart_rels_data = src_zf.read(chart_rels_src)
                dest_files[chart_rels_dst] = chart_rels_data

                # 차트 rels에서 embedding(xlsx 등) 참조 복사
                try:
                    chart_rels_root = ET.fromstring(chart_rels_data)
                    for crel in chart_rels_root.findall(f"{{{NS_REL}}}Relationship"):
                        ctarget = crel.get("Target", "")
                        if not ctarget or ctarget.startswith("http"):
                            continue
                        c_abs = _normalize_path(parent_dir, ctarget)
                        if c_abs in all_namelist and c_abs not in dest_files:
                            dest_files[c_abs] = src_zf.read(c_abs)
                except ET.ParseError:
                    pass

        # Target을 새 파일명으로 업데이트
        new_target = target.replace(basename, new_basename)
        rel.set("Target", new_target)

    # 새 rels XML 생성
    new_rels_xml = ET.tostring(
        rels_root,
        xml_declaration=True,
        encoding="UTF-8",
        short_empty_elements=False,
    )

    dest_files[f"ppt/slides/slide{dest_slide_num}.xml"] = slide_xml
    dest_files[f"ppt/slides/_rels/slide{dest_slide_num}.xml.rels"] = new_rels_xml


# ---------------------------------------------------------------------------
# Content_Types 업데이트
# ---------------------------------------------------------------------------

def _update_content_types(ct_xml: bytes, slide_count: int, extra_files: dict) -> bytes:
    """[Content_Types].xml에 슬라이드 및 새로 추가된 차트 파일 등록."""
    CT_NS = "http://schemas.openxmlformats.org/package/2006/content-types"
    SLIDE_CT = "application/vnd.openxmlformats-officedocument.presentationml.slide+xml"
    CHART_CT = "application/vnd.openxmlformats-officedocument.drawingml.chart+xml"

    root = ET.fromstring(ct_xml)

    # 기존 슬라이드 Override 제거
    existing_slides = [
        el for el in root
        if el.get("ContentType") == SLIDE_CT
    ]
    for el in existing_slides:
        root.remove(el)

    # 슬라이드 재등록
    for i in range(1, slide_count + 1):
        override = ET.SubElement(root, f"{{{CT_NS}}}Override")
        override.set("PartName", f"/ppt/slides/slide{i}.xml")
        override.set("ContentType", SLIDE_CT)

    # 새로 복사된 차트 파일 등록 (s{N}_chart*.xml)
    existing_parts = {el.get("PartName") for el in root}
    for arcname in extra_files:
        if re.match(r"ppt/charts/s\d+_chart\d+\.xml$", arcname):
            part_name = f"/{arcname}"
            if part_name not in existing_parts:
                override = ET.SubElement(root, f"{{{CT_NS}}}Override")
                override.set("PartName", part_name)
                override.set("ContentType", CHART_CT)
                existing_parts.add(part_name)

    return ET.tostring(root, xml_declaration=True, encoding="UTF-8", short_empty_elements=False)


# ---------------------------------------------------------------------------
# 메인 추출 로직
# ---------------------------------------------------------------------------

def extract_component_template(draft_path: str, source_path: str, output_path: str):
    """
    draft.json에서 레이아웃 타입별 첫 번째 슬라이드를 찾아
    source PPTX에서 해당 슬라이드를 추출해 component_template.pptx 생성.
    """

    # 1. draft.json 로드
    with open(draft_path, encoding="utf-8") as f:
        draft = json.load(f)

    slides = draft.get("slides", [])

    # 2. 레이아웃 타입별 첫 번째 슬라이드 번호 수집
    layout_to_slide: dict = {}
    for s in slides:
        layout = s.get("layout", "")
        num = s.get("slide_number")
        if layout and num is not None and layout not in layout_to_slide:
            layout_to_slide[layout] = int(num)

    if not layout_to_slide:
        raise ValueError("draft.json에서 슬라이드 레이아웃 정보를 찾을 수 없습니다.")

    print(f"  발견된 레이아웃 타입 ({len(layout_to_slide)}개):")
    for layout, num in layout_to_slide.items():
        print(f"    {layout}: slide {num}")

    # 3. source PPTX 처리
    with zipfile.ZipFile(source_path, "r") as src_zf:
        all_names = set(src_zf.namelist())

        # 공통 파일들 수집 (슬라이드 제외)
        shared_files: dict = {}
        slide_pattern = re.compile(r"ppt/slides/slide\d+\.xml$")
        slide_rels_pattern = re.compile(r"ppt/slides/_rels/slide\d+\.xml\.rels$")

        for name in src_zf.namelist():
            if slide_pattern.match(name) or slide_rels_pattern.match(name):
                continue
            shared_files[name] = src_zf.read(name)

        # 4. 레이아웃 타입별 슬라이드 복사
        dest_files: dict = {}  # 새로 생성되는 슬라이드 관련 파일들
        layout_index: dict = {}  # {layout_name: dest_slide_num}

        dest_slide_num = 0
        for layout, src_num in sorted(layout_to_slide.items(), key=lambda x: x[1]):
            dest_slide_num += 1
            try:
                copy_slide(src_zf, dest_files, src_num, dest_slide_num)
                layout_index[layout] = dest_slide_num
                print(f"  복사: {layout} (source slide {src_num} → dest slide {dest_slide_num})")
            except FileNotFoundError as e:
                print(f"  [경고] {e} — 건너뜀")
                dest_slide_num -= 1

    if dest_slide_num == 0:
        raise RuntimeError("복사된 슬라이드가 없습니다.")

    # 5. [Content_Types].xml 업데이트
    ct_path = "[Content_Types].xml"
    if ct_path in shared_files:
        shared_files[ct_path] = _update_content_types(
            shared_files[ct_path], dest_slide_num, dest_files
        )

    # 6. presentation.xml 업데이트
    _update_presentation_xml_bytes(shared_files, dest_slide_num)

    # 7. ppt/_rels/presentation.xml.rels 업데이트
    _update_prs_rels_bytes(shared_files, dest_slide_num)

    # 8. ZIP 패키징
    os.makedirs(os.path.dirname(os.path.abspath(output_path)), exist_ok=True)
    tmp_path = output_path + ".tmp.zip"
    with zipfile.ZipFile(tmp_path, "w", zipfile.ZIP_DEFLATED) as out_zf:
        # 공통 파일 먼저
        for arcname, data in shared_files.items():
            out_zf.writestr(arcname, data)
        # 슬라이드 및 관련 파일
        for arcname, data in dest_files.items():
            out_zf.writestr(arcname, data)

    if os.path.exists(output_path):
        os.remove(output_path)
    os.rename(tmp_path, output_path)

    # 9. index.json 저장
    index_path = output_path.replace(".pptx", "_index.json")
    index_data = {
        "template": output_path,
        "source": source_path,
        "layouts": layout_index,
    }
    with open(index_path, "w", encoding="utf-8") as f:
        json.dump(index_data, f, ensure_ascii=False, indent=2)

    print(f"\n  component_template: {output_path}")
    print(f"  index:              {index_path}")
    print(f"  레이아웃 매핑: {layout_index}")

    return output_path, index_path


def _update_presentation_xml_bytes(shared_files: dict, slide_count: int):
    """shared_files 내 presentation.xml의 sldIdLst 업데이트."""
    prs_key = "ppt/presentation.xml"
    if prs_key not in shared_files:
        return

    NS_PML = "http://schemas.openxmlformats.org/presentationml/2006/main"
    NS_R = "http://schemas.openxmlformats.org/officeDocument/2006/relationships"

    root = ET.fromstring(shared_files[prs_key])
    sldIdLst = root.find(f"{{{NS_PML}}}sldIdLst")
    if sldIdLst is None:
        sldIdLst = ET.SubElement(root, f"{{{NS_PML}}}sldIdLst")
    for child in list(sldIdLst):
        sldIdLst.remove(child)

    base_id = 256
    for i in range(1, slide_count + 1):
        sldId = ET.SubElement(sldIdLst, f"{{{NS_PML}}}sldId")
        sldId.set("id", str(base_id + i))
        sldId.set(f"{{{NS_R}}}id", f"rId{i}")

    shared_files[prs_key] = ET.tostring(
        root, xml_declaration=True, encoding="UTF-8", short_empty_elements=False
    )


def _update_prs_rels_bytes(shared_files: dict, slide_count: int):
    """shared_files 내 ppt/_rels/presentation.xml.rels 슬라이드 관계 업데이트."""
    rels_key = "ppt/_rels/presentation.xml.rels"
    if rels_key not in shared_files:
        return

    NS = "http://schemas.openxmlformats.org/package/2006/relationships"
    SLIDE_TYPE = "http://schemas.openxmlformats.org/officeDocument/2006/relationships/slide"

    root = ET.fromstring(shared_files[rels_key])

    # 기존 슬라이드 관계 제거
    for el in [e for e in root if e.get("Type") == SLIDE_TYPE]:
        root.remove(el)

    for i in range(1, slide_count + 1):
        rel = ET.SubElement(root, f"{{{NS}}}Relationship")
        rel.set("Id", f"rId{i}")
        rel.set("Type", SLIDE_TYPE)
        rel.set("Target", f"slides/slide{i}.xml")

    shared_files[rels_key] = ET.tostring(
        root, xml_declaration=True, encoding="UTF-8", short_empty_elements=False
    )


# ---------------------------------------------------------------------------
# 메인
# ---------------------------------------------------------------------------

def main():
    parser = argparse.ArgumentParser(
        description="extract_component_template — Plan A PPTX에서 레이아웃별 대표 슬라이드 추출"
    )
    parser.add_argument("--draft",  required=True, help="draft JSON 파일 경로")
    parser.add_argument("--source", required=True, help="Plan A PPTX 파일 경로")
    parser.add_argument("--output", required=True, help="출력 component_template PPTX 경로")
    args = parser.parse_args()

    print(f"draft : {args.draft}")
    print(f"source: {args.source}")
    print(f"output: {args.output}")
    print("추출 시작...")

    extract_component_template(args.draft, args.source, args.output)
    print("완료.")


if __name__ == "__main__":
    main()
