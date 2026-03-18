---
name: pptx-autofill-conversion
description: >
  PPTX 양식(템플릿)을 첨부하고 주제를 입력하면, 동일한 레이아웃/슬라이드 구조를
  완벽히 유지하면서 새로운 주제에 맞게 내용을 자동으로 교체하여 완성된 PPTX 파일을
  생성한다. 표·도형·그룹 구조를 XML 수준에서 분석하고 텍스트만 정밀 교체한다.

  트리거 조건 — 아래 상황 중 하나라도 해당되면 반드시 이 skill을 사용하라:
  - pptx 파일 첨부 + "이 양식으로 [주제] 만들어줘" / "내용만 바꿔줘"
  - "템플릿 유지하고 주제 바꿔줘" / "양식에 맞춰 작성해줘"
  - "pptx 자동 채우기" / "슬라이드 자동 생성"
  - PPTX 첨부 후 새 주제나 기관명·강사명·과정명 등 변경 요청 시
---

# PPTX 양식 자동 채우기 Skill

## 전체 워크플로우 요약

```
Step 0  입력 확인          → PPTX 파일 + 주제 수집
Step 1  텍스트 추출        → markitdown으로 슬라이드 구조 파악
Step 2  XML 파싱
  ├─ Step 2-A  표 심층 분석   → 행/열/셀 속성 매핑
  └─ Step 2-B  도형 심층 분석 → 유형별 교체 가능 여부 판단
Step 3  콘텐츠 생성        → 주제에 맞는 텍스트 content_map 작성
Step 4  XML 교체 & 재조립  → <a:t> 텍스트만 교체 후 팩킹
Step 5  QA                 → 내용·시각 검증 후 최종 파일 제공
```

---

## Step 0. 입력 확인

작업 시작 전 아래 두 가지를 반드시 확인하라.

**필수 입력값:**
- `template.pptx` — 레이아웃 원본 양식 파일
- `new_topic` — 변경할 주제 (예: "삼성전자 AI 활용 과정", "2026 신입사원 온보딩")

**선택 입력값 (있으면 콘텐츠 품질 향상):**
- 강사명, 기관명, 날짜, 교육 대상, 교육 시간 등 메타 정보

두 가지가 모두 확인되면 즉시 Step 1로 진행하라. 주제가 없으면 한 번만 질문하라.

---

## Step 1. 텍스트 추출 및 슬라이드 구조 파악

```bash
# 의존성 설치
pip install "markitdown[pptx]" --break-system-packages -q

# 텍스트 추출 (슬라이드 번호, 제목, 표, 리스트 구조 확인)
python -m markitdown template.pptx

# XML 압축 해제 (이후 모든 단계의 기반)
python /mnt/skills/public/pptx/scripts/office/unpack.py template.pptx unpacked/
```

추출 결과로 아래 정보를 파악하라:

| 확인 항목 | 내용 |
|---|---|
| 슬라이드 수 | 최종 출력도 동일 슬라이드 수 유지 |
| 슬라이드별 역할 | 커버/목차/본문/강사 프로필 등 |
| 표 존재 여부 | 있으면 Step 2-A 실행 |
| 도형/그룹 존재 여부 | 있으면 Step 2-B 실행 |
| 이미지 플레이스홀더 위치 | 교체 금지 목록에 기록 |

---

## Step 2-A. 표(Table) 구조 심층 분석

PPTX의 표는 `<p:graphicFrame>` 내부 `<a:tbl>`로 존재한다.
아래 코드로 슬라이드별 표 구조 전체를 추출하라.

```python
import zipfile, re
from xml.etree import ElementTree as ET

NS = {
    'a': 'http://schemas.openxmlformats.org/drawingml/2006/main',
    'p': 'http://schemas.openxmlformats.org/presentationml/2006/main',
}

with zipfile.ZipFile('template.pptx') as z:
    slide_files = sorted([f for f in z.namelist()
                          if re.match(r'ppt/slides/slide\d+\.xml', f)])
    for slide_name in slide_files:
        tree = ET.fromstring(z.read(slide_name))
        for gf in tree.findall('.//p:graphicFrame', NS):
            tbl = gf.find('.//a:tbl', NS)
            if tbl is None:
                continue

            # ① 표 전체 위치/크기
            xfrm = gf.find('p:xfrm', NS)
            off  = xfrm.find('a:off', NS)
            ext  = xfrm.find('a:ext', NS)
            print(f"\n[{slide_name}] 표 위치: x={off.get('x')}, y={off.get('y')}")
            print(f"  크기: w={ext.get('cx')}, h={ext.get('cy')}")

            # ② 열 너비 (gridCol) — 절대 변경 금지
            cols = [g.get('w') for g in tbl.findall('a:tblGrid/a:gridCol', NS)]
            print(f"  열 수: {len(cols)}, 열 너비: {cols}")

            # ③ 표 스타일 ID
            tbl_pr = tbl.find('a:tblPr', NS)
            style_id = tbl_pr.find('a:tableStyleId', NS)
            print(f"  스타일 ID: {style_id.text if style_id is not None else '없음'}")

            # ④ 행별 높이 + 셀 텍스트/서식
            for r_i, tr in enumerate(tbl.findall('a:tr', NS)):
                print(f"\n  행 {r_i+1} (높이={tr.get('h')})")
                for c_i, tc in enumerate(tr.findall('a:tc', NS)):

                    # 셀 텍스트
                    texts = [t.text for t in tc.findall('.//a:t', NS) if t.text]

                    # 셀 배경색 (solidFill 또는 gradFill)
                    tcPr = tc.find('a:tcPr', NS)
                    fill = tcPr.find('.//a:solidFill/a:srgbClr', NS) if tcPr else None
                    grad = tcPr.find('.//a:gradFill', NS) if tcPr else None
                    bg   = f"#{fill.get('val')}" if fill is not None else \
                           ('gradient' if grad is not None else 'noFill')

                    # 텍스트 서식 (첫 번째 rPr 기준)
                    rpr  = tc.find('.//a:rPr', NS)
                    sz   = rpr.get('sz') if rpr is not None else '?'
                    bold = rpr.get('b')  if rpr is not None else '?'
                    tc_fill = rpr.find('.//a:solidFill/a:srgbClr', NS) if rpr else None
                    tc_color = f"#{tc_fill.get('val')}" if tc_fill is not None else 'scheme'

                    # 테두리 정보
                    lnB = tcPr.find('a:lnB', NS) if tcPr else None
                    border_fill = lnB.find('.//a:srgbClr', NS) if lnB is not None else None
                    border_color = f"#{border_fill.get('val')}" if border_fill else 'noFill'

                    print(f"    셀[{r_i+1},{c_i+1}]")
                    print(f"      텍스트  : {'|'.join(texts) if texts else '(없음)'}")
                    print(f"      배경색  : {bg}")
                    print(f"      글자색  : {tc_color}, 크기: {sz}, 굵기: {bold}")
                    print(f"      하단 테두리: {border_color}")
```

**분석 결과를 아래 XML 형식으로 정리하라:**

```xml
<table_map slide="N" position="x=? y=?" size="w=? h=?">
  <grid cols="4" widths="1971589, 3888606, 4408371, 1003718"/>
  <style_id>{5C22544A-7EE6-4342-B048-85BDC9FD1C3A}</style_id>

  <row index="1" height="312863" role="header">
    <cell col="1" bg="#5A6580" text_color="bg1(white)" font_size="1200" bold="0">모듈</cell>
    <cell col="2" bg="#5A6580" text_color="bg1(white)" font_size="1200" bold="0">학습 목표</cell>
    <cell col="3" bg="#5A6580" text_color="bg1(white)" font_size="1200" bold="0">상세 내용</cell>
    <cell col="4" bg="#5A6580" text_color="bg1(white)" font_size="1200" bold="0">미리보기</cell>
  </row>
  <row index="2" height="980000" role="data">
    <cell col="1" bg="noFill" font_size="1100" bold="1">AI 업무자동화...</cell>
    <cell col="2" bg="noFill" font_size="1000" bold="0">생성형 AI의 동작...</cell>
    <cell col="3" bg="noFill" font_size="1000" bold="0">프롬프트 엔지니어링...</cell>
    <cell col="4" bg="noFill" font_size="1000" bold="0">(이미지)</cell>
  </row>
  <!-- 나머지 행 동일 패턴 -->
</table_map>
```

**⚠ 표에서 절대 변경 금지 속성:**

| 태그 | 이유 |
|---|---|
| `<a:tblGrid>` / `<a:gridCol w="...">` | 열 너비 — 변경 시 레이아웃 붕괴 |
| `<a:tr h="...">` | 행 높이 — 변경 시 셀 오버플로우 |
| `<a:tcPr>` 내 `<a:lnL/R/T/B>` | 셀 테두리 스타일 |
| `<a:solidFill>` (헤더 배경) | 헤더/본문 색상 구분 |
| `<a:tableStyleId>` | 표 전체 테마 |
| `<a:rPr sz="..." b="...">` | 폰트 크기/굵기 |

**✅ 표에서 교체 가능한 것:** 각 `<a:tc>` 내부의 `<a:t>텍스트</a:t>` 만

---

## Step 2-B. 도형(Shape) 구조 심층 분석

PPTX 도형은 3가지 계층으로 존재한다:

```
p:spTree (슬라이드 전체 컨테이너)
 ├── p:sp           → 단일 도형 (텍스트박스, 사각형, 원, 선 등)
 ├── p:grpSp        → 그룹 도형 (여러 도형을 하나로 묶음)
 │    ├── p:sp      → 그룹 내 개별 도형
 │    └── p:sp
 └── p:graphicFrame → 표(table) 또는 차트
```

아래 코드로 슬라이드별 도형 구조를 전수 추출하라:

```python
import zipfile, re
from xml.etree import ElementTree as ET

NS = {
    'a': 'http://schemas.openxmlformats.org/drawingml/2006/main',
    'p': 'http://schemas.openxmlformats.org/presentationml/2006/main',
}

def get_pos_size(spPr, ns):
    xfrm = spPr.find('a:xfrm', ns) if spPr is not None else None
    off  = xfrm.find('a:off', ns)  if xfrm is not None else None
    ext  = xfrm.find('a:ext', ns)  if xfrm is not None else None
    pos  = f"x={off.get('x')}, y={off.get('y')}" if off is not None else '?'
    size = f"w={ext.get('cx')}, h={ext.get('cy')}" if ext is not None else '?'
    return pos, size

def get_fill(spPr, ns):
    solid = spPr.find('.//a:solidFill/a:srgbClr', ns) if spPr is not None else None
    grad  = spPr.find('.//a:gradFill', ns)             if spPr is not None else None
    nofill= spPr.find('a:noFill', ns)                  if spPr is not None else None
    if   solid  is not None: return f"solid(#{solid.get('val')})"
    elif grad   is not None: return 'gradient'
    elif nofill is not None: return 'noFill'
    else:                    return 'unknown'

def get_border(spPr, ns):
    ln = spPr.find('a:ln', ns) if spPr is not None else None
    if ln is None: return 'none'
    w = ln.get('w', '?')
    c = ln.find('.//a:srgbClr', ns)
    color = f"#{c.get('val')}" if c is not None else 'scheme'
    return f"w={w}, color={color}"

def parse_sp(sp, ns, depth=0):
    indent = '  ' * depth
    cNvPr  = sp.find('p:nvSpPr/p:cNvPr', ns)
    sp_id  = cNvPr.get('id')   if cNvPr is not None else '?'
    sp_nm  = cNvPr.get('name') if cNvPr is not None else '?'

    spPr   = sp.find('p:spPr', ns)
    geom   = spPr.find('.//a:prstGeom', ns) if spPr is not None else None
    sp_type= geom.get('prst') if geom is not None else 'custom/freeform'

    # 라운드 사각형 모서리 조절값
    adj_el = spPr.find('.//a:prstGeom/a:avLst/a:gd[@name="adj"]', ns) if spPr else None
    adj    = adj_el.get('fmla') if adj_el is not None else None

    pos, size = get_pos_size(spPr, ns)
    fill      = get_fill(spPr, ns)
    border    = get_border(spPr, ns)

    # placeholder 여부
    ph = sp.find('p:nvSpPr/p:nvPr/p:ph', ns)
    ph_type = ph.get('type', 'body') if ph is not None else None

    # 내부 텍스트 (단락별)
    paragraphs = []
    for para in sp.findall('.//a:p', ns):
        texts = [t.text for t in para.findall('.//a:t', ns) if t.text]
        if texts:
            paragraphs.append(''.join(texts))

    # 텍스트 서식 (첫 번째 rPr)
    rpr  = sp.find('.//a:rPr', ns)
    sz   = rpr.get('sz') if rpr is not None else '?'
    bold = rpr.get('b')  if rpr is not None else '?'
    font_el = rpr.find('.//a:latin', ns) if rpr is not None else None
    font = font_el.get('typeface') if font_el is not None else '?'
    clr_el = rpr.find('.//a:solidFill/a:srgbClr', ns) if rpr is not None else None
    clr  = f"#{clr_el.get('val')}" if clr_el is not None else 'scheme'

    print(f"{indent}[도형 {sp_id}] {sp_nm}")
    print(f"{indent}  유형={sp_type}{f'(adj={adj})' if adj else ''}"
          f"{f'  PH={ph_type}' if ph_type else ''}")
    print(f"{indent}  위치={pos}  크기={size}")
    print(f"{indent}  채우기={fill}  테두리={border}")
    print(f"{indent}  폰트={font}  크기={sz}  굵기={bold}  색={clr}")
    print(f"{indent}  텍스트: {paragraphs if paragraphs else '없음'}")

with zipfile.ZipFile('template.pptx') as z:
    slide_files = sorted([f for f in z.namelist()
                          if re.match(r'ppt/slides/slide\d+\.xml', f)])
    for slide_name in slide_files:
        tree = ET.fromstring(z.read(slide_name))
        print(f"\n{'='*60}\n{slide_name}\n{'='*60}")
        spTree = tree.find('.//p:spTree', NS)

        for child in spTree:
            tag = child.tag.split('}')[-1]
            if tag == 'sp':
                parse_sp(child, NS, depth=0)

            elif tag == 'grpSp':
                grp_cNvPr = child.find('p:nvGrpSpPr/p:cNvPr', NS)
                grp_nm    = grp_cNvPr.get('name', '?') if grp_cNvPr else '?'
                grp_xfrm  = child.find('p:grpSpPr/a:xfrm', NS)
                grp_off   = grp_xfrm.find('a:off', NS)  if grp_xfrm else None
                grp_ext   = grp_xfrm.find('a:ext', NS)  if grp_xfrm else None
                chOff     = grp_xfrm.find('a:chOff', NS) if grp_xfrm else None
                chExt     = grp_xfrm.find('a:chExt', NS) if grp_xfrm else None
                print(f"\n[그룹: {grp_nm}]")
                if grp_off:
                    print(f"  그룹 위치: x={grp_off.get('x')}, y={grp_off.get('y')}"
                          f"  크기: w={grp_ext.get('cx')}, h={grp_ext.get('cy')}")
                if chOff:
                    print(f"  내부 좌표계: chOff x={chOff.get('x')}, y={chOff.get('y')}"
                          f"  chExt w={chExt.get('cx')}, h={chExt.get('cy')}")
                for sp in child.findall('.//p:sp', NS):
                    parse_sp(sp, NS, depth=1)

            elif tag == 'graphicFrame':
                gf_cNvPr = child.find('p:nvGraphicFramePr/p:cNvPr', NS)
                gf_nm    = gf_cNvPr.get('name', '?') if gf_cNvPr else '?'
                print(f"\n[GraphicFrame: {gf_nm}] → 표/차트 (Step 2-A에서 별도 분석)")
```

**분석 결과를 아래 XML 형식으로 정리하라:**

```xml
<shape_map slide="N">

  <!-- ① 제목 플레이스홀더 -->
  <shape id="22" name="제목 2" type="title_placeholder">
    <position x="385618" y="390293" w="11471420" h="535464"/>
    <style fill="none" border="none"/>
    <font typeface="나눔스퀘어 ExtraBold" size="2000" color="scheme"/>
    <text editable="true">슬라이드 제목</text>
    <rule>제목 텍스트만 교체. 폰트·위치·크기 고정.</rule>
  </shape>

  <!-- ② 원형 그라디언트 도형 (이미지 컨테이너) -->
  <shape id="5" name="타원 4" type="ellipse">
    <position x="1381876" y="1162756" w="2263472" h="2263472"/>
    <style fill="gradient(#3396F0→#3396F0)" border="w=38100 color=bg1"/>
    <text editable="false">없음</text>
    <rule>⛔ 수정 금지. 이미지 플레이스홀더.</rule>
  </shape>

  <!-- ③ 라운드 사각형 카드 박스 -->
  <shape id="2" name="직사각형 14" type="roundRect" adj="val 5773">
    <position x="562898" y="4502262" w="3856701" h="839358"/>
    <style fill="bg1(alpha=10000)" border="w=3175 color=bg1(lum85)"/>
    <font typeface="나눔스퀘어 Bold" size="1000"/>
    <text editable="true">카드 내용 텍스트</text>
    <rule>텍스트만 교체. adj(모서리 둥글기) 고정.</rule>
  </shape>

  <!-- ④ 구분선 (line) -->
  <shape id="37" name="Line 247" type="line">
    <position x="1382233" y="2315519" w="4508204" h="0"/>
    <style border="w=15875 color=#969696"/>
    <text editable="false">없음</text>
    <rule>⛔ 수정 금지. 구분선은 레이아웃 요소.</rule>
  </shape>

  <!-- ⑤ 그룹 도형 -->
  <group name="그룹 34" position="x=6395817 y=1170811 w=5362899 h=184666"
         chOff="x=527538 y=2226443" chExt="w=5362899 h=184666">
    <shape id="36" name="직사각형 35" type="rect">
      <font typeface="나눔스퀘어 Bold" size="1200" color="#009CE1"/>
      <text editable="true">섹션 레이블 텍스트</text>
      <rule>텍스트만 교체. 색상(#009CE1) 고정.</rule>
    </shape>
    <shape id="37" name="Line 247" type="line">
      <rule>⛔ 수정 금지. 그룹 내 구분선.</rule>
    </shape>
  </group>

</shape_map>
```

**도형 유형별 교체 규칙:**

| 도형 유형 | 텍스트 교체 | 절대 변경 금지 |
|---|---|---|
| `rect` / `roundRect` | ✅ `<a:t>` 텍스트 | 위치, 크기, `adj`(모서리 반지름) |
| `ellipse` (원) | ❌ 없음 | 전체 고정 (이미지 컨테이너) |
| `line` (선) | ❌ 없음 | 전체 고정 (구분선/레이아웃) |
| `title_placeholder` | ✅ 제목 텍스트 | 크기, 폰트 계열 |
| `grpSp` (그룹) | ✅ 내부 `<a:t>` | 그룹 좌표 (`chOff`, `chExt`) |
| `graphicFrame` | Step 2-A 참조 | 표 구조 전체 |

**⚠ 도형에서 절대 변경 금지 속성:**

| 속성 | 이유 |
|---|---|
| `<a:xfrm>` (위치/크기) | 도형 레이아웃 붕괴 |
| `<a:prstGeom prst="...">` | 도형 유형 변경 금지 |
| `<a:avLst>` / `<a:gd name="adj">` | 모서리·비율 등 도형 커스텀 수치 |
| `<a:gradFill>` / `<a:solidFill>` | 색상 테마 |
| `<a:ln w="...">` | 테두리/선 두께 |
| `p:grpSpPr/a:xfrm` 내 `chOff` / `chExt` | 그룹 내부 좌표계 |

**⚠ 슬라이드 외부 요소 처리 규칙:**

XML 파싱 중 `<a:off>` 의 x 또는 y 값이 음수인 도형/표/그룹은 **완전히 무시**한다.
이 요소들은 슬라이드 가시 영역 밖에 위치하므로 autofill 대상에 포함시키지 않는다.

```python
# 슬라이드 외부 요소 필터 (도형 파싱 시 적용)
off = xfrm.find('a:off', NS) if xfrm else None
if off is not None:
    x_val = int(off.get('x', 0))
    y_val = int(off.get('y', 0))
    if x_val < 0 or y_val < 0:
        continue  # 슬라이드 외부 요소 스킵
```

---

## Step 3. 주제에 맞는 콘텐츠 생성

Step 2-A·B에서 만든 `table_map`과 `shape_map`을 기반으로
새 주제의 텍스트를 `content_map` XML 형식으로 작성하라.

**콘텐츠 생성 규칙:**

1. **글자 수 제약** — 원본 텍스트의 ±30% 범위 내로 작성
   - 원본 셀 텍스트 20자 → 신규 텍스트 14~26자 이내
   - 넘칠 경우 내용을 압축하거나 줄바꿈(`<a:br/>`) 활용

2. **슬라이드 수 고정** — 슬라이드 추가/삭제 금지
   (슬라이드가 부족하면 내용을 병합, 초과하면 핵심만 선별)

3. **표 행/열 수 고정** — 데이터 행이 부족하면 원본과 동일 행 수로 내용 분산 배치

4. **메타 정보 우선 반영** — 강사명·기관명·날짜 등 Step 0에서 수집한 값 먼저 채움

5. **이미지 플레이스홀더 제외** — `<rule>⛔ 수정 금지</rule>` 도형은 content_map에 포함시키지 않음

6. **다단계 서식 구조 보존 (중요)** — 텍스트박스 내에 font size가 서로 다른 단락이 혼재할 경우, 반드시 해당 구조를 그대로 유지하라.

   **패턴 감지:** 같은 텍스트박스 안에 `sz=1400` (소제목 레이블) + `sz=900` (본문 bullet) 단락이 섞여 있는 경우가 이에 해당한다.

   ```
   원본 구조:
     단락 1: sz=1400, color=#D98F76 → "Cap. Rate"        ← 소제목 (짧은 레이블)
     단락 2: sz=900,  color=scheme  → "2025년 1분기..."  ← 본문 bullet
     단락 3: sz=900,  color=scheme  → "Cap. Rate는..."   ← 본문 bullet
     단락 4: sz=900,  color=scheme  → "금리 변동에..."   ← 본문 bullet
   ```

   **잘못된 처리 (금지):**
   ```xml
   <shape id="6" name="TextBox 6">
     <text>해상풍력 설치 현황 및 Cap Rate 추이와 금리 변동의 영향...</text>
   </shape>
   <!-- ❌ 소제목 자리에 모든 내용을 몰아넣음 → sz=1400 빨간 글씨로 전부 표시됨 -->
   ```

   **올바른 처리:**
   ```xml
   <shape id="6" name="TextBox 6">
     <paragraph>해상풍력 수익률 지표</paragraph>  <!-- 소제목: 짧게 (sz=1400 자리) -->
     <paragraph>2025년 해상풍력 Cap Rate는 약 6.5~7.5% 수준으로...</paragraph>  <!-- 본문 (sz=900 자리) -->
     <paragraph>국내 REC 가중치 조정으로 수익률 변동성 확대...</paragraph>
     <paragraph>금리 인하 기조에 따라 Cap Rate 하락 압력 지속...</paragraph>
   </shape>
   <!-- ✅ 단락 수를 원본과 맞춰 각 서식 자리에 적절한 내용 배치 -->
   ```

   **규칙 요약:**
   - 첫 번째 단락(`<paragraph>`) → 소제목 레이블 역할, **15자 이내 짧게**
   - 나머지 단락들 → 본문 bullet, **원본 단락 수 ±1 이내로 유지**
   - 원본 단락 수보다 내용이 많으면 압축, 부족하면 빈 단락(`<paragraph></paragraph>`) 추가

**출력 형식:**

```xml
<content_map topic="[새 주제명]">

  <slide number="1">
    <!-- 단순 텍스트 박스 -->
    <shape id="22" name="제목 2">
      <text>[새 슬라이드 제목]</text>
    </shape>

    <!-- 그룹 내 섹션 레이블 -->
    <group name="그룹 34">
      <shape id="36" name="직사각형 35">
        <text>[새 섹션 레이블]</text>
      </shape>
    </group>

    <!-- 일반 텍스트 박스 (bullet) -->
    <shape id="40" name="Rectangle 23">
      <paragraph>교육 대상 : [대상]</paragraph>
      <paragraph>사전 지식 : [필요 지식]</paragraph>
      <paragraph>활용 TOOL : [사용 도구]</paragraph>
    </shape>
  </slide>

  <slide number="2">
    <!-- 표 교체: 텍스트만, 행/열 구조 고정 -->
    <table id="21" name="표 20">
      <row index="1" role="header">
        <!-- 헤더는 원본 텍스트 그대로 유지하거나 새 주제에 맞게 변경 -->
        <cell col="1">[모듈명 컬럼]</cell>
        <cell col="2">[학습목표 컬럼]</cell>
        <cell col="3">[상세내용 컬럼]</cell>
        <cell col="4">[비고 컬럼]</cell>
      </row>
      <row index="2">
        <cell col="1">[모듈 1 이름]</cell>
        <cell col="2">[모듈 1 학습목표]</cell>
        <cell col="3">[모듈 1 상세내용]</cell>
        <cell col="4">[미리보기 텍스트 or 빈값]</cell>
      </row>
      <!-- 행 수는 원본과 동일하게 유지 -->
    </table>
  </slide>

  <slide number="3">
    <!-- 강사/프로필 슬라이드 -->
    <shape id="22" name="제목 2">
      <text>[강사 이름] 프로필</text>
    </shape>
    <shape id="2" name="직사각형 14">
      <text>[저서 또는 주요 이력]</text>
    </shape>
    <!-- 원형 도형(타원 4)은 content_map에서 제외 (이미지 컨테이너) -->
  </slide>

</content_map>
```

---

## Step 4. XML 교체 및 PPTX 재조립

`content_map`을 기반으로 실제 XML 파일을 수정하라.

### 4-1. 텍스트 교체 원칙

**반드시 `<a:t>` 태그 안의 텍스트만 교체하라.**
`<a:rPr>` (서식), `<a:pPr>` (단락 설정), `<a:xfrm>` (위치) 는 절대 수정 금지.

```xml
<!-- ❌ 잘못된 교체: 서식 태그까지 재작성 -->
<a:r>
  <a:t>새 텍스트</a:t>
</a:r>

<!-- ✅ 올바른 교체: rPr 유지, t 태그만 변경 -->
<a:r>
  <a:rPr lang="ko-KR" sz="1200" b="1" kern="0" dirty="0">
    <a:solidFill><a:srgbClr val="009CE1"/></a:solidFill>
    <a:latin typeface="나눔스퀘어 Bold" .../>
  </a:rPr>
  <a:t>새 텍스트</a:t>  <!-- ← 여기만 변경 -->
</a:r>
```

### 4-2. 다단락 텍스트 교체 (bullet 리스트)

여러 줄 내용은 `<a:p>` 단락을 분리하여 작성하라.
원본 단락의 `<a:pPr>` (들여쓰기, 줄간격)을 반드시 복사하라.

```xml
<!-- ✅ 올바른 다단락 교체 -->
<a:p>
  <a:pPr marL="171450" indent="-171450" fontAlgn="base">
    <a:lnSpc><a:spcPct val="150000"/></a:lnSpc>
    <a:buChar char="•"/>
  </a:pPr>
  <a:r><a:rPr lang="ko-KR" sz="1100" .../><a:t>첫 번째 항목</a:t></a:r>
</a:p>
<a:p>
  <a:pPr marL="171450" indent="-171450" fontAlgn="base">
    <a:lnSpc><a:spcPct val="150000"/></a:lnSpc>
    <a:buChar char="•"/>
  </a:pPr>
  <a:r><a:rPr lang="ko-KR" sz="1100" .../><a:t>두 번째 항목</a:t></a:r>
</a:p>
```

### 4-3. 표 셀 교체

`<a:tc>` 내부의 `<a:t>` 만 교체. `<a:tcPr>` (셀 배경/테두리) 절대 수정 금지.

```xml
<!-- ✅ 올바른 셀 텍스트 교체 -->
<a:tc>
  <a:txBody>
    <a:bodyPr/>
    <a:lstStyle/>
    <a:p>
      <a:pPr algn="ctr">...</a:pPr>
      <a:r>
        <a:rPr lang="ko-KR" sz="1100" b="1" .../>
        <a:t>새 모듈명</a:t>  <!-- ← 여기만 변경 -->
      </a:r>
    </a:p>
  </a:txBody>
  <a:tcPr ...>  <!-- ← 절대 수정 금지 -->
    ...
  </a:tcPr>
</a:tc>
```

### 4-4. 특수문자 처리

한국어 따옴표·특수문자는 XML 엔터티로 처리하라:

| 문자 | XML 엔터티 |
|---|---|
| `"` (left)  | `&#x201C;` |
| `"` (right) | `&#x201D;` |
| `'` (left)  | `&#x2018;` |
| `'` (right) | `&#x2019;` |
| `&` | `&amp;` |

### 4-5. 파일 재조립

```bash
# XML 정리 (고아 파일 제거)
python /mnt/skills/public/pptx/scripts/clean.py unpacked/

# PPTX 재패킹 (원본 바이너리 구조 보존)
python /mnt/skills/public/pptx/scripts/office/pack.py \
  unpacked/ output.pptx --original template.pptx
```

---

## Step 5. QA 검증

### 5-1. 콘텐츠 검증

```bash
# 텍스트 추출로 내용 확인
python -m markitdown output.pptx

# 잔여 원본 텍스트 체크 (원본 키워드로 검색)
python -m markitdown output.pptx | grep -iE "임성수|동원그룹|AI 실습과정|강사명_원본"
```

grep 결과가 나오면 해당 슬라이드 XML로 돌아가 교체가 누락된 `<a:t>` 를 찾아 수정하라.

### 5-2. 구조 검증

```python
# 슬라이드 수 일치 확인
import zipfile, re

original_slides = len([f for f in zipfile.ZipFile('template.pptx').namelist()
                       if re.match(r'ppt/slides/slide\d+\.xml', f)])
output_slides   = len([f for f in zipfile.ZipFile('output.pptx').namelist()
                       if re.match(r'ppt/slides/slide\d+\.xml', f)])

assert original_slides == output_slides, \
    f"슬라이드 수 불일치: 원본 {original_slides} vs 출력 {output_slides}"
print(f"✅ 슬라이드 수 일치: {output_slides}장")
```

### 5-3. 시각 검증 (최종 확인)

```bash
# PDF 변환 후 이미지로 렌더링
python /mnt/skills/public/pptx/scripts/office/soffice.py \
  --headless --convert-to pdf output.pptx
pdftoppm -jpeg -r 150 output.pdf slide

# 생성 이미지 목록 확인
ls slide-*.jpg
```

렌더링 이미지로 아래 항목을 육안 확인하라:

| 확인 항목 | 판단 기준 |
|---|---|
| 텍스트 오버플로우 | 텍스트가 박스 밖으로 나가지 않음 |
| 표 구조 유지 | 행/열 수, 헤더 색상 원본과 동일 |
| 도형 위치/크기 | 원본과 동일한 레이아웃 |
| 이미지 플레이스홀더 | 원형/사각형 컨테이너 그대로 존재 |
| 폰트/색상 | 원본 서식 유지 |
| 잔여 원본 텍스트 없음 | 모든 교체 대상 텍스트 변경 완료 |

이슈 발견 시 해당 슬라이드 XML로 복귀하여 수정 후 재패킹·재검증하라.

---

## 최종 파일 제공

모든 QA 통과 후:

```bash
cp output.pptx /mnt/user-data/outputs/[주제명]_완성.pptx
```

`present_files` 도구로 파일을 사용자에게 전달하라.

---

## 빠른 참조 — 교체 가능/불가 요약

```
✅ 교체 가능
  <a:t> 태그 내 텍스트 (도형, 표 셀, 제목 플레이스홀더)

⛔ 절대 변경 금지
  <a:xfrm>       위치/크기
  <a:prstGeom>   도형 유형
  <a:avLst>      도형 커스텀 수치 (모서리 등)
  <a:solidFill>  색상
  <a:gradFill>   그라디언트
  <a:ln>         테두리/선
  <a:rPr>        텍스트 서식 (폰트·크기·굵기)
  <a:pPr>        단락 설정 (줄간격·들여쓰기)
  <a:tblGrid>    표 열 너비
  <a:tr h="..."> 표 행 높이
  <a:tcPr>       셀 배경·테두리
  p:grpSpPr      그룹 좌표계
```
