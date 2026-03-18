"""
pptxgenjs 파일럿 슬라이드를 PNG로 변환해 HTML 프리뷰와 비교 이미지 생성
pymupdf 사용 (Poppler 불필요)
"""
import fitz  # pymupdf
import sys, os

pptx_path = os.path.join(os.path.dirname(__file__),
    "../../outputs/pilot_pptxgenjs_slides_5_8.pptx")
out_dir = os.path.join(os.path.dirname(__file__), "../../outputs/pilot_png")
os.makedirs(out_dir, exist_ok=True)

# pymupdf는 직접 PPTX를 열 수 없음 — LibreOffice 없이는 PPTX→PDF 변환 불가
# 대신 PPTX 내부 슬라이드 PNG 추출 시도 (pptx 썸네일)
import zipfile
from PIL import Image
import io

with zipfile.ZipFile(pptx_path) as z:
    names = z.namelist()
    # 썸네일 추출 (있을 경우)
    thumbs = [n for n in names if 'thumbnail' in n.lower() or 'preview' in n.lower()]
    slides = sorted([n for n in names if n.startswith('ppt/slides/slide') and not 'Rel' in n])

    print(f"슬라이드 파일: {slides}")
    print(f"썸네일: {thumbs}")

    if thumbs:
        for i, t in enumerate(thumbs):
            img_data = z.read(t)
            img = Image.open(io.BytesIO(img_data))
            out_path = os.path.join(out_dir, f"thumbnail_{i+1}.png")
            img.save(out_path)
            print(f"저장: {out_path} ({img.size})")
    else:
        print("썸네일 없음 — PowerPoint에서 직접 열어 repair 메시지 확인 필요")

print("\n[검증 체크리스트]")
print("1. outputs/pilot_pptxgenjs_slides_5_8.pptx 를 PowerPoint에서 열기")
print("2. repair 메시지 없으면 → pptxgenjs 전환 가능")
print("3. 차트·표·텍스트 레이아웃 확인")
print("4. 폰트(Pretendard) 렌더링 확인")
