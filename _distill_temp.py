"""
임시 ref-distiller 스크립트:
[SRCIG] 2호_디지털인프라 섹터의 데이터센터_22.4Q.pdf 파싱 및 캐시 생성
"""
import os, sys, json, hashlib, datetime

FILE_PATH = r"C:\Side project\ppt generator\references\report\narratives\[SRCIG] 2호_디지털인프라 섹터의 데이터센터_22.4Q.pdf"
CACHE_DIR = r"C:\Side project\ppt generator\references\report\narratives\.cache"
CACHE_FILE = os.path.join(CACHE_DIR, "[SRCIG] 2호_디지털인프라 섹터의 데이터센터_22.4Q.pdf.json")

# ── 파일 메타데이터 ─────────────────────────────────────────────────────────
stat = os.stat(FILE_PATH)
file_size  = stat.st_size
file_mtime = stat.st_mtime

def file_hash(path):
    h = hashlib.sha256()
    with open(path, 'rb') as f:
        for chunk in iter(lambda: f.read(65536), b''):
            h.update(chunk)
    return h.hexdigest()

fhash = file_hash(FILE_PATH)

# ── 캐시 유효성 검사 ────────────────────────────────────────────────────────
os.makedirs(CACHE_DIR, exist_ok=True)
try:
    existing = json.load(open(CACHE_FILE, encoding='utf-8'))
    if (existing.get("file_size") == file_size and
            existing.get("file_mtime") == file_mtime):
        print("캐시 유효 (size+mtime 일치): 재처리 건너뜀")
        sys.exit(0)
    if existing.get("file_hash") == f"sha256:{fhash}":
        print("캐시 유효 (SHA256 일치): 재처리 건너뜀")
        sys.exit(0)
except FileNotFoundError:
    pass

# ── PDF 파싱 ────────────────────────────────────────────────────────────────
# pdfplumber 시도 → 실패 시 PyMuPDF(fitz) 시도
pages_text = []
parser_used = None

try:
    import pdfplumber
    with pdfplumber.open(FILE_PATH) as pdf:
        total_pages = len(pdf.pages)
        for i, page in enumerate(pdf.pages):
            txt = page.extract_text() or ""
            pages_text.append({"page": i+1, "text": txt})
    parser_used = "pdfplumber"
    print(f"pdfplumber 파싱 완료: 총 {total_pages}페이지")
except Exception as e1:
    print(f"pdfplumber 실패: {e1}")
    try:
        import fitz  # PyMuPDF
        doc = fitz.open(FILE_PATH)
        total_pages = len(doc)
        for i, page in enumerate(doc):
            txt = page.get_text("text") or ""
            pages_text.append({"page": i+1, "text": txt})
        doc.close()
        parser_used = "pymupdf"
        print(f"PyMuPDF 파싱 완료: 총 {total_pages}페이지")
    except Exception as e2:
        print(f"PyMuPDF 실패: {e2}")
        print("PDF 파싱 불가 — 파서 없음")
        sys.exit(1)

# ── 텍스트 출력 (분석용) ────────────────────────────────────────────────────
print("\n========== 전체 텍스트 미리보기 ==========")
for p in pages_text:
    print(f"\n--- 페이지 {p['page']} ---")
    print(p['text'][:800])  # 페이지당 최대 800자 출력
print("\n========== 끝 ==========")
print(f"\n총 페이지 수: {total_pages}")
print(f"파서: {parser_used}")
print(f"파일 해시: sha256:{fhash}")
print(f"파일 크기: {file_size}")
print(f"파일 mtime: {file_mtime}")
