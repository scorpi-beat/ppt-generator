"""
파일럿 슬라이드 5,8을 HTML로 렌더링해서 pptxgenjs 결과물과 나란히 비교하는 페이지 생성
python src/pilot/generate_pilot_html.py
"""
import json, os, sys
sys.stdout.reconfigure(encoding='utf-8')

BASE = os.path.join(os.path.dirname(__file__), "../..")
draft = json.load(open(f"{BASE}/outputs/draft_report_solar_power_invest_v2.json", encoding="utf-8"))
style = json.load(open(f"{BASE}/outputs/style_report.json", encoding="utf-8"))

slides = {s["slide_number"]: s for s in draft["slides"]}
s5, s8 = slides[5], slides[8]

COLORS = style["colors"]
FONTS  = style["fonts"]

C_PRIMARY = COLORS["primary"].lstrip("#")
C_SALMON  = COLORS["accent_salmon"].lstrip("#")
C_ACCENT  = COLORS["accent_dark"].lstrip("#")
C_WARM    = COLORS["accent_warm"].lstrip("#")
C_TEXT    = COLORS["text_main"].lstrip("#")
C_NEUTRAL = COLORS["neutral_light"].lstrip("#")

FONT     = FONTS["primary_typeface"]
BODY_PT  = max(FONTS["body"]["size_pt"], 9)
TABLE_PT = max(FONTS["table_body"]["size_pt"], 9)
HEAD_PT  = FONTS["head_message"]["size_pt"]


def header_html(slide, page_num):
    sec = slide.get("section_number", "")
    sec_label = f'<div class="sec-label">Section {sec}</div>' if sec else ""
    head = slide.get("head_message", "")
    return f"""
<div class="slide-header">
  {sec_label}
  <div class="slide-title">{slide.get('title','')}</div>
  <div class="page-num">{page_num}</div>
</div>
<div class="head-banner">{head}</div>
"""


def bar_chart_svg(chart_data, width=420, height=220):
    data   = chart_data["data"]
    labels = [d["label"] for d in data]
    values = [d["value"] for d in data]
    mx     = max(values) * 1.1
    n      = len(data)
    ML, MR, MT, MB = 32, 8, 12, 28
    aw = width - ML - MR
    ah = height - MT - MB
    bar_w = aw / n * 0.6
    gap   = aw / n

    bars_svg = ""
    for i, (lbl, val) in enumerate(zip(labels, values)):
        bh   = ah * val / mx
        bx   = ML + i * gap + (gap - bar_w) / 2
        by   = MT + ah - bh
        bars_svg += f'<rect x="{bx:.1f}" y="{by:.1f}" width="{bar_w:.1f}" height="{bh:.1f}" fill="#{C_PRIMARY}" rx="1"/>\n'
        bars_svg += f'<text x="{bx + bar_w/2:.1f}" y="{by - 3:.1f}" text-anchor="middle" font-size="7" fill="#{C_TEXT}">{val}</text>\n'
        bars_svg += f'<text x="{bx + bar_w/2:.1f}" y="{MT + ah + 12:.1f}" text-anchor="middle" font-size="7" fill="#{C_TEXT}">{lbl}</text>\n'

    # Y축 눈금선 4개
    grid_svg = ""
    for k in range(1, 5):
        gy = MT + ah * (1 - k / 4)
        gv = int(mx * k / 4)
        grid_svg += f'<line x1="{ML}" y1="{gy:.1f}" x2="{ML+aw}" y2="{gy:.1f}" stroke="#DDDDDD" stroke-width="0.5"/>\n'
        grid_svg += f'<text x="{ML-3}" y="{gy+3:.1f}" text-anchor="end" font-size="7" fill="#999">{gv}</text>\n'

    title = chart_data.get("title", "")
    return f"""<svg viewBox="0 0 {width} {height}" xmlns="http://www.w3.org/2000/svg" style="width:100%;height:100%;">
  <text x="{width//2}" y="10" text-anchor="middle" font-size="8" fill="#{C_ACCENT}" font-family="{FONT}">{title}</text>
  {grid_svg}
  {bars_svg}
</svg>"""


def table_html(tbl, font_size=None, first_col_bold=True, stripe=True):
    fs = font_size or TABLE_PT
    rows_html = ""
    for ri, row in enumerate(tbl["rows"]):
        bg = f"#{C_NEUTRAL}" if (stripe and ri % 2 == 0) else "#FFFFFF"
        cells = ""
        for ci, cell in enumerate(row):
            bold = "font-weight:600;" if (first_col_bold and ci == 0) else ""
            align = "left" if ci == 0 else "center"
            cells += f'<td style="padding:4px 6px;font-size:{fs}px;text-align:{align};{bold}background:{bg};border:1px solid #E0E0E0;">{cell}</td>'
        rows_html += f"<tr>{cells}</tr>"
    headers = "".join(
        f'<th style="padding:4px 6px;font-size:{fs}px;background:#{C_PRIMARY};color:#FFF;text-align:{"left" if i==0 else "center"};border:1px solid #{C_PRIMARY};">{h}</th>'
        for i, h in enumerate(tbl["headers"])
    )
    return f'<table style="width:100%;border-collapse:collapse;font-family:{FONT};"><thead><tr>{headers}</tr></thead><tbody>{rows_html}</tbody></table>'


def render_slide5(s):
    chart_svg = bar_chart_svg(s["main_zone"]["chart"], width=420, height=200)
    bullets = s["sub_zone_top"]["bullets"]
    bullet_html = "".join(
        f'<div style="margin-bottom:5px;font-size:{BODY_PT}px;color:#{C_TEXT};line-height:1.5;">'
        f'<span style="color:#{C_WARM};font-weight:600;margin-right:4px;">{i+1}.</span>{b}</div>'
        for i, b in enumerate(bullets)
    )
    mini_tbl = table_html(s["sub_zone_bottom"]["table"], font_size=TABLE_PT)
    note = s.get("note_box", {})
    note_html = f'<div class="note-box">{note.get("content","")}</div>' if note else ""

    return f"""
<div class="slide" id="slide5">
  {header_html(s, 5)}
  <div class="content-area" style="display:flex;gap:10px;">
    <div style="flex:0 0 53%;height:100%;">{chart_svg}</div>
    <div style="flex:1;display:flex;flex-direction:column;gap:8px;overflow:hidden;">
      <div style="flex:1;">{bullet_html}</div>
      <div style="flex:0 0 auto;">{mini_tbl}</div>
    </div>
  </div>
  {note_html}
</div>"""


def render_slide8(s):
    tbl = table_html(s["table"])
    kps = s.get("key_points", [])
    kp_html = ""
    if kps:
        items = "".join(
            f'<div style="font-size:{BODY_PT}px;color:#{C_TEXT};margin-bottom:4px;line-height:1.5;">'
            f'<span style="color:#{C_PRIMARY};margin-right:4px;">●</span>{kp}</div>'
            for kp in kps
        )
        kp_html = f'<div style="border-top:1px solid #DDD;padding-top:6px;margin-top:8px;">{items}</div>'

    return f"""
<div class="slide" id="slide8">
  {header_html(s, 8)}
  <div class="content-area">
    {tbl}
    {kp_html}
  </div>
</div>"""


CSS = f"""
* {{ box-sizing: border-box; margin: 0; padding: 0; }}
body {{ background: #E8E8E8; font-family: '{FONT}', 'Pretendard', sans-serif;
       padding: 20px; display: flex; flex-direction: column; gap: 24px; align-items: center; }}
.slide {{
  width: 960px; height: 540px; background: #FFFFFF;
  box-shadow: 0 4px 16px rgba(0,0,0,0.2);
  display: flex; flex-direction: column; overflow: hidden; position: relative;
}}
.slide-header {{
  background: #{C_PRIMARY}; height: 50px;
  padding: 5px 18px; display: flex; flex-direction: column;
  justify-content: center; position: relative;
}}
.sec-label {{ font-size: 9px; color: rgba(255,255,255,0.65); line-height: 1.2; margin-bottom: 1px; }}
.slide-title {{ font-size: 15px; color: #FFFFFF; font-weight: 800; line-height: 1.2; }}
.page-num {{ position: absolute; right: 16px; top: 50%; transform: translateY(-50%);
             font-size: 9px; color: rgba(255,255,255,0.55); }}
.head-banner {{
  background: #FFFFFF; border-left: 3px solid #{C_SALMON};
  padding: 4px 18px 4px 14px; min-height: 28px;
  font-size: {HEAD_PT}px; color: #{C_SALMON}; font-weight: 600;
  display: flex; align-items: center; line-height: 1.3;
}}
.content-area {{
  flex: 1; padding: 8px 16px 4px 16px; overflow: hidden;
}}
.note-box {{
  font-size: 7px; color: #888; padding: 2px 16px 3px;
  border-top: 1px solid #EEE; font-style: italic;
}}
h2 {{ font-size: 14px; color: #333; margin-bottom: 6px; }}
"""

html = f"""<!DOCTYPE html>
<html lang="ko">
<head>
<meta charset="UTF-8">
<title>pptxgenjs 파일럿 비교 — 슬라이드 5, 8</title>
<style>{CSS}</style>
</head>
<body>
<h2>Slide 5: composite_split (bar chart + bullets + mini-table)</h2>
{render_slide5(s5)}
<h2>Slide 8: wide_table (LCOE 비교)</h2>
{render_slide8(s8)}
</body>
</html>"""

out = f"{BASE}/outputs/pilot_compare_slides_5_8.html"
with open(out, "w", encoding="utf-8") as f:
    f.write(html)
print(f"저장: {out}")
