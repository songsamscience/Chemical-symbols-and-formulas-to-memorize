# -*- coding: utf-8 -*-
"""4쪽 학습지 PDF 생성: HTML 조판 → 헤드리스 Chrome 인쇄 → PyMuPDF로 용량 정리.
글꼴은 페이퍼로지(Paperlogy) TTF를 PDF 안에 직접 심으므로 윈도우·맥 어디서 열어도 같게 보입니다."""
import base64, glob, html, os, shutil, subprocess, sys, time
from chemdata import ELEMENTS, MOLECULES, IONS, COMPOUNDS, formula_html, is_polyatomic

HERE = os.path.dirname(os.path.abspath(__file__))
OUT_DIR = os.path.join(HERE, "out")
OUT_HTML = os.path.join(OUT_DIR, "worksheet.html")
OUT_PDF = os.path.join(OUT_DIR, "worksheet.pdf")

CHROME_CANDIDATES = [
    "/Applications/Google Chrome.app/Contents/MacOS/Google Chrome",
    "/Applications/Microsoft Edge.app/Contents/MacOS/Microsoft Edge",
    r"C:\Program Files\Google\Chrome\Application\chrome.exe",
    r"C:\Program Files (x86)\Google\Chrome\Application\chrome.exe",
    r"C:\Program Files (x86)\Microsoft\Edge\Application\msedge.exe",
    "/usr/bin/google-chrome", "/usr/bin/chromium", "/usr/bin/chromium-browser",
]
FONT_DIRS = [
    os.path.join(HERE, "fonts"),                       # _build/fonts/ 에 복사해 두면 이곳을 먼저 씁니다
    os.path.expanduser("~/Library/Fonts"), "/Library/Fonts",
    os.path.join(os.environ.get("LOCALAPPDATA", ""), "Microsoft", "Windows", "Fonts"),
    os.path.join(os.environ.get("WINDIR", r"C:\Windows"), "Fonts"),
]
FONT_FILES = {400: "Paperlogy-4Regular.ttf", 700: "Paperlogy-7Bold.ttf", 800: "Paperlogy-8ExtraBold.ttf"}

def find_font(filename):
    for d in FONT_DIRS:
        p = os.path.join(d, filename)
        if os.path.exists(p):
            return p
    sys.exit(f"페이퍼로지 글꼴을 찾지 못했습니다: {filename}\n  → 페이퍼로지를 설치하거나 TTF 파일을 _build/fonts/ 폴더에 복사한 뒤 다시 실행하세요.")

def font_face_css():
    rules = []
    for weight, filename in FONT_FILES.items():
        data = base64.b64encode(open(find_font(filename), "rb").read()).decode()
        rules.append(f'@font-face {{ font-family: "PaperlogyPDF"; font-weight: {weight}; font-style: normal; src: url("data:font/ttf;base64,{data}") format("truetype"); }}')
    return "\n".join(rules)

core_elements = [e for e in ELEMENTS if e["core"]]
core_ions = [i for i in IONS if i["core"]]

def esc(s): return html.escape(s, quote=False)

def split2(items):
    half = (len(items) + 1) // 2
    return items[:half], items[half:]

MOL_HEAD = '<span class="head-chips"><span class="chip element">원소 — 한 가지 원소</span><span class="chip compound">화합물 — 두 가지 이상의 원소</span></span>'
ION_LEGEND = '<div class="legend"><span class="chip cation">양이온</span><span class="chip anion">음이온</span><span class="chip poly">다원자 이온 (보라 바탕)</span></div>'
CMP_LEGEND = '<div class="legend"><span class="chip same">개수비 1 : 1 — 전하 크기가 같음</span><span class="chip diff">개수비 다름 — 아래첨자로 이온 수 표시</span></div>'

def section(letter, title, caption, items, render, cls, row_cls="", legend="", head_chips=""):
    left, right = split2(items)
    rows = lambda part, start: "".join(render(i + start, it) for i, it in enumerate(part))
    return f'''
    <section class="sec {cls}">
      <div class="sec-head"><span class="letter">{letter}</span><span class="sec-title">{title}</span>{head_chips}<span class="sec-caption">{caption}</span></div>
      {legend}
      <div class="two-col">
        <div class="rows {row_cls}">{rows(left, 1)}</div>
        <div class="rows {row_cls}">{rows(right, len(left) + 1)}</div>
      </div>
    </section>'''

# ---- 행 그리기 ------------------------------------------------------------
def el_symbol_row(n, e):
    return f'<div class="row"><span class="z">{e["z"]}</span><span class="prompt sym">{e["symbol"]}</span><span class="blank"></span></div>'

def el_name_row(n, e):
    label = esc(e["name"]) + (f'<small>({"·".join(e["alias"])})</small>' if e["alias"] else "")
    return f'<div class="row"><span class="z">{e["z"]}</span><span class="prompt">{label}</span><span class="blank short"></span></div>'

def mol_class(m):
    return "element" if m["kind"] == "원소" else "compound"

def mol_formula_row(n, m):
    return f'<div class="row"><span class="n">{n}</span><span class="prompt sym {mol_class(m)}">{formula_html(m["formula"])}</span><span class="blank"></span></div>'

def mol_name_row(n, m):
    return f'<div class="row"><span class="n">{n}</span><span class="prompt {mol_class(m)}">{esc(m["name"])}</span><span class="blank short"></span></div>'

def ion_classes(i):
    """양이온/음이온은 글자색, 다원자 이온은 행 바탕색으로 구분"""
    return ("cation" if i["charge"] > 0 else "anion"), (" poly" if is_polyatomic(i["formula"]) else "")

def ion_formula_row(n, i):
    charge, poly = ion_classes(i)
    return f'<div class="row{poly}"><span class="n">{n}</span><span class="prompt sym {charge}">{formula_html(i["formula"])}</span><span class="blank"></span></div>'

def ion_name_row(n, i):
    charge, poly = ion_classes(i)
    return f'<div class="row{poly}"><span class="n">{n}</span><span class="prompt {charge}">{esc(i["name"])}</span><span class="blank short"></span></div>'

def ratio_class(c):
    return "same" if c["cation_count"] == c["anion_count"] else "diff"

def cmp_formula_row(n, c):
    return f'<div class="row"><span class="n">{n}</span><span class="prompt sym {ratio_class(c)}">{formula_html(c["formula"])}</span><span class="blank"></span></div>'

def cmp_name_row(n, c):
    return f'''<div class="row tall"><span class="n">{n}</span><div class="stack"><span class="prompt {ratio_class(c)}">{esc(c["name"])}</span>
      <span class="parts"><span class="lab">이온</span><span class="mini"></span><span class="plus">+</span><span class="mini"></span><span class="lab">개수비</span><span class="mini xs"></span><span class="lab">화학식</span><span class="mini wide"></span></span></div></div>'''

# ---- 쪽 틀 ---------------------------------------------------------------
def page(num, total, banner, body, note=""):
    return f'''
  <article class="page">
    <header class="head">
      <div class="brand-row"><span class="brand">송쌤과학</span><span class="eyebrow">2022 개정 교육과정 중학교 과학 · Ⅳ. 물질의 구성</span></div>
      <div class="title-row">
        <div><h1>꼭 알아둘 화학 기호와 화학식</h1><p class="subtitle">원소 · 분자 · 이온 · 이온 결합 화합물 4쪽 학습지</p></div>
        <div class="fields"><span><b>학년·반</b><i></i></span><span><b>이름</b><i></i></span></div>
      </div>
      <div class="banner"><span class="pg">{num} / {total}</span>{banner}</div>
    </header>
    {body}
    {f'<div class="note">{note}</div>' if note else ''}
    <footer><span>송쌤과학 · 아주중학교  |  정답은 학습 사이트에서 확인하세요.</span><span>{num} / {total}</span></footer>
  </article>'''

def build_pages():
    p1 = page(1, 4, "원소 기호 · 분자식  —  기호와 식을 보고 <b>이름</b> 쓰기",
        section("A", "원소 기호 → 원소 이름", "앞의 숫자는 원자 번호입니다", core_elements, el_symbol_row, "navy")
        + section("B", "분자식 → 물질 이름", "", MOLECULES, mol_formula_row, "green", head_chips=MOL_HEAD))
    p2 = page(2, 4, "원소 기호 · 분자식  —  이름을 보고 <b>기호와 식</b> 쓰기",
        section("A", "원소 이름 → 원소 기호", "앞의 숫자는 원자 번호입니다 · 괄호 안은 IUPAC 이름", core_elements, el_name_row, "navy")
        + section("B", "물질 이름 → 분자식", "원자 수는 아래첨자로", MOLECULES, mol_name_row, "green", head_chips=MOL_HEAD))
    p3 = page(3, 4, "이온식 · 이온 결합 화합물  —  식을 보고 <b>이름</b> 쓰기",
        section("C", "이온식 → 이온 이름", "", core_ions, ion_formula_row, "blue", "roomy-xl", ION_LEGEND)
        + section("D", "이온 결합 화합물의 화학식 → 이름", "", COMPOUNDS, cmp_formula_row, "coral", "roomy-xl", CMP_LEGEND),
        note=formula_html("<b>이름 규칙</b>  양이온: 원소 이름 + <u>이온</u> (나트륨 → 나트륨 이온 Na⁺)  ·  음이온: 원소 이름 + <u>화 이온</u> (황 → 황화 이온 S²⁻), 이름이 ‘소’로 끝나면 ‘소’를 떼고 (염소 → 염화 이온 Cl⁻, 산소 → 산화 이온 O²⁻)  ·  다원자 이온은 고유한 이름 (암모늄·수산화·질산·탄산·황산·과망가니즈산 이온)"))
    p4 = page(4, 4, "이온식 · 이온 결합 화합물  —  이름을 보고 <b>식</b> 쓰기",
        section("C", "이온 이름 → 이온식", "전하까지 함께 씁니다", core_ions, ion_name_row, "blue", "roomy", ION_LEGEND)
        + section("D", "화합물 이름 → 이온 · 개수비 · 화학식", "전체 전하의 합 = 0", COMPOUNDS, cmp_name_row, "coral", "", CMP_LEGEND),
        note=formula_html("<b>쓰는 법</b>  전하는 위첨자(Cu²⁺), 원자 수는 아래첨자(H₂O)  ·  전하가 1이면 숫자 1을 생략(Na⁺, Cl⁻)  ·  숫자는 전하 기호 앞에(Cu²⁺ ○, Cu⁺² ×)  ·  다원자 이온이 2개 이상이면 괄호로 묶기(Ca(OH)₂)"))
    return p1 + p2 + p3 + p4

CSS = r'''
@page { size: A4; margin: 0; }
* { box-sizing: border-box; }
html, body { margin: 0; padding: 0; }
body { font-family: "PaperlogyPDF", "Paperlogy", "Apple SD Gothic Neo", "Malgun Gothic", sans-serif; color: #172235; font-size: 10.5pt; line-height: 1.25; -webkit-print-color-adjust: exact; print-color-adjust: exact; }
.page { position: relative; width: 210mm; height: 297mm; padding: 11mm 12mm 9mm; page-break-after: always; overflow: hidden; display: flex; flex-direction: column; }
.page:last-child { page-break-after: auto; }
.head { flex: none; }
.brand-row { display: flex; align-items: center; gap: 8px; }
.brand { padding: 2px 9px; border-radius: 6px; background: #1f3a5f; color: #fff; font-weight: 700; font-size: 8.5pt; }
.eyebrow { color: #5a6b80; font-size: 8.5pt; font-weight: 400; }
.title-row { display: flex; align-items: flex-end; justify-content: space-between; margin-top: 3mm; }
h1 { margin: 0; font-size: 20pt; font-weight: 800; letter-spacing: -.02em; color: #1f3a5f; line-height: 1.1; }
.subtitle { margin: 2mm 0 0; color: #6b7684; font-size: 8.5pt; }
.fields { display: flex; gap: 14px; }
.fields span { display: flex; align-items: flex-end; gap: 6px; font-size: 8pt; color: #5a6b80; }
.fields b { font-weight: 700; }
.fields i { display: inline-block; width: 26mm; border-bottom: 1px solid #8a97a8; height: 5mm; }
.banner { margin-top: 3mm; padding: 2mm 3.5mm; border-radius: 3mm; background: #eef3f8; color: #1f3a5f; font-weight: 700; font-size: 11pt; display: flex; align-items: center; gap: 3mm; }
.banner .pg { padding: .6mm 2mm; border-radius: 2mm; background: #1f3a5f; color: #fff; font-size: 8.5pt; }
.banner b { color: #d65a64; }

.sec { margin-top: 3.2mm; border: 1px solid #d8e0ea; border-radius: 3mm; overflow: hidden; }
.sec-head { display: flex; align-items: center; gap: 2mm; padding: 1.6mm 3mm; color: #fff; font-weight: 700; font-size: 10pt; }
.sec.navy .sec-head { background: #1f3a5f; } .sec.green .sec-head { background: #1f7a68; }
.sec.blue .sec-head { background: #0f6fb8; } .sec.coral .sec-head { background: #d65a64; }
.letter { display: inline-grid; place-items: center; width: 5.2mm; height: 5.2mm; border-radius: 1.5mm; background: rgba(255,255,255,.22); font-size: 9pt; }
.sec-caption { margin-left: auto; font-weight: 400; font-size: 8pt; opacity: .92; }
.legend { display: flex; flex-wrap: wrap; gap: 1.6mm; padding: 1.1mm 3mm; background: #fbfcfe; border-bottom: 1px solid #e3e9f0; font-size: 7.6pt; }
.chip { padding: .4mm 2mm; border-radius: 999px; border: 1px solid; font-weight: 700; }
.head-chips { display: inline-flex; gap: 1.4mm; margin-left: 2mm; font-size: 7.4pt; }
.head-chips .chip { border-color: transparent; background: rgba(255,255,255,.94); }
.chip.cation { color: #0f6fb8; border-color: #b9d6f2; background: #eaf3fc; }
.chip.anion { color: #c4404b; border-color: #f2c4c8; background: #fdeeef; }
.chip.poly { color: #7a3fb0; border-color: #dcc6f3; background: #f3ebfb; }
.chip.element { color: #1f7a68; border-color: #bfe3d9; background: #e9f7f3; }
.chip.compound { color: #1f3a5f; border-color: #c9d5e3; background: #eef3f8; }
.prompt.element { color: #1f7a68; } .prompt.compound { color: #1f3a5f; }
.chip.same { color: #1f7a68; border-color: #bfe3d9; background: #e9f7f3; }
.chip.diff { color: #b7791f; border-color: #f1dcb0; background: #fff4dc; }
.prompt.cation { color: #0f6fb8; } .prompt.anion { color: #c4404b; }
.prompt.same { color: #1f7a68; } .prompt.diff { color: #b7791f; }
.row.poly, .row.poly:nth-child(even) { background: #f3ebfb; }
.two-col { display: grid; grid-template-columns: 1fr 1fr; }
.two-col .rows + .rows { border-left: 1px solid #e3e9f0; }
.row { display: flex; align-items: center; gap: 2mm; height: 9.6mm; padding: 0 3mm; border-bottom: 1px solid #edf1f5; }
.rows.roomy .row { height: 10mm; }
.rows.roomy-xl .row { height: 11.6mm; }
.row.tall { height: 14.4mm; }
.row:last-child { border-bottom: 0; }
.row:nth-child(even) { background: #f7f9fc; }
.sec.coral .row:nth-child(even) { background: #fff5f4; }
.n { flex: none; width: 5.5mm; color: #8a97a8; font-size: 8pt; text-align: right; }
.z { flex: none; min-width: 7.2mm; padding: 0 1.2mm; height: 5.4mm; border-radius: 1.4mm; background: #e6eef8; color: #1f3a5f; font-size: 8pt; font-weight: 700; text-align: center; line-height: 5.4mm; }
.prompt { flex: none; min-width: 27mm; font-size: 11.5pt; font-weight: 700; }
.prompt small { margin-left: 1mm; font-size: 8pt; color: #6b7684; font-weight: 400; }
.prompt.sym { font-size: 13pt; font-weight: 700; letter-spacing: .01em; }
.prompt.poly { color: #7a3fb0; }
sub, sup { font-size: .62em; line-height: 0; position: relative; vertical-align: baseline; }
sub { top: .32em; } sup { top: -.5em; }
.blank { flex: 1; height: 6.8mm; border-bottom: 1.2px solid #9aa7b8; margin-right: 1mm; }
.blank.short { flex: 1; max-width: 40mm; margin-left: auto; }
.stack { flex: 1; display: flex; flex-direction: column; gap: 1.2mm; min-width: 0; }
.parts { display: flex; align-items: flex-end; gap: 1.6mm; font-size: 8pt; color: #6b7684; }
.parts .lab { flex: none; }
.parts .plus { flex: none; color: #1f3a5f; font-weight: 700; font-size: 10pt; }
.mini { flex: 1; height: 5mm; border-bottom: 1px solid #9aa7b8; }
.mini.xs { flex: .55; } .mini.wide { flex: 1.6; }
.note { margin-top: auto; padding: 2mm 3mm; border-radius: 2.5mm; background: #f3f7fb; color: #4c5d72; font-size: 8pt; line-height: 1.5; }
.note b { color: #1f3a5f; margin-right: 1.5mm; }
.note u { text-decoration: none; color: #d65a64; font-weight: 700; }
footer { flex: none; display: flex; justify-content: space-between; margin-top: 3mm; padding-top: 2mm; border-top: 1px solid #d8e0ea; color: #8a97a8; font-size: 8pt; }
'''

def find_chrome():
    for c in CHROME_CANDIDATES:
        if os.path.exists(c):
            return c
    sys.exit("Chrome 또는 Edge를 찾지 못했습니다. make_pdf.py의 CHROME_CANDIDATES에 경로를 추가하세요.")

def print_to_pdf():
    os.makedirs(OUT_DIR, exist_ok=True)
    doc = f'<!doctype html><html lang="ko"><head><meta charset="utf-8"><title>꼭 알아둘 화학 기호와 화학식 학습지</title><style>{font_face_css()}\n{CSS}</style></head><body>{build_pages()}</body></html>'
    open(OUT_HTML, "w", encoding="utf-8").write(doc)
    profile = os.path.join(OUT_DIR, ".chrome-profile")
    shutil.rmtree(profile, ignore_errors=True)
    if os.path.exists(OUT_PDF):
        os.remove(OUT_PDF)
    cmd = [find_chrome(), "--headless=new", "--disable-gpu", "--no-first-run", "--no-default-browser-check",
           f"--user-data-dir={profile}", "--no-pdf-header-footer", f"--print-to-pdf={OUT_PDF}", "file://" + OUT_HTML]
    proc = subprocess.Popen(cmd, stdout=subprocess.DEVNULL, stderr=subprocess.DEVNULL)
    for _ in range(240):                      # Chrome이 파일을 쓴 뒤 종료하지 않는 경우가 있어 파일 생성만 기다립니다
        time.sleep(0.5)
        if os.path.exists(OUT_PDF) and os.path.getsize(OUT_PDF) > 1000:
            time.sleep(2); break
    proc.kill()
    shutil.rmtree(profile, ignore_errors=True)
    if not os.path.exists(OUT_PDF):
        sys.exit("PDF가 만들어지지 않았습니다.")

def optimize():
    import fitz  # PyMuPDF
    d = fitz.open(OUT_PDF)
    d.set_metadata({"title": "꼭 알아둘 화학 기호와 화학식 · 4쪽 학습지", "author": "송쌤과학", "subject": "2022 개정 중학교 과학 Ⅳ. 물질의 구성"})
    cat = d.pdf_catalog()
    for k in ("StructTreeRoot", "MarkInfo"):    # 접근성 태그 트리는 크기만 키우므로 뺍니다
        if k in d.xref_get_keys(cat):
            d.xref_set_key(cat, k, "null")
    d.save(OUT_PDF + ".tmp", garbage=4, deflate=True, clean=True, use_objstms=1)
    d.close()
    os.replace(OUT_PDF + ".tmp", OUT_PDF)

def report():
    import fitz
    d = fitz.open(OUT_PDF)
    fonts = sorted({f[3].split("+")[-1] for p in d for f in p.get_fonts()})
    print(f"PDF: {len(d)}쪽, {os.path.getsize(OUT_PDF)//1024} KB, 심어진 글꼴: {', '.join(fonts)}")

if __name__ == "__main__":
    print_to_pdf()
    optimize()
    report()
