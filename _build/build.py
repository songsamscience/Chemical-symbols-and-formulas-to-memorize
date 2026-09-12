# -*- coding: utf-8 -*-
"""
한 번에 다시 만들기:  python3 _build/build.py
  1) out/worksheet.pdf  (4쪽 학습지, 페이퍼로지 글꼴 내장)
  2) out/workbook.xlsx  (통합 엑셀 자료)
  3) 학습 사이트 HTML 안의 데이터 블록과 다운로드 메뉴(base64)를 chemdata.py 기준으로 갱신
필요한 것: Python 3 + openpyxl, pymupdf(fitz) / Chrome 또는 Edge / 페이퍼로지 글꼴
"""
import base64, json, os, re, shutil, subprocess, sys

HERE = os.path.dirname(os.path.abspath(__file__))
PROJECT = os.path.dirname(HERE)
SITE = os.path.join(PROJECT, "송쌤과학_2022개정_화학기호와화학식_학습사이트.html")
INDEX = os.path.join(PROJECT, "index.html")
OUT = os.path.join(HERE, "out")
PDF_NAME = "송쌤과학_화학기호와화학식_학습지_4쪽.pdf"
XLSX_NAME = "송쌤과학_화학기호와화학식_통합정리.xlsx"

sys.path.insert(0, HERE)
from chemdata import ELEMENTS, MOLECULES, IONS, COMPOUNDS

def js(value):
    return json.dumps(value, ensure_ascii=False, separators=(",", ":"))

def data_block():
    """사이트 스크립트의 chemistryData 블록. 각 항목: [식/기호, 이름, 핵심(1/0), 다른 이름 목록, 추가 정보]"""
    rows = {
        "elements": [js([e["symbol"], e["name"], int(e["core"]), e["alias"], e["z"]]) for e in ELEMENTS],
        "molecules": [js([m["formula"], m["name"], int(m["core"]), [], m["kind"], int(m["pages"] != "–")]) for m in MOLECULES],
        "ions": [js([i["formula"], i["name"], int(i["core"]), i["alias"]]) for i in IONS],
        "compounds": [js([c["formula"], c["name"], 1, c["alias"], f'{c["cation"]} + {c["anion"]} · {c["ratio"]}', int(c["cation_count"] == c["anion_count"])]) for c in COMPOUNDS],
    }
    body = ",\n".join(f"      {key}: [\n        {','.join(items)}\n      ]" for key, items in rows.items())
    return f"    const chemistryData = {{\n{body}\n    }};\n"

def downloads_block():
    pdf = base64.b64encode(open(os.path.join(OUT, "worksheet.pdf"), "rb").read()).decode()
    xlsx = base64.b64encode(open(os.path.join(OUT, "workbook.xlsx"), "rb").read()).decode()
    core_el = sum(1 for e in ELEMENTS if e["core"])
    return f'''        <div class="download-menu" id="downloadMenu" hidden>
          <a href="data:application/pdf;base64,{pdf}" download="{PDF_NAME}">
            <span class="file-icon" aria-hidden="true">PDF</span>
            <span><strong>4쪽 학습지</strong><span>이름 ↔ 기호·식 · 원소 {core_el}종(원자 번호 표시)</span></span>
          </a>
          <a href="data:application/vnd.openxmlformats-officedocument.spreadsheetml.sheet;base64,{xlsx}" download="{XLSX_NAME}">
            <span class="file-icon" aria-hidden="true">XLSX</span>
            <span><strong>통합 엑셀 자료</strong><span>원소·분자·이온·화합물 정리 (7개 시트)</span></span>
          </a>
        </div>
'''

def replace_between(text, start_marker, end_marker, replacement):
    pattern = re.compile(re.escape(start_marker) + r".*?\n(.*?)" + re.escape(end_marker), re.S)
    matches = pattern.findall(text)
    if len(matches) != 1:
        sys.exit(f"표식을 찾지 못했습니다: {start_marker[:30]}… ({len(matches)}개)")
    return pattern.sub(lambda m: m.group(0).replace(m.group(1), replacement), text, count=1)

def run(script):
    r = subprocess.run([sys.executable, os.path.join(HERE, script)], cwd=HERE)
    if r.returncode != 0:
        sys.exit(f"{script} 실패")

if __name__ == "__main__":
    only_site = "--site-only" in sys.argv
    if not only_site:
        run("make_pdf.py")
        run("make_xlsx.py")
    html = open(SITE, encoding="utf-8").read()
    html = replace_between(html, "/* build:data:start", "    /* build:data:end */", data_block())
    html = replace_between(html, "<!-- build:downloads:start", "        <!-- build:downloads:end -->", downloads_block())
    open(SITE, "w", encoding="utf-8").write(html)
    if os.path.exists(INDEX):                 # GitHub Pages용 index.html 사본도 같은 내용으로 맞춥니다
        shutil.copy(SITE, INDEX)
    shutil.copy(os.path.join(OUT, "worksheet.pdf"), os.path.join(OUT, PDF_NAME))
    shutil.copy(os.path.join(OUT, "workbook.xlsx"), os.path.join(OUT, XLSX_NAME))
    print(f"사이트 갱신 완료: {os.path.basename(SITE)} ({os.path.getsize(SITE)//1024} KB)")
    print(f"따로 쓸 파일: _build/out/{PDF_NAME}, _build/out/{XLSX_NAME}")
