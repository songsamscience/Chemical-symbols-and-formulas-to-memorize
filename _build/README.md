# _build — 학습지 PDF · 엑셀 · 사이트 데이터 만들기

학습 사이트(`../송쌤과학_2022개정_화학기호와화학식_학습사이트.html`)에 들어가는 내용은
모두 **`chemdata.py` 한 곳**에서 나옵니다.

| 파일 | 하는 일 |
| --- | --- |
| `chemdata.py` | 원소 · 분자식 · 이온 · 이온 결합 화합물 · 이온 이름 규칙 데이터 (여기만 고치면 됨) |
| `make_pdf.py` | 4쪽 학습지 PDF. HTML로 조판 → Chrome 인쇄 → 페이퍼로지 글꼴을 PDF 안에 내장 |
| `make_xlsx.py` | 통합 엑셀 자료(7개 시트) |
| `build.py` | 위 둘을 실행한 뒤 사이트 HTML의 데이터 블록과 다운로드 메뉴(base64)를 바꿔 넣음 |
| `out/` | 생성물 (자동으로 다시 만들어지므로 git에 올리지 않음) |

## 다시 만들기

```bash
python3 _build/build.py
```

- 데이터만 바꾸고 PDF·엑셀은 그대로 둘 때: `python3 _build/build.py --site-only`
- 필요한 것: Python 3, `pip install openpyxl pymupdf`, Chrome 또는 Edge, **페이퍼로지(Paperlogy) 글꼴 설치**
  (설치 대신 `_build/fonts/` 폴더에 `Paperlogy-4Regular.ttf`, `Paperlogy-7Bold.ttf`, `Paperlogy-8ExtraBold.ttf`를 넣어도 됩니다)
- 사이트 HTML 안의 `build:data` / `build:downloads` 표식 사이는 자동 생성 구간이니 직접 고치지 마세요.
  화면·문구·기능은 표식 바깥(HTML/CSS/JS)에서 자유롭게 고치면 됩니다.

## GitHub에 올릴 때

- 사이트는 **HTML 파일 하나로 완결**됩니다(PDF·엑셀이 안에 들어 있음). GitHub Pages는 이 파일만 있으면 됩니다.
- `_build/` 폴더는 사이트 동작과 무관합니다. 나중에 다시 만들 수 있도록 함께 올려 두는 것을 권하지만, 빼도 사이트는 그대로 돌아갑니다.
- 내용을 바꿀 때 순서: `chemdata.py` 수정 → `python3 _build/build.py` → 바뀐 HTML(과 chemdata.py)을 커밋·푸시. `_build/out/`은 `.gitignore`로 제외돼 있습니다.
- Pages 주소를 짧게 쓰려면 HTML 파일을 `index.html`로 복사(또는 이름 변경)해 저장소 루트에 두면 됩니다. `build.py`의 `SITE` 경로도 같이 바꿔 주세요.
