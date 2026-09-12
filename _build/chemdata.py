# -*- coding: utf-8 -*-
"""
학습 사이트 · PDF 학습지 · 엑셀 자료가 공통으로 쓰는 데이터입니다.
내용을 바꾼 뒤 `python3 _build/build.py` 를 실행하면 세 곳에 모두 반영됩니다.

core=True  : 학습 핵심 (사이트의 '교과서·학습 핵심만' 범위이자 PDF 학습지에 실리는 항목, 엑셀 연노란색)
core=False : 수업 확장 (사이트에서 '핵심만'을 끄면 나옴, 엑셀에 색 없이 수록)
"""

def element(z, symbol, name, core, pages="–", alias=()):
    return dict(z=z, symbol=symbol, name=name, core=core, pages=pages, alias=list(alias))

def molecule(formula, name, core, kind, pages="–"):
    return dict(formula=formula, name=name, core=core, kind=kind, pages=pages)

def ion(formula, name, core, element_name, symbol, z, charge, textbook, pages, note, alias=()):
    return dict(formula=formula, name=name, core=core, element_name=element_name, symbol=symbol, z=z,
                charge=charge, textbook=textbook, pages=pages, note=note, alias=list(alias))

def compound(name, cation, cation_charge, cation_count, anion, anion_charge, anion_count, formula, memo, alias=()):
    return dict(name=name, cation=cation, cation_charge=cation_charge, cation_count=cation_count,
                anion=anion, anion_charge=anion_charge, anion_count=anion_count,
                ratio=f"{cation_count} : {anion_count}", formula=formula, memo=memo, alias=list(alias))

# ---------------------------------------------------------------- 원소 기호 (원자 번호 순)
ELEMENTS = [
    element(1, "H", "수소", True, "140·145쪽"),
    element(2, "He", "헬륨", True, "140·145쪽"),
    element(3, "Li", "리튬", True, "140·145쪽"),
    element(4, "Be", "베릴륨", True, "140·145쪽"),
    element(5, "B", "붕소", True, "145쪽"),
    element(6, "C", "탄소", True, "140·145쪽"),
    element(7, "N", "질소", True, "145·160쪽"),
    element(8, "O", "산소", True, "144·145쪽"),
    element(9, "F", "플루오린", True, "145·154쪽"),
    element(10, "Ne", "네온", True, "140·145쪽"),
    element(11, "Na", "나트륨", True, "140·145쪽", alias=["소듐"]),
    element(12, "Mg", "마그네슘", True, "140·145쪽"),
    element(13, "Al", "알루미늄", True, "140·145쪽"),
    element(14, "Si", "규소", True, "140·145쪽"),
    element(15, "P", "인", True, "145쪽"),
    element(16, "S", "황", True, "140·145쪽"),
    element(17, "Cl", "염소", True, "140·145쪽"),
    element(18, "Ar", "아르곤", True, "145·147쪽"),
    element(19, "K", "칼륨", True, "145·157쪽", alias=["포타슘"]),
    element(20, "Ca", "칼슘", True, "140·145쪽"),
    element(21, "Sc", "스칸듐", False),
    element(22, "Ti", "타이타늄", False),
    element(23, "V", "바나듐", False),
    element(24, "Cr", "크로뮴", False),
    element(25, "Mn", "망가니즈", True),
    element(26, "Fe", "철", True, "140쪽"),
    element(27, "Co", "코발트", False),
    element(28, "Ni", "니켈", False),
    element(29, "Cu", "구리", True, "140·157쪽"),
    element(30, "Zn", "아연", True),
    element(35, "Br", "브로민", True),
    element(38, "Sr", "스트론튬", False),
    element(47, "Ag", "은", True),
    element(48, "Cd", "카드뮴", False),
    element(50, "Sn", "주석", False),
    element(53, "I", "아이오딘", True),
    element(56, "Ba", "바륨", False),
    element(78, "Pt", "백금", False),
    element(79, "Au", "금", True),
    element(80, "Hg", "수은", True, "140쪽"),
    element(82, "Pb", "납", True),
]

# ---------------------------------------------------------------- 분자식
MOLECULES = [
    molecule("H₂", "수소", True, "원소", "153·165쪽"),
    molecule("O₂", "산소", True, "원소", "160·163쪽"),
    molecule("O₃", "오존", True, "원소", "163쪽"),
    molecule("N₂", "질소", False, "원소"),
    molecule("H₂O", "물", True, "화합물", "141·153쪽"),
    molecule("CH₄", "메테인", True, "화합물", "141·153쪽"),
    molecule("NH₃", "암모니아", True, "화합물", "141·153쪽"),
    molecule("CO", "일산화 탄소", True, "화합물", "141·153쪽"),
    molecule("CO₂", "이산화 탄소", True, "화합물", "141·153쪽"),
    molecule("HCl", "염화 수소", True, "화합물", "153쪽"),
    molecule("H₂O₂", "과산화 수소", False, "화합물"),
    molecule("SO₂", "이산화 황", False, "화합물"),
    molecule("NO₂", "이산화 질소", False, "화합물"),
    molecule("C₆H₁₂O₆", "포도당", False, "화합물"),
]

# ---------------------------------------------------------------- 이온식 (양이온 → 음이온 순서로 사이트·PDF에 나옵니다)
CATIONS = [
    ion("H⁺", "수소 이온", True, "수소", "H", 1, 1, "교과서", "155쪽 표 Ⅳ-1", "표 Ⅳ-1 수록"),
    ion("Li⁺", "리튬 이온", True, "리튬", "Li", 3, 1, "교과서", "154쪽 그림 Ⅳ-11", "그림 Ⅳ-11 수록"),
    ion("Na⁺", "나트륨 이온", True, "나트륨", "Na", 11, 1, "교과서", "155쪽 그림 Ⅳ-12", "잃은 전자가 1개이므로 숫자 1을 생략"),
    ion("K⁺", "칼륨 이온", True, "칼륨", "K", 19, 1, "교과서", "155쪽 표 Ⅳ-1", "질산 칼륨·과망가니즈산 칼륨 수용액에 들어 있음"),
    ion("Be²⁺", "베릴륨 이온", True, "베릴륨", "Be", 4, 2, "교과서", "155쪽 표 Ⅳ-1", "표 Ⅳ-1 수록"),
    ion("Mg²⁺", "마그네슘 이온", True, "마그네슘", "Mg", 12, 2, "교과서", "155쪽 표 Ⅳ-1", "표 Ⅳ-1 수록"),
    ion("Ca²⁺", "칼슘 이온", True, "칼슘", "Ca", 20, 2, "교과서", "155쪽 표 Ⅳ-1", "표 Ⅳ-1 수록"),
    ion("Cu²⁺", "구리 이온", True, "구리", "Cu", 29, 2, "교과서", "155쪽 그림 Ⅳ-12", "잃은 전자 수 2를 전하 기호 앞에 씀 · 구리(Ⅱ) 이온", alias=["구리(Ⅱ) 이온"]),
    ion("NH₄⁺", "암모늄 이온", True, "-", "-", None, 1, "교과서", "155쪽 본문", "질소 1개와 수소 4개가 결합한 다원자 이온"),
    ion("Al³⁺", "알루미늄 이온", False, "알루미늄", "Al", 13, 3, "추가", "–", "전자 3개를 잃는 대표적인 양이온"),
    ion("Zn²⁺", "아연 이온", False, "아연", "Zn", 30, 2, "추가", "–", "금속과 산의 반응에서 자주 등장"),
    ion("Ag⁺", "은 이온", False, "은", "Ag", 47, 1, "추가", "–", "앙금 생성 반응에서 자주 등장"),
    ion("Ba²⁺", "바륨 이온", False, "바륨", "Ba", 56, 2, "추가", "–", "황산 이온과 흰색 앙금을 만듦"),
    ion("Fe²⁺", "철 이온", False, "철", "Fe", 26, 2, "추가", "–", "철은 Fe²⁺, Fe³⁺ 두 가지 이온을 만듦 · 철(Ⅱ) 이온", alias=["철(Ⅱ) 이온"]),
    ion("Pb²⁺", "납 이온", False, "납", "Pb", 82, 2, "추가", "–", "아이오딘화 이온과 노란색 앙금을 만듦"),
]
ANIONS = [
    ion("Cl⁻", "염화 이온", True, "염소", "Cl", 17, -1, "교과서", "155쪽 그림 Ⅳ-12", "'염소'는 '소'를 떼고 '염화 이온'이라 부름"),
    ion("O²⁻", "산화 이온", True, "산소", "O", 8, -2, "교과서", "155쪽 표 Ⅳ-1", "'산소'는 '소'를 떼고 '산화 이온'이라 부름"),
    ion("S²⁻", "황화 이온", True, "황", "S", 16, -2, "교과서", "155쪽 그림 Ⅳ-12", "그림 Ⅳ-12 수록"),
    ion("F⁻", "플루오린화 이온", True, "플루오린", "F", 9, -1, "교과서", "154쪽 그림 Ⅳ-11", "원소 이름 뒤에 '화 이온'을 붙임"),
    ion("OH⁻", "수산화 이온", True, "-", "-", None, -1, "교과서", "155쪽 본문", "산소 1개와 수소 1개가 결합한 다원자 이온"),
    ion("NO₃⁻", "질산 이온", True, "-", "-", None, -1, "교과서", "155쪽 표 Ⅳ-1", "질소 1개와 산소 3개가 결합한 다원자 이온"),
    ion("CO₃²⁻", "탄산 이온", True, "-", "-", None, -2, "교과서", "155쪽 표 Ⅳ-1", "탄소 1개와 산소 3개가 결합한 다원자 이온"),
    ion("SO₄²⁻", "황산 이온", True, "-", "-", None, -2, "교과서", "157쪽 그림 Ⅳ-13", "황산 구리(Ⅱ) 수용액에 들어 있음"),
    ion("MnO₄⁻", "과망가니즈산 이온", True, "-", "-", None, -1, "교과서", "155쪽 표 Ⅳ-1", "과망가니즈산 칼륨 수용액에 들어 있음(보라색)"),
    ion("Br⁻", "브로민화 이온", False, "브로민", "Br", 35, -1, "추가", "–", "원소 이름 뒤에 '화 이온'을 붙임"),
    ion("I⁻", "아이오딘화 이온", False, "아이오딘", "I", 53, -1, "추가", "–", "원소 이름 뒤에 '화 이온'을 붙임"),
]
IONS = CATIONS + ANIONS

# ---------------------------------------------------------------- 이온 결합 화합물 (모두 핵심)
COMPOUNDS = [
    compound("염화 나트륨", "Na⁺", 1, 1, "Cl⁻", -1, 1, "NaCl", "교과서 141쪽 · 가장 기본적인 이온 결합 화합물"),
    compound("염화 칼륨", "K⁺", 1, 1, "Cl⁻", -1, 1, "KCl", "가장 간단한 1 : 1"),
    compound("염화 마그네슘", "Mg²⁺", 2, 1, "Cl⁻", -1, 2, "MgCl₂", "음이온이 2개"),
    compound("황산 나트륨", "Na⁺", 1, 2, "SO₄²⁻", -2, 1, "Na₂SO₄", "양이온이 2개"),
    compound("수산화 나트륨", "Na⁺", 1, 1, "OH⁻", -1, 1, "NaOH", "다원자 이온이 1개이면 괄호 없이 씀"),
    compound("탄산 칼슘", "Ca²⁺", 2, 1, "CO₃²⁻", -2, 1, "CaCO₃", "전하의 크기가 같음"),
    compound("염화 구리(Ⅱ)", "Cu²⁺", 2, 1, "Cl⁻", -1, 2, "CuCl₂", "염화 이온이 2개", alias=["염화 구리"]),
    compound("수산화 칼슘", "Ca²⁺", 2, 1, "OH⁻", -1, 2, "Ca(OH)₂", "다원자 이온이 2개이므로 괄호 사용"),
    compound("과망가니즈산 칼륨", "K⁺", 1, 1, "MnO₄⁻", -1, 1, "KMnO₄", "MnO₄⁻는 과망가니즈산 이온"),
    compound("질산 칼륨", "K⁺", 1, 1, "NO₃⁻", -1, 1, "KNO₃", "질산 이온은 NO₃⁻"),
    compound("황산 구리", "Cu²⁺", 2, 1, "SO₄²⁻", -2, 1, "CuSO₄", "중학교에서는 '황산 구리(Ⅱ)'로도 표기", alias=["황산 구리(Ⅱ)"]),
]

# ---------------------------------------------------------------- 이온 이름 규칙 (엑셀)
RULES = [
    ("양이온은 원소 이름 뒤에 '~ 이온'을 붙인다.", "나트륨 → 나트륨 이온(Na⁺) / 구리 → 구리 이온(Cu²⁺)", "나트륨화 이온"),
    ("음이온은 원소 이름 뒤에 '~화 이온'을 붙인다.", "황 → 황화 이온(S²⁻) / 플루오린 → 플루오린화 이온(F⁻)", "황 이온"),
    ("원소 이름이 '소'로 끝나면 '소'를 떼고 '~화 이온'을 붙인다.", "염소 → 염화 이온(Cl⁻) / 산소 → 산화 이온(O²⁻)", "염소화 이온, 산소화 이온"),
    ("잃거나 얻은 전자가 1개일 때는 숫자 1을 생략한다.", "Na⁺, Cl⁻", "Na¹⁺, Cl¹⁻"),
    ("숫자는 전하의 종류 기호 앞에 쓴다.", "Cu²⁺, O²⁻", "Cu⁺², O⁻²"),
    ("여러 원자가 결합한 다원자 이온은 고유한 이름을 그대로 쓴다.", "NH₄⁺ 암모늄 이온 / OH⁻ 수산화 이온 / NO₃⁻ 질산 이온", "질소수소 이온"),
    ("이온이 되어도 양성자 수는 변하지 않으므로 원소의 종류도 그대로다.", "Na(양성자 11개) → Na⁺(양성자 11개)", "이온이 되면 다른 원소가 된다"),
    ("전류를 흘리면 양이온은 (−)극, 음이온은 (+)극으로 이동한다.", "Cu²⁺ → (−)극 / MnO₄⁻ → (+)극", "양이온이 (+)극으로 이동한다"),
]

# ---------------------------------------------------------------- 공용 도우미
SUB = "₀₁₂₃₄₅₆₇₈₉"
SUP = "⁰¹²³⁴⁵⁶⁷⁸⁹⁺⁻"

def plain(formula):
    """첨자를 보통 글자로: SO₄²⁻ → SO42- (엑셀 '입력용 표기')"""
    table = {c: str(i) for i, c in enumerate(SUB)}
    table.update({c: str(i) for i, c in enumerate(SUP[:10])})
    table.update({"⁺": "+", "⁻": "-"})
    return "".join(table.get(c, c) for c in formula)

def formula_html(text):
    """유니코드 첨자 → <sub>/<sup> HTML (연속된 첨자는 하나로 묶음)"""
    out, mode, buf = [], None, ""
    def flush():
        nonlocal buf
        if buf:
            out.append({"sub": "<sub>%s</sub>", "sup": "<sup>%s</sup>"}.get(mode, "%s") % buf)
        buf = ""
    for ch in text:
        if ch in SUB:
            m, p = "sub", str(SUB.index(ch))
        elif ch in SUP:
            i = SUP.index(ch); m, p = "sup", (str(i) if i < 10 else ("+" if i == 10 else "−"))
        else:
            m, p = None, ch
        if m != mode:
            flush(); mode = m
        buf += p
    flush()
    return "".join(out)

def is_polyatomic(formula):
    """대문자(원소 기호 시작)가 2개 이상이면 다원자 이온"""
    return sum(1 for ch in formula if ch.isupper()) >= 2
