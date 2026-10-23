"""사전·사후 값만 입력하는 간단 양식 (VO2max, FMS 총점, FMS 하위항목) + 논문 표용 자동 결과.

사용: python3 build_simple.py <출력.xlsx> [--sample]
"""
import sys
import random
from openpyxl import Workbook
from openpyxl.styles import Font, PatternFill, Alignment, Border, Side
from openpyxl.worksheet.datavalidation import DataValidation
from openpyxl.worksheet.formula import ArrayFormula
from openpyxl.formatting.rule import FormulaRule
from openpyxl.utils import get_column_letter as L

OUT = sys.argv[1]
SAMPLE = "--sample" in sys.argv

FONT = "맑은 고딕"
F = lambda **k: Font(name=FONT, size=k.pop("size", 10), **k)
FILL_IN = PatternFill("solid", fgColor="FFF2CC")
FILL_CALC = PatternFill("solid", fgColor="E7E6E6")
FILL_HDR = PatternFill("solid", fgColor="1F4E78")
FILL_SUB = PatternFill("solid", fgColor="D9E1F2")
FILL_EX = PatternFill("solid", fgColor="F2F2F2")
FILL_WARN = PatternFill("solid", fgColor="F8CBAD")
THIN = Side(style="thin", color="BFBFBF")
BOX = Border(left=THIN, right=THIN, top=THIN, bottom=THIN)
C = Alignment(horizontal="center", vertical="center", wrap_text=True)

EX, FIRST, LAST = 4, 5, 34          # 예시 행, 입력 행(30명분)
S_IN, S_OUT = "입력", "결과"
REF = f"'{S_IN}'"

VARS = [  # (코드, 표시 이름, 단위/범위, 정수 여부)
    ("VO2", "VO2max (mL/kg/min)", None, False),
    ("FMS", "FMS 총점 (0–21)", (0, 21), True),
    ("DS", "Deep Squat", (0, 3), True),
    ("HS", "Hurdle Step", (0, 3), True),
    ("ILL", "In-line Lunge", (0, 3), True),
    ("SM", "Shoulder Mobility", (0, 3), True),
    ("ASLR", "Active Straight Leg Raise", (0, 3), True),
    ("TSPU", "Trunk Stability Push-up", (0, 3), True),
    ("RS", "Rotary Stability", (0, 3), True),
]
COL = {}  # (코드, 'pre'/'post') -> 열 문자
c = 3
for code, *_ in VARS:
    COL[(code, "pre")], COL[(code, "post")] = L(c), L(c + 1)
    c += 2
SUM_PRE, SUM_POST, CHK = L(c), L(c + 1), L(c + 2)

wb = Workbook()
wi = wb.active
wi.title = S_IN
wo = wb.create_sheet(S_OUT)

# ------------------------------------------------------------------ 입력 시트
wi["A1"] = ("노란 칸에 사전·사후 값만 입력하세요. 집단: 1=실험군, 2=통제군. 측정하지 않은 값은 빈칸(0 아님). "
            "FMS 하위항목은 최종 항목 점수(좌우 중 낮은 점수, 통증 시 0점)를 입력합니다. 하위항목이 없으면 비워 두셔도 됩니다. "
            "결과는 '결과' 시트에 자동으로 계산됩니다.")
wi["A1"].font = F(size=9, color="595959")
wi["A1"].alignment = Alignment(wrap_text=True, vertical="center")
wi.merge_cells(f"A1:{CHK}1")
wi.row_dimensions[1].height = 34

for col, lab in (("A", "ID"), ("B", "집단\n1=실험\n2=통제")):
    wi.merge_cells(f"{col}2:{col}3")
    cell = wi[f"{col}2"]
    cell.value, cell.font, cell.fill, cell.alignment = lab, F(bold=True, color="FFFFFF"), FILL_HDR, C
for code, name, *_ in VARS:
    a, b = COL[(code, "pre")], COL[(code, "post")]
    wi.merge_cells(f"{a}2:{b}2")
    wi[f"{a}2"].value = name
    for x in (a, b):
        wi[f"{x}2"].font, wi[f"{x}2"].fill, wi[f"{x}2"].alignment = F(bold=True, color="FFFFFF"), FILL_HDR, C
    wi[f"{a}3"], wi[f"{b}3"] = "사전", "사후"
    for x in (a, b):
        wi[f"{x}3"].font, wi[f"{x}3"].fill, wi[f"{x}3"].alignment = F(bold=True), FILL_SUB, C
        wi.column_dimensions[x].width = 9 if code != "VO2" else 10
wi.merge_cells(f"{SUM_PRE}2:{SUM_POST}2")
wi[f"{SUM_PRE}2"] = "하위항목 합 (자동)"
for x in (SUM_PRE, SUM_POST, CHK):
    wi[f"{x}2"].font, wi[f"{x}2"].fill, wi[f"{x}2"].alignment = F(bold=True), FILL_CALC, C
wi[f"{SUM_PRE}3"], wi[f"{SUM_POST}3"] = "사전", "사후"
wi[f"{CHK}2"] = "자동 점검"
wi.merge_cells(f"{CHK}2:{CHK}3")
for x in (SUM_PRE, SUM_POST):
    wi[f"{x}3"].font, wi[f"{x}3"].fill, wi[f"{x}3"].alignment = F(bold=True), FILL_CALC, C
    wi.column_dimensions[x].width = 8
wi.column_dimensions["A"].width = 8
wi.column_dimensions["B"].width = 9
wi.column_dimensions[CHK].width = 42
wi.row_dimensions[2].height = 40
for r in (2, 3):
    for j in range(1, c + 3):
        wi.cell(row=r, column=j).border = BOX

item_codes = [code for code, *_ in VARS[2:]]


def row_formulas(r):
    pre_items = ",".join(f"{COL[(k, 'pre')]}{r}" for k in item_codes)
    post_items = ",".join(f"{COL[(k, 'post')]}{r}" for k in item_codes)
    fp, fq = f"{COL[('FMS', 'pre')]}{r}", f"{COL[('FMS', 'post')]}{r}"
    sp, sq = f"{SUM_PRE}{r}", f"{SUM_POST}{r}"
    parts = [
        f'IF(AND(ISNUMBER({sp}),ISNUMBER({fp})),IF({sp}<>{fp},"사전: FMS 총점≠하위항목 합 ",""),"")',
        f'IF(AND(ISNUMBER({sq}),ISNUMBER({fq})),IF({sq}<>{fq},"사후: FMS 총점≠하위항목 합 ",""),"")',
        f'IF(ISNUMBER({fp}),IF({fp}<>INT({fp}),"사전 총점 정수 아님 ",""),"")',
        f'IF(ISNUMBER({fq}),IF({fq}<>INT({fq}),"사후 총점 정수 아님 ",""),"")',
        f'IF(AND(COUNT({pre_items})>0,COUNT({pre_items})<7),"사전 하위항목 일부 빈칸 ","")',
        f'IF(AND(COUNT({post_items})>0,COUNT({post_items})<7),"사후 하위항목 일부 빈칸 ","")',
    ]
    return (f'=IF(COUNT({pre_items})=7,SUM({pre_items}),"")',
            f'=IF(COUNT({post_items})=7,SUM({post_items}),"")',
            f'=IF(A{r}="","",TRIM({"&".join(parts)}))')


for r in [EX] + list(range(FIRST, LAST + 1)):
    ex = r == EX
    for j in range(1, c):
        cell = wi.cell(row=r, column=j)
        cell.border, cell.alignment = BOX, Alignment(horizontal="center", vertical="center")
        cell.fill = FILL_EX if ex else FILL_IN
        cell.font = F(italic=True, color="7F7F7F") if ex else F(color="0000FF")
        if j >= 3:
            cell.number_format = "0.00" if j <= 4 else "0"
    a, b, chk = row_formulas(r)
    for col, v, fmt in ((SUM_PRE, a, "0"), (SUM_POST, b, "0"), (CHK, chk, "@")):
        cell = wi[f"{col}{r}"]
        cell.value, cell.fill, cell.border, cell.number_format = v, FILL_CALC, BOX, fmt
        cell.font = F(italic=ex, color="7F7F7F" if ex else "000000")
        cell.alignment = Alignment(horizontal="left" if col == CHK else "center", vertical="center")
# 예시 행 (집단을 비워 분석에서 제외)
ex_vals = {"A": "예시", ("VO2", "pre"): 48.30, ("VO2", "post"): 52.40, ("FMS", "pre"): 14, ("FMS", "post"): 16,
           ("DS", "pre"): 2, ("DS", "post"): 3, ("HS", "pre"): 2, ("HS", "post"): 2, ("ILL", "pre"): 2,
           ("ILL", "post"): 3, ("SM", "pre"): 2, ("SM", "post"): 2, ("ASLR", "pre"): 2, ("ASLR", "post"): 2,
           ("TSPU", "pre"): 2, ("TSPU", "post"): 2, ("RS", "pre"): 2, ("RS", "post"): 2}
for k, v in ex_vals.items():
    wi[f"{k}{EX}" if k == "A" else f"{COL[k]}{EX}"] = v
wi[f"B{EX}"] = "(예시)"
# 대상자 ID 미리 채움
ids = [f"E{k:02d}" for k in range(1, 14)] + [f"C{k:02d}" for k in range(1, 13)]
for k, pid in enumerate(ids):
    wi[f"A{FIRST + k}"], wi[f"B{FIRST + k}"] = pid, (1 if pid.startswith("E") else 2)

dv = DataValidation(type="list", formula1='"1,2"', allow_blank=True, showErrorMessage=True,
                    errorTitle="집단", error="1(실험군) 또는 2(통제군)")
dv.add(f"B{FIRST}:B{LAST}")
wi.add_data_validation(dv)
for code, name, rng, integer in VARS:
    lo, hi = rng if rng else (15, 95)
    v = DataValidation(type="decimal", operator="between", formula1=str(lo), formula2=str(hi), allow_blank=True,
                       showErrorMessage=True, errorStyle="warning", errorTitle=code[:30],
                       error=f"{name}: {lo}–{hi} 범위인지 확인하세요." + (" FMS 점수는 보통 정수입니다." if integer else ""))
    v.add(f"{COL[(code, 'pre')]}{FIRST}:{COL[(code, 'post')]}{LAST}")
    wi.add_data_validation(v)
wi.conditional_formatting.add(f"{CHK}{FIRST}:{CHK}{LAST}", FormulaRule(formula=[f"LEN({CHK}{FIRST})>0"], fill=FILL_WARN))
wi.freeze_panes = f"C{FIRST}"

# ------------------------------------------------------------------ 결과 시트
wo["A1"] = "결과 (논문 표 3·표 4용) — 입력 시트에서 집단이 1/2이고 사전·사후가 모두 있는 대상자만 계산"
wo["A1"].font = F(size=12, bold=True)
wo["A2"] = ("t는 SPSS 대응표본 검정과 같은 '사전−사후' 기준(증가하면 음수). 상호작용 F는 2(집단)×2(시기) 반복측정 분산분석의 "
            "집단×시기 상호작용과 같은 값(변화량의 독립표본 t²)입니다. FMS 하위항목의 p는 대응표본 t검정 기준입니다.")
wo["A2"].font = F(size=9, color="595959")
wo["A2"].alignment = Alignment(wrap_text=True, vertical="center")
wo.merge_cells("A2:P2")
wo.row_dimensions[2].height = 30

GRP = f"{REF}!$B${FIRST}:$B${LAST}"


def rng(code, t):
    col = COL[(code, t)]
    return f"{REF}!${col}${FIRST}:${col}${LAST}"


def hdr(ws, r, labels):
    for j, h in enumerate(labels, start=1):
        cell = ws.cell(row=r, column=j, value=h)
        cell.font, cell.fill, cell.alignment, cell.border = F(bold=True, color="FFFFFF", size=9), FILL_HDR, C, BOX
    ws.row_dimensions[r].height = 30


def stat_row(ws, r, code, g):
    pre, post = rng(code, "pre"), rng(code, "post")
    cnd = f"({GRP}={g})*ISNUMBER({pre})*ISNUMBER({post})"
    ws[f"C{r}"] = f"=SUMPRODUCT({cnd})"
    for col, fn, src, nmin in (("D", "AVERAGE", pre, 1), ("E", "STDEV", pre, 2), ("F", "AVERAGE", post, 1),
                               ("G", "STDEV", post, 2)):
        ws[f"{col}{r}"] = ArrayFormula(f"{col}{r}", f'=IF(C{r}<{nmin},"",{fn}(IF({cnd},{src})))')
    ws[f"H{r}"] = ArrayFormula(f"H{r}", f'=IF(C{r}<2,"",STDEV(IF({cnd},{post}-{pre})))')   # 변화량 SD
    ws[f"I{r}"] = f'=IF(OR(C{r}<2,H{r}=""),"",IF(ROUND(H{r},10)=0,"",(D{r}-F{r})/(H{r}/SQRT(C{r}))))'   # t (사전−사후)
    ws[f"J{r}"] = f'=IF(I{r}="","",C{r}-1)'
    ws[f"K{r}"] = f'=IF(I{r}="","",TDIST(ABS(I{r}),J{r},2))'
    # 논문 표기
    ws[f"L{r}"] = f'=IF(D{r}="","",TEXT(D{r},"0.00")&"±"&IF(E{r}="","",TEXT(E{r},"0.00")))'
    ws[f"M{r}"] = f'=IF(F{r}="","",TEXT(F{r},"0.00")&"±"&IF(G{r}="","",TEXT(G{r},"0.00")))'
    ws[f"N{r}"] = f'=IF(K{r}="","",TEXT(I{r},"0.000")&" ("&IF(K{r}<0.001,"<.001",TEXT(K{r},".000"))&")")'
    fmts = {"C": "0", "D": "0.00", "E": "0.00", "F": "0.00", "G": "0.00", "H": "0.00", "I": "0.000", "J": "0",
            "K": "0.0000"}
    for j in range(1, 15):
        cell = ws.cell(row=r, column=j)
        cell.border, cell.alignment = BOX, Alignment(horizontal="center", vertical="center")
        if j >= 3:
            cell.fill, cell.font = FILL_CALC, F(size=9)
            cell.number_format = fmts.get(L(j), "@")


def inter_cells(ws, r, re_, rc):
    """r행 O–P 열에 상호작용 F(p)와 부분 η² (실험군 행 re_, 통제군 행 rc)."""
    d1, d2 = f"(F{re_}-D{re_})", f"(F{rc}-D{rc})"
    n1, n2, s1, s2 = f"C{re_}", f"C{rc}", f"H{re_}", f"H{rc}"
    sp = f"SQRT((({n1}-1)*{s1}^2+({n2}-1)*{s2}^2)/({n1}+{n2}-2))"
    ok = f'OR({n1}<2,{n2}<2,{s1}="",{s2}="")'
    ws[f"O{r}"] = f'=IF({ok},"",IF(ROUND({sp},10)=0,"",(({d1}-{d2})/({sp}*SQRT(1/{n1}+1/{n2})))^2))'
    ws[f"P{r}"] = f'=IF(O{r}="","",{n1}+{n2}-2)'
    ws[f"Q{r}"] = f'=IF(O{r}="","",FDIST(O{r},1,P{r}))'
    ws[f"R{r}"] = f'=IF(O{r}="","",O{r}/(O{r}+P{r}))'
    ws[f"S{r}"] = (f'=IF(O{r}="","","F(1, "&P{r}&") = "&TEXT(O{r},"0.000")&", p "&IF(Q{r}<0.001,"< .001","= "&TEXT(Q{r},".000"))'
                   f'&", ηp² = "&TEXT(R{r},".000"))')
    for col, fmt in (("O", "0.000"), ("P", "0"), ("Q", "0.0000"), ("R", "0.000"), ("S", "@")):
        cell = ws[f"{col}{r}"]
        cell.fill, cell.font, cell.border, cell.number_format = FILL_CALC, F(size=9), BOX, fmt
        cell.alignment = Alignment(horizontal="left" if col == "S" else "center", vertical="center")


HEAD = ["변인", "집단", "n", "사전 M", "사전 SD", "사후 M", "사후 SD", "변화량 SD", "t\n(사전−사후)", "df", "p",
        "사전\n(논문 표기)", "사후\n(논문 표기)", "t (p)\n(논문 표기)", "상호작용 F", "df(오차)", "p", "부분 η²",
        "상호작용 (논문 표기)"]

wo["A4"] = "표 3. VO2max와 FMS 총점"
wo["A4"].font = F(bold=True, size=11)
hdr(wo, 5, HEAD)
r = 6
for code, name in (("VO2", "VO2max"), ("FMS", "FMS 총점")):
    re_, rc = r, r + 1
    for g, gname in ((1, "실험군"), (2, "통제군")):
        wo[f"A{r}"], wo[f"B{r}"] = (name if g == 1 else ""), gname
        stat_row(wo, r, code, g)
        r += 1
    inter_cells(wo, re_, re_, rc)
    for col in "OPQRS":
        wo.merge_cells(f"{col}{re_}:{col}{rc}")

r += 1
wo[f"A{r}"] = "표 4. FMS 하위항목"
wo[f"A{r}"].font = F(bold=True, size=11)
r += 1
hdr(wo, r, HEAD[:14])
r += 1
for code, name, *_ in VARS[2:]:
    for g, gname in ((1, "실험군"), (2, "통제군")):
        wo[f"A{r}"], wo[f"B{r}"] = (name if g == 1 else ""), gname
        stat_row(wo, r, code, g)
        r += 1
wo[f"A{r}"] = ("FMS 하위항목은 0–3점 서열척도이므로 논문에서는 Wilcoxon 부호순위검정을 권합니다. "
               "이 시트의 p는 참고용 대응표본 t검정 값입니다.")
wo[f"A{r}"].font = F(size=9, color="595959")

widths = {"A": 22, "B": 8, "C": 5, "D": 8, "E": 8, "F": 8, "G": 8, "H": 9, "I": 9, "J": 5, "K": 8,
          "L": 13, "M": 13, "N": 15, "O": 10, "P": 7, "Q": 8, "R": 8, "S": 36}
for k, w in widths.items():
    wo.column_dimensions[k].width = w
wo.freeze_panes = "C6"

for ws in (wi, wo):
    ws.sheet_view.zoomScale = 90

# ------------------------------------------------------------------ 검산용 가상자료
if SAMPLE:
    rnd = random.Random(7)
    for k, pid in enumerate(ids):
        rr = FIRST + k
        g = 1 if pid.startswith("E") else 2
        vp = round(rnd.gauss(48.3, 4), 2)
        wi[f"{COL[('VO2', 'pre')]}{rr}"] = vp
        if pid != "C12":
            wi[f"{COL[('VO2', 'post')]}{rr}"] = round(vp + (rnd.gauss(4, 2.3) if g == 1 else rnd.gauss(0.5, 1.5)), 2)
        for t in ("pre", "post"):
            items = []
            for code in item_codes:
                v = rnd.choice([1, 2, 2, 2, 3])
                if t == "post" and g == 1 and v < 3 and rnd.random() < 0.4:
                    v += 1
                items.append(v)
                if pid != "E07":   # 하위항목 없는 대상자
                    wi[f"{COL[(code, t)]}{rr}"] = v
            tot = sum(items) + (1 if (pid == "E03" and t == "pre") else 0)   # 총점 불일치 1건
            wi[f"{COL[('FMS', t)]}{rr}"] = tot

wb.save(OUT)
print("saved", OUT)
