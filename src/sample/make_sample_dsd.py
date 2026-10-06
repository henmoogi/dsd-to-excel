# 공개용 예시 DSD 생성기: 가상 회사 '샘플전자(주)' 2025년 감사보고서.
# 회사·감사인·은행·주주 이름과 모든 숫자·문장은 지어낸 것이다(실제 회사 자료를 쓰지 않음).
# 검증 기능 시연용으로 주석 6 유형자산 변동표의 '합계' 행 감가상각비를 일부러 틀리게 적는다(ERR 표시).
import zipfile, sys, os

OUT = sys.argv[1] if len(sys.argv) > 1 else "샘플전자_감사보고서_2025_예시.dsd"
WITH_ERROR = "--no-error" not in sys.argv

def esc(t):
    return t.replace("&", "&amp;").replace("<", "&lt;").replace(">", "&gt;")

def won(v):
    if v is None:
        return "　"
    if v == 0:
        return "-"
    s = f"{abs(v):,}"
    return f"({s})" if v < 0 else s

M = 1_000_000  # 숫자는 백만원 단위로 적고 원으로 바꾼다

def m(v):
    return None if v is None else int(round(v * M))

out = []
w = out.append

# ---------- 머리 ----------
w('<?xml version="1.0" encoding="utf-8"?>\n')
w('<DOCUMENT xmlns:xsi="http://www.w3.org/2001/XMLSchema-instance" xsi:noNamespaceSchemaLocation="dart4.xsd">\n')
w('<DOCUMENT-HEADER AEXT-CLASS="Y">\n<DOCUMENT-NAME ACODE="00760">감사보고서</DOCUMENT-NAME>\n')
w('<FORMULA-VERSION ADATE="20221027">1.0</FORMULA-VERSION>\n<COMPANY-NAME AREGCIK="00000000" AACCOUNTTYPE="A">샘플전자(주)</COMPANY-NAME>\n')
w('<SUMMARY>\n')
for k, v in [("TOT_ASSETS", "9481"), ("TOT_DEBTS", "3397"), ("TOT_SALES", "18500"), ("TOT_EMPL", "25")]:
    w(f'<EXTRACTION ACODE="{k}" AFEATURE="BOTH">{v}</EXTRACTION>\n')
w('</SUMMARY>\n</DOCUMENT-HEADER>\n<BODY ATOCID="9">\n')

def P(text="", mark=None):
    attr = f' USERMARK="{mark}"' if mark else ""
    w(f"<P{attr}>{text}</P>\n")

CR = "&amp;cr;"

def simple_table(rows, width=600, border=0, align="CENTER"):
    w(f'<TABLE ACLASS="NORMAL" AFIXTABLE="N" WIDTH="{width}" BORDER="{border}">\n<COLGROUP WIDTH="{width}">\n<COL WIDTH="{width}"></COL>\n</COLGROUP>\n<TBODY>\n')
    for r in rows:
        w(f'<TR ACOPY="Y" ADELETE="Y" HEIGHT="30">\n<TD ALIGN="{align}" WIDTH="{width - 9}" HEIGHT="23" VALIGN="MIDDLE">{esc(r)}</TD>\n</TR>\n')
    w("</TBODY>\n</TABLE>\n")

# ---------- 표지 ----------
w("<COVER>\n")
P(CR)
P("샘플전자 주식회사", "F-18 ")
P(CR + CR)
P("재 무 제 표 에 대 한", "F-18 ")
w('<COVER-TITLE ATOC="Y" ATOCID="1" ENG="Audit Report">감   사   보   고   서</COVER-TITLE>\n')
P(CR + CR)
simple_table(["제 10 기", "2025년 01월 01일 부터", "2025년 12월 31일 까지", "", "제 9 기", "2024년 01월 01일 부터", "2024년 12월 31일 까지"])
P(CR + CR)
P("샘플회계법인", "F-14 ")
w("</COVER>\n")

# ---------- 감사보고서 ----------
w('<SECTION-1 ACLASS="MANDATORY" APARTSOURCE="SOURCE">\n')
w('<TITLE ATOC="Y" AASSOCNOTE="D-0-0-1-0" ATOCID="2" ENG="Independent Auditor\'s Report">독립된 감사인의 감사보고서</TITLE>\n')
P("샘플전자 주식회사")
P("주주 및 이사회 귀중")
P(CR)
P("감사의견", "B")
P("우리는 샘플전자 주식회사(이하 '회사')의 재무제표를 감사하였습니다. 해당 재무제표는 2025년 12월 31일 현재의 재무상태표, 동일로 종료되는 보고기간의 손익계산서, 자본변동표, 현금흐름표 그리고 중요한 회계정책 정보를 포함한 재무제표의 주석으로 구성되어 있습니다.")
P("우리의 의견으로는 별첨된 회사의 재무제표는 회사의 2025년 12월 31일 현재의 재무상태와 동일로 종료되는 보고기간의 재무성과 및 현금흐름을 일반기업회계기준에 따라, 중요성의 관점에서 공정하게 표시하고 있습니다.")
P(CR)
P("감사의견근거", "B")
P("우리는 대한민국의 회계감사기준에 따라 감사를 수행하였습니다. 우리는 재무제표 감사와 관련된 대한민국의 윤리적 요구사항에 따라 회사로부터 독립적이며, 그러한 요구사항에 따른 기타의 윤리적 책임들을 이행하였습니다. 우리가 입수한 감사증거가 감사의견을 위한 근거로서 충분하고 적합하다고 우리는 믿습니다.")
P(CR)
P("이 감사보고서는 파일 변환 도구를 시연하기 위해 만든 가상의 예시입니다. 회사명, 감사인, 금액은 모두 실제와 관계가 없습니다.", "B")
P(CR)
P("샘플회계법인")
P("대표이사  홍 길 동")
P("2026년 3월 10일")
w("</SECTION-1>\n")

# ---------- 첨부 재무제표 표지 ----------
w('<SECTION-1 ACLASS="MANDATORY" APARTSOURCE="SOURCE">\n')
w('<TITLE ATOC="Y" AASSOCNOTE="D-0-0-0-0" ATOCID="3" ENG="(Attachment) Financial Statements">(첨부)재 무 제 표</TITLE>\n')
w('<TABLE-GROUP ACLASS="COVER2" ADELETETABLE="N">\n')
simple_table(["샘플전자 주식회사", "제 10 기", "2025년 01월 01일 부터", "2025년 12월 31일 까지", "", "제 9 기", "2024년 01월 01일 부터", "2024년 12월 31일 까지"])
P(CR)
P('"첨부된 재무제표는 당사가 작성한 것입니다."')
P("샘플전자 주식회사 대표이사 김 샘 플")
w("</TABLE-GROUP>\n")

# ---------- 재무제표 공통: 4열(당기 세부·합계, 전기 세부·합계) ----------
def fin_head(gclass, title_id, title, eng, period_lines, unit="(단위 : 원)"):
    w(f'<TABLE-GROUP ACLASS="{gclass}" ADELETETABLE="N">\n')
    w(f'<TITLE ATOC="Y" ATOCID="{title_id}" ENG="{eng}">{title}</TITLE>\n')
    w('<TABLE ACLASS="EXTRACTION" AFIXTABLE="Y" WIDTH="600" BORDER="0">\n<COLGROUP WIDTH="600">\n<COL WIDTH="300"></COL>\n<COL WIDTH="300"></COL>\n</COLGROUP>\n<TBODY>\n')
    for line in period_lines:
        w(f'<TR ACOPY="Y" ADELETE="Y" HEIGHT="30">\n<TD COLSPAN="2" ALIGN="CENTER" VALIGN="MIDDLE" WIDTH="591" HEIGHT="23">{line}</TD>\n</TR>\n')
    w(f'<TR ACOPY="Y" ADELETE="Y" HEIGHT="30">\n<TD VALIGN="BOTTOM" WIDTH="291" HEIGHT="23">샘플전자 주식회사</TD>\n<TU ALIGN="RIGHT" VALIGN="BOTTOM" WIDTH="291" HEIGHT="23" AUNIT="WON" AUNITVALUE="1" ENG="(Unit: KRW)">{unit}</TU>\n</TR>\n</TBODY>\n</TABLE>\n')

def fin_table(header, rows):
    # rows: (이름, 깊이, 당기세부, 당기합계, 전기세부, 전기합계)  금액은 백만원
    w('<TABLE ACLASS="FINANCE" AFIXTABLE="N" WIDTH="725" BORDER="1" FRAME="BORDER" RULES="ALL">\n<COLGROUP WIDTH="725">\n')
    w('<COL WIDTH="217"></COL>\n' + '<COL WIDTH="127"></COL>\n' * 4 + '</COLGROUP>\n<THEAD>\n<TR ACOPY="N" ADELETE="N" HEIGHT="30">\n')
    w(f'<TH VALIGN="MIDDLE" WIDTH="208" HEIGHT="23">{header[0]}</TH>\n')
    w(f'<TH COLSPAN="2" VALIGN="MIDDLE" WIDTH="245" HEIGHT="23">{header[1]}</TH>\n')
    w(f'<TH COLSPAN="2" VALIGN="MIDDLE" WIDTH="245" HEIGHT="23">{header[2]}</TH>\n</TR>\n</THEAD>\n<TBODY>\n')
    for name, lvl, a1, a2, b1, b2 in rows:
        mark = {0: "", 1: ' USERMARK="P16"', 2: ' USERMARK="P32"'}[lvl]
        w('<TR ACOPY="Y" ADELETE="Y" HEIGHT="30">\n')
        w(f'<TE ADELIM="0" ALEVEL="{lvl}"{mark} WIDTH="208" HEIGHT="23">{esc(name)}</TE>\n')
        for i, v in enumerate([a1, a2, b1, b2], 1):
            w(f'<TE ALIGN="RIGHT" ADELIM="{i}" WIDTH="118" HEIGHT="23">{won(m(v))}</TE>\n')
        w("</TR>\n")
    w("</TBODY>\n</TABLE>\n</TABLE-GROUP>\n")

PER_BS = ["제 10(당)기 2025년 12월 31일 현재", "제 9(전)기 2024년 12월 31일 현재"]
PER_PL = ["제 10(당)기 2025년 01월 01일부터 2025년 12월 31일까지", "제 9(전)기 2024년 01월 01일부터 2024년 12월 31일까지"]
HDR = ["과                        목", "제 10(당)기", "제 9(전)기"]

# 재무상태표
fin_head("PBS-2:A", 4, "재 무 상 태 표", "Statement of Financial Position", PER_BS)
fin_table(HDR, [
    ("자산", 0, None, None, None, None),
    ("Ⅰ.유동자산", 0, None, 5250, None, 4548),
    ("(1)당좌자산", 1, None, 4230, None, 3548),
    ("현금및현금성자산(주석3)", 2, 1250, None, 980, None),
    ("단기금융상품", 2, 500, None, 400, None),
    ("매출채권(주석4)", 2, 2400, None, 2100, None),
    ("대손충당금", 2, -24, None, -21, None),
    ("미수금", 2, 86, None, 74, None),
    ("선급비용", 2, 18, None, 15, None),
    ("(2)재고자산(주석5)", 1, None, 1020, None, 1000),
    ("제품", 2, 640, None, 590, None),
    ("원재료", 2, 380, None, 410, None),
    ("Ⅱ.비유동자산", 0, None, 4231, None, 4234),
    ("(1)투자자산", 1, None, 300, None, 300),
    ("장기금융상품", 2, 300, None, 300, None),
    ("(2)유형자산(주석6)", 1, None, 3766, None, 3754),
    ("토지", 2, 1500, None, 1500, None),
    ("건물", 2, 2000, None, 2000, None),
    ("감가상각누계액", 2, -450, None, -400, None),
    ("기계장치", 2, 1200, None, 1000, None),
    ("감가상각누계액", 2, -520, None, -400, None),
    ("차량운반구", 2, 90, None, 90, None),
    ("감가상각누계액", 2, -54, None, -36, None),
    ("(3)무형자산", 1, None, 45, None, 60),
    ("소프트웨어", 2, 45, None, 60, None),
    ("(4)기타비유동자산", 1, None, 120, None, 120),
    ("보증금", 2, 120, None, 120, None),
    ("자산총계", 0, None, 9481, None, 8782),
    ("부채", 0, None, None, None, None),
    ("Ⅰ.유동부채", 0, None, 2287, None, 2368),
    ("매입채무", 2, 1150, None, 1080, None),
    ("미지급금", 2, 210, None, 190, None),
    ("단기차입금(주석7)", 2, 800, None, 1000, None),
    ("미지급법인세", 2, 95, None, 70, None),
    ("예수금", 2, 32, None, 28, None),
    ("Ⅱ.비유동부채", 0, None, 1110, None, 1110),
    ("장기차입금(주석7)", 2, 1000, None, 1000, None),
    ("퇴직급여충당부채", 2, 410, None, 360, None),
    ("퇴직연금운용자산", 2, -300, None, -250, None),
    ("부채총계", 0, None, 3397, None, 3478),
    ("자본", 0, None, None, None, None),
    ("Ⅰ.자본금(주석8)", 0, None, 1000, None, 1000),
    ("보통주자본금", 2, 1000, None, 1000, None),
    ("Ⅱ.자본잉여금", 0, None, 500, None, 500),
    ("주식발행초과금", 2, 500, None, 500, None),
    ("Ⅲ.이익잉여금", 0, None, 4584, None, 3804),
    ("이익준비금", 2, 60, None, 50, None),
    ("미처분이익잉여금", 2, 4524, None, 3754, None),
    ("자본총계", 0, None, 6084, None, 5304),
    ("부채와자본총계", 0, None, 9481, None, 8782),
])

# 손익계산서
fin_head("PIS-2:A", 5, "손 익 계 산 서", "Income Statement", PER_PL)
fin_table(HDR, [
    ("Ⅰ.매출액", 0, None, 18500, None, 16800),
    ("Ⅱ.매출원가", 0, None, 14060, None, 12900),
    ("Ⅲ.매출총이익", 0, None, 4440, None, 3900),
    ("Ⅳ.판매비와관리비", 0, None, 3320, None, 3050),
    ("급여", 2, 1620, None, 1500, None),
    ("퇴직급여", 2, 130, None, 120, None),
    ("복리후생비", 2, 210, None, 190, None),
    ("지급수수료", 2, 480, None, 450, None),
    ("감가상각비", 2, 85, None, 80, None),
    ("무형자산상각비", 2, 15, None, 15, None),
    ("기타판매비와관리비", 2, 780, None, 695, None),
    ("Ⅴ.영업이익", 0, None, 1120, None, 850),
    ("Ⅵ.영업외수익", 0, None, 48, None, 36),
    ("이자수익", 2, 48, None, 36, None),
    ("Ⅶ.영업외비용", 0, None, 98, None, 112),
    ("이자비용", 2, 98, None, 112, None),
    ("Ⅷ.법인세비용차감전순이익", 0, None, 1070, None, 774),
    ("Ⅸ.법인세비용(주석9)", 0, None, 190, None, 134),
    ("Ⅹ.당기순이익", 0, None, 880, None, 640),
])

# 자본변동표
w('<TABLE-GROUP ACLASS="PEF-1" ADELETETABLE="N">\n<TITLE ATOC="Y" ATOCID="6" ENG="Statement of Changes in Equity">자 본 변 동 표</TITLE>\n')
w('<TABLE ACLASS="EXTRACTION" AFIXTABLE="Y" WIDTH="600" BORDER="0">\n<COLGROUP WIDTH="600">\n<COL WIDTH="300"></COL>\n<COL WIDTH="300"></COL>\n</COLGROUP>\n<TBODY>\n')
for line in PER_PL:
    w(f'<TR ACOPY="Y" ADELETE="Y" HEIGHT="30">\n<TD COLSPAN="2" ALIGN="CENTER" VALIGN="MIDDLE" WIDTH="591" HEIGHT="23">{line}</TD>\n</TR>\n')
w('<TR ACOPY="Y" ADELETE="Y" HEIGHT="30">\n<TD VALIGN="BOTTOM" WIDTH="291" HEIGHT="23">샘플전자 주식회사</TD>\n<TU ALIGN="RIGHT" VALIGN="BOTTOM" WIDTH="291" HEIGHT="23" AUNIT="WON" AUNITVALUE="1">(단위 : 원)</TU>\n</TR>\n</TBODY>\n</TABLE>\n')
w('<TABLE ACLASS="FINANCE" AFIXTABLE="N" WIDTH="725" BORDER="1" FRAME="BORDER" RULES="ALL">\n<COLGROUP WIDTH="725">\n<COL WIDTH="205"></COL>\n' + '<COL WIDTH="130"></COL>\n' * 4 + '</COLGROUP>\n<THEAD>\n<TR ACOPY="N" ADELETE="N" HEIGHT="30">\n')
for h in ["과 목", "자본금", "자본잉여금", "이익잉여금", "총 계"]:
    w(f'<TH VALIGN="MIDDLE" WIDTH="121" HEIGHT="23">{h}</TH>\n')
w("</TR>\n</THEAD>\n<TBODY>\n")
for name, vals in [
    ("2024.01.01(전기초)", [1000, 500, 3244, 4744]),
    ("연차배당", [None, None, -80, -80]),
    ("당기순이익", [None, None, 640, 640]),
    ("2024.12.31(전기말)", [1000, 500, 3804, 5304]),
    ("2025.01.01(당기초)", [1000, 500, 3804, 5304]),
    ("연차배당", [None, None, -100, -100]),
    ("당기순이익", [None, None, 880, 880]),
    ("2025.12.31(당기말)", [1000, 500, 4584, 6084]),
]:
    w('<TR ACOPY="Y" ADELETE="Y" HEIGHT="30">\n')
    w(f'<TE ADELIM="0" ALEVEL="0" WIDTH="196" HEIGHT="23">{name}</TE>\n')
    for i, v in enumerate(vals, 1):
        w(f'<TE ALIGN="RIGHT" ADELIM="{i}" WIDTH="121" HEIGHT="23">{won(m(v))}</TE>\n')
    w("</TR>\n")
w("</TBODY>\n</TABLE>\n</TABLE-GROUP>\n")

# 현금흐름표 (간접법)
fin_head("PCF-I:A", 7, "현 금 흐 름 표", "Statement of Cash Flows", PER_PL)
fin_table(HDR, [
    ("Ⅰ.영업활동으로인한현금흐름", 0, None, 870, None, 620),
    ("1.당기순이익", 1, 880, None, 640, None),
    ("2.현금의유출이없는비용등의가산", 1, 276, None, 252, None),
    ("감가상각비", 2, 188, None, 170, None),
    ("무형자산상각비", 2, 15, None, 15, None),
    ("퇴직급여", 2, 70, None, 65, None),
    ("대손상각비", 2, 3, None, 2, None),
    ("3.영업활동으로인한자산부채의변동", 1, -286, None, -272, None),
    ("매출채권의 증가", 2, -300, None, -250, None),
    ("미수금의 증가", 2, -12, None, -8, None),
    ("선급비용의 감소(증가)", 2, -3, None, 2, None),
    ("재고자산의 증가", 2, -20, None, -60, None),
    ("매입채무의 증가", 2, 70, None, 90, None),
    ("미지급금의 증가(감소)", 2, 20, None, -10, None),
    ("미지급법인세의 증가", 2, 25, None, 15, None),
    ("예수금의 증가", 2, 4, None, 4, None),
    ("퇴직금의 지급", 2, -20, None, -15, None),
    ("퇴직연금운용자산의 증가", 2, -50, None, -40, None),
    ("Ⅱ.투자활동으로인한현금흐름", 0, None, -300, None, -250),
    ("1.투자활동으로인한현금유출액", 1, -300, None, -250, None),
    ("단기금융상품의 증가", 2, -100, None, -100, None),
    ("기계장치의 취득", 2, -200, None, -150, None),
    ("Ⅲ.재무활동으로인한현금흐름", 0, None, -300, None, -90),
    ("1.재무활동으로인한현금유입액", 1, 0, None, 200, None),
    ("단기차입금의 차입", 2, 0, None, 200, None),
    ("2.재무활동으로인한현금유출액", 1, -300, None, -290, None),
    ("단기차입금의 상환", 2, -200, None, 0, None),
    ("장기차입금의 상환", 2, 0, None, -210, None),
    ("배당금의 지급", 2, -100, None, -80, None),
    ("Ⅳ.현금의증가(Ⅰ+Ⅱ+Ⅲ)", 0, None, 270, None, 280),
    ("Ⅴ.기초의현금", 0, None, 980, None, 700),
    ("Ⅵ.기말의현금", 0, None, 1250, None, 980),
])
w("</SECTION-1>\n")

# ---------- 주석 ----------
w('<SECTION-1 ACLASS="MANDATORY" APARTSOURCE="SOURCE">\n')
w('<TITLE ATOC="Y" AASSOCNOTE="D-0-1-0-0" ATOCID="8" ENG="Notes to the Financial Statements">주석</TITLE>\n')
simple_table(["제 10(당)기 2025년 12월 31일 현재", "제 9(전)기 2024년 12월 31일 현재", "샘플전자 주식회사"])

def note_table(head, rows, widths=None):
    n = len(head)
    widths = widths or [240] + [int(360 / (n - 1))] * (n - 1)
    tw = sum(widths)
    w(f'<TABLE ACLASS="NORMAL" AFIXTABLE="N" BORDER="1" WIDTH="{tw}">\n<COLGROUP WIDTH="{tw}">\n')
    for cw in widths:
        w(f'<COL WIDTH="{cw}"></COL>\n')
    w("</COLGROUP>\n<TBODY>\n<TR ACOPY=\"Y\" ADELETE=\"Y\" HEIGHT=\"30\">\n")
    for h, cw in zip(head, widths):
        w(f'<TD WIDTH="{cw - 9}" VALIGN="MIDDLE" ALIGN="CENTER" HEIGHT="23" USERMARK="BC0XF2F2F2">{esc(h)}</TD>\n')
    w("</TR>\n")
    for r in rows:
        w('<TR ACOPY="Y" ADELETE="Y" HEIGHT="30">\n')
        for j, (c, cw) in enumerate(zip(r, widths)):
            if j == 0 or isinstance(c, str):
                w(f'<TD VALIGN="MIDDLE" ALIGN="{"CENTER" if j == 0 else "RIGHT"}" WIDTH="{cw - 9}" HEIGHT="23">{esc(c)}</TD>\n')
            else:
                w(f'<TD VALIGN="MIDDLE" ALIGN="RIGHT" WIDTH="{cw - 9}" HEIGHT="23">{won(m(c))}</TD>\n')
        w("</TR>\n")
    w("</TBODY>\n</TABLE>\n")
    P(CR)

def unit_line(t="(단위 : 원)"):
    simple_table([t], align="RIGHT")

P("1. 회사의 개요", "B")
P("샘플전자 주식회사(이하 '당사')는 2016년에 설립되어 전자부품을 제조·판매하고 있습니다. 당사의 본점은 서울특별시에 있으며, 당기말 현재 자본금은 1,000,000,000원입니다." + CR + "당기말 현재 주주 현황은 다음과 같습니다.")
note_table(["주주명", "소유주식수(주)", "지분율(%)"], [
    ["김샘플", "120,000", "60.00"], ["이예시", "50,000", "25.00"], ["기타", "30,000", "15.00"], ["합계", "200,000", "100.00"]])

P("2. 중요한 회계정책", "B")
P("2.1 재무제표 작성기준")
P("당사는 주식회사 등의 외부감사에 관한 법률에 따라 일반기업회계기준을 적용하여 재무제표를 작성하였습니다.")
P("2.2 유형자산")
P("유형자산은 취득원가에서 감가상각누계액을 차감한 금액으로 표시하며, 아래의 내용연수에 따라 정액법으로 상각합니다.")
note_table(["구분", "내용연수"], [["건물", "40년"], ["기계장치", "10년"], ["차량운반구", "5년"]])

P("3. 현금및현금성자산", "B")
P("보고기간말 현재 현금및현금성자산의 내역은 다음과 같습니다.")
unit_line()
note_table(["구분", "당기말", "전기말"], [["현금", 1.0, 1.5], ["보통예금", 1249.0, 978.5], ["합계", 1250, 980]])

P("4. 매출채권", "B")
P("(1) 보고기간말 현재 매출채권의 내역은 다음과 같습니다.")
unit_line()
note_table(["구분", "당기말", "전기말"], [["매출채권", 2400, 2100], ["대손충당금", -24, -21], ["장부금액", 2376, 2079]])
P("(2) 당기와 전기 중 대손충당금의 변동내역은 다음과 같습니다.")
unit_line()
note_table(["구분", "당기", "전기"], [["기초", 21, 19], ["대손상각비", 3, 2], ["기말", 24, 21]])

P("5. 재고자산", "B")
P("보고기간말 현재 재고자산의 내역은 다음과 같습니다.")
unit_line()
note_table(["구분", "당기말", "전기말"], [["제품", 640, 590], ["원재료", 380, 410], ["합계", 1020, 1000]])

P("6. 유형자산", "B")
P("당기 중 유형자산 장부금액의 변동내역은 다음과 같습니다.")
unit_line()
dep_total = -178 if WITH_ERROR else -188   # ERR: 일부러 틀린 합계(맞는 값은 188,000,000)
note_table(["구분", "기초", "취득", "처분", "감가상각비", "기말"], [
    ["토지", 1500, None, 0, 0, 1500],
    ["건물", 1600, None, 0, -50, 1550],
    ["기계장치", 600, 200, 0, -120, 680],
    ["차량운반구", 54, None, 0, -18, 36],
    ["합계", 3754, 200, 0, dep_total, 3766],
], widths=[150, 90, 90, 90, 90, 90])
P("당기 감가상각비 188,000,000원 중 103,000,000원은 제조원가에, 85,000,000원은 판매비와관리비에 포함되어 있습니다.")

P("7. 차입금", "B")
P("(1) 보고기간말 현재 단기차입금의 내역은 다음과 같습니다.")
unit_line()
note_table(["차입처", "연이자율(%)", "당기말", "전기말"], [["가나은행(운전자금)", "4.50", 800, 1000], ["합계", "", 800, 1000]], widths=[180, 120, 150, 150])
P("(2) 보고기간말 현재 장기차입금의 내역은 다음과 같습니다.")
unit_line()
note_table(["차입처", "연이자율(%)", "당기말", "전기말"], [["다라은행(시설자금)", "4.20", 1000, 1000], ["합계", "", 1000, 1000]], widths=[180, 120, 150, 150])

P("8. 자본금", "B")
P("당사가 발행할 주식의 총수는 1,000,000주이고, 보고기간말 현재 발행한 주식수는 보통주 200,000주(1주당 금액 5,000원)이며, 자본금은 1,000,000,000원입니다.")

P("9. 법인세비용", "B")
P("당기 법인세비용은 190,000,000원(전기 134,000,000원)이며, 전액 당기 법인세부담액입니다.")
w("</SECTION-1>\n")
w("</BODY>\n</DOCUMENT>\n")

contents = "".join(out)
meta = ('<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\n<METAINFO>\n\t<GENERATOR schema="dart4.xsd" editver="1.0"/>\n'
        '\t<DOCUMENT-HEADER regcik="00000000" regname="샘플전자(주)"/>\n\t<DOCUMENT-INFO docver="6.0" bsn-id="00760" rpt-id="00760" doc-id="00760" iscorrection="N"/>\n'
        '\t<EXTRACT type="header">\n\t\t<ITEM name="TOT_ASSETS" value="9481"/>\n\t\t<ITEM name="TOT_DEBTS" value="3397"/>\n\t\t<ITEM name="TOT_SALES" value="18500"/>\n\t</EXTRACT>\n</METAINFO>\n')
with zipfile.ZipFile(OUT, "w", zipfile.ZIP_DEFLATED) as z:
    z.writestr("contents.xml", contents)
    z.writestr("meta.xml", meta)
print("만듦:", OUT, len(contents), "자", "| 일부러 틀린 합계:", WITH_ERROR)
