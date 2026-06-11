#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""올더베러 제안 PPT 빌더 — PPT v1.2 인덱스 순서 기반 + 내용 docx + VOC 인사이트 통합"""
from pptx import Presentation
from pptx.util import Inches, Pt, Emu
from pptx.dml.color import RGBColor
from pptx.enum.text import PP_ALIGN, MSO_ANCHOR

BROWN   = RGBColor(0x4A, 0x35, 0x26)
BROWN2  = RGBColor(0x6B, 0x4F, 0x3A)
ACCENT  = RGBColor(0xC8, 0x7A, 0x3C)   # 강조 오렌지브라운
LIGHT   = RGBColor(0xF2, 0xEC, 0xE3)
GRAY    = RGBColor(0x55, 0x55, 0x55)
WHITE   = RGBColor(0xFF, 0xFF, 0xFF)
FONT    = "맑은 고딕"

prs = Presentation()
prs.slide_width  = Inches(13.333)
prs.slide_height = Inches(7.5)
BLANK = prs.slide_layouts[6]
SW, SH = prs.slide_width, prs.slide_height

def _set_font(run, size, bold=False, color=BROWN):
    run.font.name = FONT
    run.font.size = Pt(size)
    run.font.bold = bold
    run.font.color.rgb = color

def add_rect(slide, x, y, w, h, color):
    sp = slide.shapes.add_shape(1, x, y, w, h)  # rectangle
    sp.fill.solid(); sp.fill.fore_color.rgb = color
    sp.line.fill.background()
    sp.shadow.inherit = False
    return sp

def cover():
    s = prs.slides.add_slide(BLANK)
    add_rect(s, 0, 0, SW, SH, BROWN)
    add_rect(s, Inches(0.9), Inches(3.05), Inches(2.2), Inches(0.09), ACCENT)
    tb = s.shapes.add_textbox(Inches(0.9), Inches(2.0), Inches(11.5), Inches(1.0)).text_frame
    tb.word_wrap = True
    r = tb.paragraphs[0].add_run(); r.text = "올리브영 플랫폼 PB 마케팅 제안"
    _set_font(r, 40, True, WHITE)
    tb2 = s.shapes.add_textbox(Inches(0.9), Inches(3.25), Inches(11.5), Inches(1.6)).text_frame
    tb2.word_wrap = True
    for i, line in enumerate([
        "올더베러(All the Better) 브랜드 빌딩 캠페인",
        "“팔리는 브랜드를, 빅 프로모션의 주인공으로”",
        "캠페인 예산 2.5억 원 · 기간 2026.07.01 ~ 12.31 (6개월)",
    ]):
        p = tb2.paragraphs[0] if i == 0 else tb2.add_paragraph()
        r = p.add_run(); r.text = line
        _set_font(r, 17 if i < 2 else 13, i == 1, LIGHT)
    tb3 = s.shapes.add_textbox(Inches(0.9), Inches(6.6), Inches(5), Inches(0.5)).text_frame
    r = tb3.paragraphs[0].add_run(); r.text = "BAT"; _set_font(r, 16, True, ACCENT)

def divider(num, en, ko=""):
    s = prs.slides.add_slide(BLANK)
    add_rect(s, 0, 0, SW, SH, BROWN)
    add_rect(s, Inches(0.9), Inches(3.5), Inches(1.6), Inches(0.08), ACCENT)
    tb = s.shapes.add_textbox(Inches(0.9), Inches(2.7), Inches(11), Inches(0.9)).text_frame
    r = tb.paragraphs[0].add_run(); r.text = num; _set_font(r, 22, True, ACCENT)
    tb2 = s.shapes.add_textbox(Inches(0.9), Inches(3.7), Inches(11), Inches(1.2)).text_frame
    tb2.word_wrap = True
    r = tb2.paragraphs[0].add_run(); r.text = en; _set_font(r, 40, True, WHITE)
    if ko:
        r2 = tb2.add_paragraph().add_run(); r2.text = ko; _set_font(r2, 18, False, LIGHT)

def header(s, title, sub=None):
    add_rect(s, 0, 0, SW, Inches(0.16), BROWN)
    tb = s.shapes.add_textbox(Inches(0.6), Inches(0.4), Inches(12.1), Inches(0.95)).text_frame
    tb.word_wrap = True
    r = tb.paragraphs[0].add_run(); r.text = title; _set_font(r, 25, True, BROWN)
    if sub:
        r2 = tb.add_paragraph().add_run(); r2.text = sub; _set_font(r2, 13, False, ACCENT)
    add_rect(s, Inches(0.62), Inches(1.5), Inches(1.4), Inches(0.06), ACCENT)

def content(title, bullets, sub=None, table=None, note=None, body_top=1.8):
    s = prs.slides.add_slide(BLANK)
    header(s, title, sub)
    top = Inches(body_top)
    if bullets:
        bh = Inches(4.9) if not table else Inches(2.2)
        tf = s.shapes.add_textbox(Inches(0.65), top, Inches(12.0), bh).text_frame
        tf.word_wrap = True
        for i, b in enumerate(bullets):
            lvl = 0
            txt = b
            if isinstance(b, tuple):
                lvl, txt = b
            p = tf.paragraphs[0] if i == 0 else tf.add_paragraph()
            p.space_after = Pt(6)
            mark = "■ " if lvl == 0 else "– "
            r = p.add_run(); r.text = mark + txt
            _set_font(r, 14 if lvl == 0 else 12.5, lvl == 0, BROWN if lvl == 0 else GRAY)
            p.level = 0
        top = Emu(int(top) + int(bh) + Inches(0.1))
    if table:
        headers, rows = table
        ncol = len(headers); nrow = len(rows) + 1
        gtop = top if bullets else Inches(body_top)
        gh = Inches(min(4.6, 0.42 * nrow + 0.2))
        gx = Inches(0.65); gw = Inches(12.0)
        g = s.shapes.add_table(nrow, ncol, gx, gtop, gw, gh).table
        for c, htxt in enumerate(headers):
            cell = g.cell(0, c)
            cell.fill.solid(); cell.fill.fore_color.rgb = BROWN
            tf = cell.text_frame; tf.word_wrap = True
            r = tf.paragraphs[0].add_run(); r.text = htxt
            _set_font(r, 11, True, WHITE)
        for ri, row in enumerate(rows, start=1):
            for c in range(ncol):
                cell = g.cell(ri, c)
                cell.fill.solid(); cell.fill.fore_color.rgb = LIGHT if ri % 2 else WHITE
                tf = cell.text_frame; tf.word_wrap = True
                r = tf.paragraphs[0].add_run(); r.text = row[c] if c < len(row) else ""
                _set_font(r, 10.5, c == 0, BROWN if c == 0 else GRAY)
    if note:
        nb = s.shapes.add_textbox(Inches(0.65), Inches(6.95), Inches(12), Inches(0.4)).text_frame
        r = nb.paragraphs[0].add_run(); r.text = note; _set_font(r, 10, False, ACCENT)

# ============ 00. Index / 재해석 ============
cover()

content("목차 (INDEX) — RFP / Campaign Brief 재해석", [
    "PART 00. 우리가 이해한 과제 (재해석)",
    "PART 01. Strategy — 올영세일 1등 브랜드로 만드는 전략 구조",
    "PART 02. Product & Message — 제품 우선순위 · 유형별 메시지 · 타겟",
    "PART 03. Seeding Content — 시딩 1순위 · 크리에이터 · 콘텐츠 앵글",
    "PART 04. Performance Boost — 콘텐츠+퍼포 결합 · LMF · 올영 인앱 · CRM",
    "PART 05. Operation — 로드맵 · KPI&예산 · 심의 · 리포팅 · TF",
    "PART 06. Why Us",
], sub="00 Index")

content("올더베러는 안 팔리는 브랜드가 아닙니다", [
    "이미 올리브영이라는 강력한 홈그라운드 안에서 판매 가능성을 증명 (웰니스 매출 상위 · 구매 후 평균 3~4개 추가 구매)",
    "문제는 소비자가 올더베러를 ‘이름으로’ 선택하지 않는다는 점",
    "매대에서 제품을 집어들 수는 있지만, ‘올더베러를 선택했다’는 기억으로 저장되지 않음",
    "아직 독립 브랜드로 각인되어 있다고 보기 어렵다",
], sub="00 재해석 ①")

content("2.5억으로 가장 빠르게 브랜드 자산을 만드는 법", [
    "이번 캠페인의 목표는 단순히 올더베러를 ‘많이 노출’하는 것이 아니다",
    "제한된 2.5억 안에서 가장 빠른 길 = 올영세일에서 상위 랭킹/1위라는 ‘선택 근거’를 먼저 만드는 것",
    "그 성과를 콘텐츠 · 후기 · 검색 · CRM 자산으로 반복 전환",
    "브랜드 자산이 약한 올더베러에게 ‘랭킹’은 가장 빠른 인지 자산",
], sub="00 재해석 ②")

content("시딩-콘텐츠-퍼포먼스 연결로 비즈니스 임팩트를 만든다", [
    "단순 시딩 콘텐츠 ‘생산’이 아니라, 시딩 → 콘텐츠 → 퍼포먼스 연결 구조로 설계",
    "UGC + 퍼포 결합 레퍼런스 제시 (→ PART 04 / 대표 레퍼런스)",
    "‘브랜드를 알리는 캠페인’이 아니라 ‘올영세일에서 선택받는 구조를 만드는 캠페인’",
], sub="00 재해석 ③")

# ============ 01. Strategy ============
divider("01", "Strategy", "올영세일 1등 브랜드로 만드는 전략 구조")

content("웰니스 시장: 폭발 직전, 그러나 갇혀 있다", [
    "건기식 시장 ’23 6.1조 → ’24 5.95조 → ’25 6.4조(+5.2%) → ’28 7.5조 전망 (재성장 국면)",
    "온라인 구매 비중 54%→59% — 디지털 콘텐츠가 구매를 만든다",
    "진짜 경쟁자는 전통 건기식(종근당·고려은단·KGC)이 아니라 낫띵베럴 같은 라이프스타일 웰니스 브랜드",
    "소비자 머릿속에 각인된 ‘일상형 웰니스 브랜드’는 아직 부재 → 선점 기회",
], sub="01-1 시장 진단")

content("올리브영 / 올리브베러는 올더베러의 홈그라운드", [
    "주 무대 = 올리브영·올리브베러 생태계. 소비자는 이미 매대·앱에서 탐색하고 구매를 결정",
    "올리브영 앱 · 매장 · 기획전 · 리뷰 · 랭킹을 하나의 ‘미디어’로 본다",
    "외부에서 인지를 처음부터 만들기보다, 올영세일 시점에 소비자의 선택을 집중시키는 방식이 효율적",
], sub="01-2 홈그라운드 어드밴티지")

content("2.5억 · 6개월, 어디에 쓸 것인가", [
    "대형 모델 플레이 · 대규모 OOH로 넓고 얕은 노출을 만들지 않는다",
    "구매 판단에 영향을 주는 시딩 콘텐츠 자산 + 전환형 소재 + 퍼포먼스 부스팅에 집중",
    "‘알리는 캠페인’이 아니라 ‘올영/올영세일에서 선택받는 구조를 만드는 캠페인’",
], sub="01-3 예산 운용 원칙")

content("[과제 재정의] 올더베러를 올리브영/올리브베러 1등 브랜드로", [
    "지금 확보할 자산 = 추상적 인지·호감 X → 제품 각인 + 명확한 선택 근거",
    "콘텐츠 인지 + 사용 후기 + 올리브영 내 랭킹 + 반복 노출 = “많이 사는 / 선택해도 되는 제품” 신호",
    "올더베러는 인지가 쌓여 1등이 되는 게 아니라, ‘1등이라는 선택 신호’를 먼저 만들어 인지를 쌓는다",
], sub="01-4 캠페인 과제 재정의")

content("캠페인 구조: SEED → PROVE → BOOST → RETAIN", None,
    sub="01-5 4-STEP 프레임",
    table=(["STEP", "내용"], [
        ["STEP 1 · SEED", "제품 경험량·후기량을 만든다 — 챌린저스, 마이크로/나노 시딩, KOL 시딩으로 구매 판단 근거 확보"],
        ["STEP 2 · PROVE", "구매 이유를 증명한다 — 리뷰·해시태그·UGC를 축적해 선택할 이유를 만든다"],
        ["STEP 3 · BOOST", "시딩 성과를 퍼포먼스로 증폭 — PA·올영 인앱광고·LMF 테스트로 세일 유입·전환 집중 견인"],
        ["STEP 4 · RETAIN", "9월 성과를 12월까지 — 위닝 소재·후기·CRM(플친)으로 재구매·쟁임 구매 유도"],
    ]),
    note="* 9월·12월 올영세일 & 11월 블프 이슈 고려")

content("올리브영 1등 달성 케이스  [대표 레퍼런스 추가 예정]", None,
    sub="01-6 우리는 이미 1위를 만들어 봤다",
    table=(["항목", "내용"], [
        ["브랜드 / 카테고리", "[추가 예정]"],
        ["당시 문제·니즈", "[추가 예정]"],
        ["적용 방법론", "시딩 + 퍼포먼스(PA) + 챌린저스 랭킹 견인 + 콘텐츠 절대량"],
        ["랭킹 변화", "○위 → 1위 [추가 예정]"],
        ["매출·성과 / 인사이트", "[추가 예정] — 올영세일/랭킹 캡처·그래프 첨부"],
    ]))

# ============ 02. Product & Message ============
divider("02", "Product & Message", "제품 우선순위 · 유형별 메시지 · 타겟")

content("올더베러는 웰니스 브랜드이지만, 제품마다 다른 언어가 필요", [
    "제품에 따라 적용되는 규제·심의 정도에 차이가 존재",
    "효능을 말할 수 있는 제품(건기식)과, 상황·루틴으로 풀어야 하는 제품(일반식품)이 다르다",
    "심의 난이도/기간 ↔ 크리에이터 풀:  뷰티·패션·식품  <  건기식  <  …  <  금융",
], sub="02-1 제품 유형별 언어 분리")

content("제품 포트폴리오 우선순위 — 선발대 & 본대", None,
    sub="02-2 7월 선발대 → 심의 후 본대 합류",
    table=(["구분", "제품", "역할"], [
        ["선발대 (7월~)", "엑스트라버진 올리브오일 · 멜라나잇 구미 · 콜라겐 구미 (일반/가공식품)", "심의 부담↓ · TPO 콘텐츠화 쉬움 → 시장 언어 테스트·시딩 자산 선확보"],
        ["본대 (심의 후)", "철분 · 올인원 · 루테인 · 비타민C · 바나바잎 · 비오틴 (건기식 구미)", "기능성 메시지로 구매 이유 강화 · 12월 세일 라인업 규모감·쟁임 구매"],
        ["확장 효과", "“하나 사보고 여러 개 담는” PB 라인 확장", "히어로 1종 → 라인 전체 매출 견인(낙수)"],
    ]))

content("제품 유형별 메시지 설계 — RTB를 소비자 언어로", None,
    sub="02-3 공급자 언어 → 소비자 언어 (심의 분리)",
    table=(["제품(유형)", "공급자 언어(버림)", "소비자 언어(가져감)"], [
        ["올리브오일 (일반식품)", "폴리페놀 542mg/kg 고함량", "공복에 한 캡슐, 비린내 없이 끝까지 먹는"],
        ["멜라나잇 (일반식품)", "식물성 멜라토닌 1mg", "잠 못 드는 밤이 아니라 ‘자기 전 리추얼’"],
        ["콜라겐 (일반식품)", "피부 탄력 콜라겐", "디저트 대신 챙기는 한 알"],
        ["루테인 (건기식)", "눈이 좋아지는", "루테인, 눈 건강에 도움을 줄 수 있음(인정문구)"],
        ["바나바잎 (건기식)", "혈당 낮춰주는", "식후 혈당상승 억제에 도움을 줄 수 있음"],
    ]))

content("타겟 전략 — 30 · 20 · 40 + 방한 외국인 (+부모)", None,
    sub="02-4 키워드는 공통(‘채움’), 훅·채널·크리에이터는 분리",
    table=(["트랙", "타겟", "역할 / TPO"], [
        ["각인 (1순위)", "30대 직장인", "웰니스 루틴 형성 — 브랜드명+키워드 회상"],
        ["입문·재미", "20대", "웰니스 입문템 — 밈·확산·해시태그 점유"],
        ["전환", "40대 여성", "실질 구매 전환 — 올리브오일 검색 주도"],
        ["K-웰니스", "방한 외국인", "중장기 브랜드 자산 — 실제 웰니스 씬 노출"],
        ["가족 확대", "자녀를 둔 부모(3040)", "자녀 챙김 → 본인·가족 루틴, 객단가↑"],
    ]))

# ============ 03. Seeding Content ============
divider("03", "Seeding Content", "시딩 1순위 · 크리에이터 · 콘텐츠 앵글")

content("왜 시딩이 1순위인가", [
    "올더베러는 아직 브랜드 이름 자체로 소비자를 끌어당기는 힘이 약하다",
    "선택을 망설이지 않도록 후기·리뷰·해시태그·콘텐츠 볼륨을 우선 확보",
    "시딩 = 노출 수량이 아니라 ‘구매 판단 근거’를 먼저 만드는 일",
    "VOC 근거: 전체 재구매율 15.5% · 구매이유 1위가 ‘가격·할인’(160건) → 가격 외 명확한 선택 근거 자산 필요",
], sub="03-1 Seeding First")

content("크리에이터 전략", None,
    sub="03-2 역할 구분 · 리스트업 · 관계 빌드업(KOL 무가시딩)",
    table=(["등급", "관계 빌드업 방식"], [
        ["슈퍼스타급", "연 4회+ 정기 접촉 · 커스텀 키트(개인 이벤트 연계) · 행보 모니터링 → 자연스러운 연결"],
        ["메가/매크로", "무가 시딩 + 커스텀 레터/키트로 차별화"],
        ["마이크로/나노", "제품 제공 · 동일 메시지 동시다발 게시(반복 노출)"],
        ["전문가", "약사·영양사 우선 편성 — 신뢰 + 심의 통과율 확보"],
    ]))

content("크리에이터 콘텐츠 기획 — 제품별 콘텐츠 앵글 (VOC 기반)", None,
    sub="성공 훅 = 명확한 타겟 × 구체적 상황/에피소드 × 반전 또는 팩트",
    table=(["제품", "컨셉 / 훅", "길이", "VOC 근거"], [
        ["캡슐 A", "생올리브오일 챌린지 → 캡슐 반전 (Before/After)", "15s", "섭취 편의 55.1%"],
        ["캡슐 B", "‘알약 잘 못 먹는 사람용’ 목넘김 실연", "12s", "목넘김 단점 32.6% 정면 방어"],
        ["구미·루테인", "모니터 앞 직장인의 눈 챙기기 (인정문구)", "15s", "건기식 인정 기능성"],
        ["구미·바나바잎", "식후 디저트처럼 (식후 혈당상승 억제 도움)", "15s", "식후 상황 연출"],
        ["구미·멜라나잇", "자기 전 무드 리추얼 (수면효능 배제·무드만)", "12s", "일반식품 — 효능 금지"],
        ["구미·콜라겐", "디저트 대신 한 알 (피부효능 배제)", "12s", "일반식품 — 효능 금지"],
        ["스틱 A", "가방에 쏙, 출근길 한 포", "12s", "섭취 편의 48.3%"],
        ["스틱 B", "오일 맛 솔직 리뷰 (선제 인정→섭취팁)", "15s", "맛 불호 69.8% 방어"],
    ]),
    note="※ 각 컨셉은 VOC 키워드를 후킹으로 사용 · 식품 분류별 심의 가드 + 풀 대본/자막 타임코드 보유")

content("오프라인 인증형 시딩 콘텐츠", [
    "광화문 등 올리브영/올리브베러 오프라인 매장을 시딩 콘텐츠의 ‘현장 인증 접점’으로 활용",
    "“올영에서 실제로 파는 제품”이라는 현장감 → 자연 발견형 바이럴",
    "방문 동선 콘텐츠 + 방문 크리에이터 전용 콘텐츠로 온·오프 연결",
], sub="03-3 매장 = 현장 인증")

content("시딩 프로그램 구조 — 점점 커지는 물량", [
    "① 마이크로/나노 시딩 (EGC·전문가)",
    (1, "② + 챌린저스 / 체험단 / 어필리에이트"),
    (1, "③ + KOL 무가시딩 / 파워페이지 / 바이럴 / GEO(AEO)"),
    "7월부터 시장 언어·콘텐츠 앵글을 테스트하고, 8월 말~9월 초 올영세일 직전에 콘텐츠가 충분히 적재되도록 물량 집중",
], sub="03-4 단계별 물량 확대 (바이럴 파트 별도 분리 검토)")

content("챌린저스 / 체험단 / 어필리에이트", [
    "챌린저스 다수 동시 게시(실시간 랭킹 견인) 전략으로 올영세일 시즌 카테고리 랭킹 상위/1위 확보",
    "올리브영 쇼핑 큐레이터 / 쿠팡 파트너스 연계로 전환 보강",
], sub="03-5 랭킹 견인 엔진")

content("KOL 무가시딩 / 파워페이지 / 바이럴 / GEO(AEO)", [
    "인스타그램·블로그·X 마이크로 바이럴 + 뷰티 정보 플랫폼·대형 커뮤니티 체험단/파워페이지",
    "PR = 검색 답변 시대의 전략 — 단순 푸시가 아니라 검색 풀(Pull) 자산",
    "키워드 설계 → SEO / AEO / GEO 영역 구축 → 자연 유입(검색 풀) 증가",
], sub="03-6 검색 풀(Pull) 자산화")

# ============ 04. Performance Boost ============
divider("04", "Performance Boost", "콘텐츠+퍼포 결합 · LMF · 올영 인앱 · CRM")

content("시딩만으로는 1등 브랜드를 만들기 어렵다", [
    "콘텐츠와 후기를 확보하더라도, 세일 기간 구매 전환으로 연결하지 않으면 랭킹 견인력은 제한적",
    "따라서 시딩 결과물은 반드시 퍼포먼스 소재 + 올영 인앱 유입·전환으로 활용되어야 한다",
], sub="04-1 시딩 → 전환 연결")

content("콘텐츠 + 퍼포 결합 구조", [
    "시딩 콘텐츠 수집  →  위닝 콘텐츠 선별  →  LMF 테스트  →  콘텐츠+배너 광고 집행  →  올영세일 집중 운영",
    "뷰티 & 타 카테고리에서 올영세일 성과를 기록한 프로세스를 그대로 적용",
    "UGC + 퍼포 결합 레퍼런스 제시 (→ 대표 레퍼런스)",
], sub="04-2 검증된 전환 프로세스")

content("비주얼 가이드 · KV · 퍼포 소재 레이아웃", [
    "올더베러 톤앤무드를 적용한 올리브오일/웰니스 구미 제품별 KV",
    "매장 디스플레이·POP 연계 + 퍼포 소재 레이아웃",
    "히트 패턴 앵커링 — 검증된 히트 팩트를 전 채널에 반복 적용",
], sub="04-3 KV & 소재 시스템")

content("LMF 테스트 — 위닝 메시지 발굴", [
    "어떤 문구·썸네일·후킹 구조가 실제 반응을 만드는지 먼저 확인",
    "검증된 위닝 소재를 9월 올영세일과 12월 재부스팅에 활용",
    "초개인화 메시지 / VOC 인사이트 기반 검증-활용-검증-활용 루프",
], sub="04-4 Language-Market-Fit")

content("퍼포 매체 활용 방안", [
    "올영 인앱광고를 코어로, 메타·유튜브 등 전환 매체 믹스",
    "위닝 소재 중심으로 세일 기간 유입·전환 집중",
], sub="04-5 매체 믹스")

content("올영 인앱광고", [
    "올리브영 플랫폼 내 올더베러 제품 반복 노출",
    "후기와 위닝 소재를 올영 인앱광고로 연결 → 세일 기간 유입과 랭킹 상승을 집중 견인",
], sub="04-6 In-App")

content("CRM — 카카오 플친", [
    "올영세일 기간 재구매 연결 CRM 소재 기획",
    "오프라인 샘플 증정 · 제품 조회자 대상 리타겟 · 재구매 · 라인 확장 구매",
    "‘채움’ 루틴 리마인드로 재구매 고착",
], sub="04-7 재구매·확장")

# ============ 05. Operation ============
divider("05", "Operation", "로드맵 · KPI&예산 · 심의 · 리포팅 · TF")

content("6개월 운영 로드맵 — 9월에 만들고, 11~12월에 다시 키운다", None,
    sub="05-1 1·2차 웨이브",
    table=(["월", "핵심 활동"], [
        ["7월", "콘텐츠 설계/테스트 — 제품별 위닝 앵글·메시지 확보"],
        ["8월", "시딩 본격화 — 올영세일 전 콘텐츠·후기 축적"],
        ["9월", "1차 부스팅 — 올영세일 집중, 퍼포 연계 유입 증대 및 1위 확보"],
        ["10월", "낙수효과 유지 — 후기 확산, 콘텐츠 보강, 위닝 소재 지속"],
        ["11월", "세일 준비 — 위닝 소재 테스트, 콘텐츠 확보 (블프)"],
        ["12월", "2차 부스팅 — 올영세일 집중 집행, 확장·재구매 활성화"],
    ]))

content("KPI & Budget", None,
    sub="05-2 성과 지표 & 예산 배분",
    table=(["구분", "내용"], [
        ["KPI · 검색량", "브랜드 월 검색량 3,000 → 10,000 (연말) + 익년 1Q 추가 상승"],
        ["KPI · 랭킹", "올영세일 카테고리 랭킹 상위 진입/1위 · 시그니처 키워드 노출률 90%+"],
        ["예산 원칙", "2.5억 — 퍼포 광고 운영 비용 확보 필수 (콘텐츠·시딩·증폭 집중)"],
        ["월별 배분", "7–8월 25% / 9–10월 40% / 11–12월 35% — 세일 2개월에 약 53% 집중"],
    ]))

content("심의 대응 체계 + AI 심의 가이드 검증", [
    "올더베러는 일반식품·일반/가공식품·건강기능식품이 함께 존재 → ‘선제작 후검수’는 리스크가 크다",
    "제품 유형별 표현 가능 범위·금지 표현·크리에이터 발화 가이드를 사전 정리",
    "AI 기반 심의 가이드라인 검증 봇 → 콘텐츠 사전 필터링으로 반려↓·심의 기간 단축 (삼성증권 사례)",
], sub="05-3 차별화 포인트",
    table=(["제품", "❌ 위험 표현", "✅ 안전 표현"], [
        ["올리브오일(일반)", "혈관에 좋은", "엑스트라 버진 올리브오일을 한 포로"],
        ["멜라나잇(일반)", "숙면 / 잠 잘 오는", "자기 전 리추얼 / 하루 마무리"],
        ["콜라겐(일반)", "피부·주름", "디저트 대신 챙기는 젤리"],
        ["루테인(건기식)", "눈이 좋아지는", "눈 건강에 도움을 줄 수 있음"],
        ["바나바잎(건기식)", "혈당 낮춰주는 / 당뇨", "식후 혈당상승 억제에 도움"],
    ]),
    body_top=1.7, note="※ 일반식품은 건강·기능 효능 일체 금지 · 건기식은 인정 기능성 문구 그대로 · 표시광고 표기 필수")

content("성과 리포팅 솔루션", [
    "정성·정량 데이터 종합 분석 + 경쟁사 동향 → 후속 캠페인 방향 제안",
    "실시간 대시보드 기반 리포팅 (시딩·랭킹·전환·검색량 추이)",
    "VOC 딥다이브 대시보드를 캠페인 중에도 갱신해 메시지·앵글 최적화",
], sub="05-4 데이터 리포팅")

content("올영 PB TF 인력 구성", None,
    sub="05-5 총괄 PM급 시니어 1인 외 실무 리소스 필수 편성",
    table=(["구분", "R&R"], [
        ["디렉터", "제안 전략·크리에이티브 방향성 검토, 최종 구조 의사결정"],
        ["PM(총괄 시니어)", "전체 구조 설계, 산출물 취합/정리, 일정·결정 관리"],
        ["콘텐츠 / IMC", "제품별 메시지·RTB, TPO 앵글, 바이럴/시딩 운영안"],
        ["디자인 / 영상", "KV·소재 배너·오프라인 비주얼 / 제품·크리에이터 영상"],
        ["퍼포먼스 마케터", "PA 집행 / 올영 인앱광고 / CRM"],
    ]))

# ============ 06. Why Us ============
divider("06", "Why Us")

content("올더베러 캠페인 구조 — 한판 요약", [
    "문제: 팔리는데 기억되지 않는 브랜드 (검색량 월 3천, 시그니처 키워드 0)",
    "전략: ‘우연 → 의도’ 각인 / 시그니처 키워드 「채움」 선점",
    "구조: SEED → PROVE → BOOST → RETAIN (시딩-콘텐츠-퍼포 연결)",
    "목표: 올영세일 카테고리 랭킹 1위 → 검색·후기·CRM 자산으로 반복 전환",
], sub="06-1 Summary")

content("BAT가 잘할 수 있는 이유", [
    "브랜드 콘텐츠 전략과 퍼포먼스 운영을 분리하지 않고, 하나의 구매 흐름 안에서 설계",
    "웰니스 카테고리의 심의/표현 리스크를 이해 (금융 등 다양한 카테고리 마이크로 IMC 경험)",
    "감도 높은 영상 크리에이티브를 세일즈로 연결",
    "우리는 이미 올리브영에서 1위를 만들어 봤다 (→ 추가 레퍼런스 제시)",
], sub="06-2 Why BAT")

# closing
s = prs.slides.add_slide(BLANK)
add_rect(s, 0, 0, SW, SH, BROWN)
tb = s.shapes.add_textbox(Inches(0.9), Inches(3.0), Inches(11.5), Inches(1.5)).text_frame
tb.word_wrap = True
r = tb.paragraphs[0].add_run(); r.text = "감사합니다."; _set_font(r, 40, True, WHITE)
r2 = tb.add_paragraph().add_run(); r2.text = "올더베러를 검색되고, 떠올려지고, 1위로 집어지는 브랜드로."
_set_font(r2, 16, False, LIGHT)

prs.save("proposal/올더베러_제안서_v1.pptx")
print("slides:", len(prs.slides._sldIdLst))
