# -*- coding: utf-8 -*-
"""올더베러 견적서 — 예시 폼 서식 + 월별 블록 구조(풍성).
견적제안 / 견적상세(월별 블록 + 무상 하단) / 월별 액션플랜.
규칙: 마이크로·나노 5:5(197/197) · EGC=마이크로×2 / HalfEGC=×2.5 · 콘텐츠제작 월150만×6
· 커뮤니티·챌린저스만 마크업20% · 총 249,450,000(2.5억 내) · 퍼포먼스 없음."""
import openpyxl
from openpyxl.styles import Font, PatternFill, Alignment, Border, Side
from openpyxl.utils import get_column_letter

FONT="맑은 고딕"
ACC='_-* #,##0_-;-* #,##0_-;_-* "-"_-;_-@_-'
DARK="404040"; GREEN_HL="82DC28"; OLIVE="62A81B"; DGREEN="082A08"; SUB="D9D9D9"; LGRN="EAF3DA"; WHITE="FFFFFF"; RED="FF0000"; INK="262626"
thin=Side(style="thin",color="BFBFBF"); bd=Border(left=thin,right=thin,top=thin,bottom=thin)
def C(ws,r,c,v=None,b=False,col=INK,fill=None,al="center",fmt=None,sz=10,wrap=False,border=True):
    cc=ws.cell(row=r,column=c)
    if v is not None: cc.value=v
    cc.font=Font(name=FONT,size=sz,bold=b,color=col)
    if fill: cc.fill=PatternFill("solid",fgColor=fill)
    cc.alignment=Alignment(horizontal=al,vertical="center",wrap_text=wrap)
    if border: cc.border=bd
    if fmt: cc.number_format=fmt
    return cc
def mrg(ws,r1,c1,r2,c2): ws.merge_cells(start_row=r1,start_column=c1,end_row=r2,end_column=c2)
wb=openpyxl.Workbook()

# ===================================== 견적제안 (프레임형)
ws=wb.active; ws.title="견적제안"
GRAY="F2F2F2"
for i,w in enumerate([2.6,5.8,23.7,3.0,96,5.8]): ws.column_dimensions[get_column_letter(i+1)].width=w
LC,VC,FL,FR=3,5,2,6; TOP=2
sections=[
 ("· 프로젝트명",[("올리브영 PB 올더베러 브랜드 빌딩 캠페인",True,False)]),
 ("· 집행 기간",[("2026년 7월 1일 ~ 2026년 12월 31일 (6개월) · 인플루언서 시딩·바이럴 통합",False,False)]),
 ("· 집행 내용",[("① 인플루언서 시딩(나노·마이크로 코어) 중심 진성 UGC 대량 확산 — ‘선택 증거’ 확보",False,False),
            ("② 바이럴(커뮤니티·파워페이지·챌린저스·어필리에이트·GEO/AEO) — 구매 직전 ‘확신’·올영 랭킹 1위 견인",False,False),
            ("③ 위닝 콘텐츠·시딩 자산을 올영세일 구매 전환으로 직접 연결",False,False)]),
 ("· 견적 상세",[("· 총 견적 : KRW 249,450,000원 (VAT 별도)   ※ 마크업 포함, 2.5억 내",False,False),
            ("· 크리에이터·집행 원고료 : 244,450,000원 / 예상 마크업 : 5,000,000원 (커뮤니티·챌린저스 한정 20%)",False,False),
            ("· 목표 시딩 수량 : IG·TikTok·YT·Blog 등 총 499건 (+ EGC·Half EGC 18건)",False,False)]),
 ("· A/C 견적 상세",[("· 크리에이터 시딩 : 177,750,000원  (마이크로 197 · 나노 197 · 매크로 15 · KOL 60 · X 15 · 블로그 15)",False,True),
              ("· 콘텐츠 EGC : 19,500,000원  (EGC 12 × 100만 · Half EGC 6 × 125만)",False,True),
              ("· 콘텐츠 제작(KV·영상·이미지) : 9,000,000원  (월 150만 × 6개월)",False,True),
              ("· 콘텐츠·바이럴 : 38,200,000원  (파워페이지 9 · 커뮤니티 3 · 챌린저스 2회 · 어필리에이트 60)",False,True),
              ("· 마크업 (커뮤니티·챌린저스 20%) : 5,000,000원",False,True),
              ("· 총 합계 : 249,450,000원 (VAT 별도)",True,True)]),
 ("· 원고료/마크업 안내",[("· 마크업은 ‘커뮤니티 체험단’·‘챌린저스’ 2개 항목에만 20% 적용",False,False),
                ("· 나노 25만 / 마이크로 50만 / 매크로 80만 (IG·TikTok 동일 단가)",False,False),
                ("· EGC = 마이크로 ×2(100만) / Half EGC = 마이크로 ×2.5(125만) · 콘텐츠 제작 월 150만",False,False)]),
 ("· 콘텐츠 제작 안내",[("· 콘텐츠 제작(KV·디자인·영상·이미지 소재) 월 150만 × 6개월 별도 계상 · 마이크로·나노 5:5",False,False),
                ("· UGC·EGC·Half EGC 영상은 시딩 원고료 내 제작·확보 · 퍼포먼스 항목 없음",False,False)]),
 ("· 무상 지원 내역",[("· [대행 기간 솔루션 무상 지원] 대행 수행 시 2개 솔루션 이용료를 대행 종료시점까지 무상 지원 예정",False,False),
               ("- 스프레이ai 솔루션 월 이용료 420만원(브랜드 1개 기준)  ※ 연간 5,040만원 상당","RED",False),
               ("- 피처링 스탠다드형 솔루션 (연간 이용료 378만원 상당)",False,False),
               ("· 고성과 리포트 대시보드 / 올리비아 AI 어시스턴트 / VOC 딥다이브 / Claude AI 심의 검수 / 월간 리포팅",False,False)]),
 ("· 특이사항",[("· 시장 단가 변동 및 모집 기간에 따라 총 원고료는 소폭 변동될 수 있습니다.",False,False),
            ("· 항목별 상세는 ‘견적상세’(월별), 집행 일정은 ‘월별 액션플랜’ 시트를 확인 바랍니다.",False,False)]),
]
r=TOP+1
for label,vals in sections:
    r0=r
    for (txt,big,shade) in vals:
        red=(big=="RED")
        cc=C(ws,r,VC,txt,b=(big is True or red),al="left",sz=(18 if big is True else 10),col=(RED if red else INK),border=False)
        if shade: cc.fill=PatternFill("solid",fgColor=GRAY)
        ws.row_dimensions[r].height=(26 if big is True else 18); r+=1
    C(ws,r0,LC,label,b=True,al="left",border=False)
    if r-r0>1: mrg(ws,r0,LC,r-1,LC)
    r+=1
BOT=r
for rr in range(TOP,BOT+1):
    for cc in (FL,FR): ws.cell(rr,cc).fill=PatternFill("solid",fgColor=GRAY)
    ws.cell(rr,LC).border=Border(left=Side(style="hair",color="D9D9D9"),right=Side(style="hair",color="D9D9D9"))
for cc in range(FL,FR+1):
    ws.cell(TOP,cc).fill=PatternFill("solid",fgColor=GRAY); ws.cell(BOT,cc).fill=PatternFill("solid",fgColor=GRAY)
ws.row_dimensions[TOP].height=10; ws.row_dimensions[BOT].height=10
med=Side(style="thin",color="808080")
for rr in range(TOP,BOT+1):
    ws.cell(rr,FL).border=Border(left=med,top=(med if rr==TOP else None),bottom=(med if rr==BOT else None))
    ws.cell(rr,FR).border=Border(right=med,top=(med if rr==TOP else None),bottom=(med if rr==BOT else None))
for cc in range(FL,FR+1):
    ws.cell(TOP,cc).border=Border(top=med,left=ws.cell(TOP,cc).border.left,right=ws.cell(TOP,cc).border.right)
    ws.cell(BOT,cc).border=Border(bottom=med,left=ws.cell(BOT,cc).border.left,right=ws.cell(BOT,cc).border.right)
C(ws,TOP,LC,"BAT  |  견적 제안서",b=True,sz=9,al="left",col="808080",border=False)

# ===================================== 견적상세 (월별 블록)
ds=wb.create_sheet("견적상세")
cols=[("항목",16),("채널",11),("캠페인 및 활용 목적 / 앵글",36),("타겟 페르소나",22),
      ("필수 해시태그",22),("목표수량",8),("단가",11),("마크업(20%)",11),("총비용",13)]
for i,(h,w) in enumerate(cols): ds.column_dimensions[get_column_letter(i+1)].width=w
C(ds,1,1,"올더베러 26년 하반기 견적 상세 — 월별 캠페인   (단위: 원 / VAT 별도)",b=True,sz=13,al="left",border=False)
C(ds,2,1,"※ Moonshot Rocket Launch · 마크업은 커뮤니티·챌린저스 항목에만 20% · 퍼포먼스 항목 없음 · 단가/금액은 월별 액션플랜과 동일",col="808080",sz=9,al="left",border=False)
hr=4
hf={6:GREEN_HL,7:OLIVE,8:OLIVE,9:DGREEN}; ht={6:INK,7:WHITE,8:WHITE,9:WHITE}
for i,(h,w) in enumerate(cols,1): C(ds,hr,i,h,b=True,col=ht.get(i,WHITE),fill=hf.get(i,DARK),al="center",wrap=True)
ds.row_dimensions[hr].height=28

META={
"마이크로 (UGC)":("IG·TT","직장인 루틴형 2535","#올더베러 #오늘도베러",500000,False),
"나노":("IG","웰니스 입문형 2030","#올더베러 #오늘도베러 #웰니스입문",250000,False),
"매크로":("IG·TT·YT","헬시플레저 광역 4050+","#올더베러 #올영세일",800000,False),
"KOL 무가시딩":("IG·YT","30대 웰니스 KOL","#올더베러 #오늘도베러",100000,False),
"X (트위터) 시딩":("X","리뷰·밈 2030","#올더베러 #올영템",500000,False),
"네이버 블로그":("Blog","정보탐색 검색자","#올더베러 #웰니스구미추천",300000,False),
"EGC 콘텐츠":("IG·TT","정보 신뢰 구매고려층","#올더베러 #직원피셜",1000000,False),
"Half EGC 콘텐츠":("IG·TT","건기식 관심 직장인","#올더베러 #웰니스루틴",1250000,False),
"파워페이지":("SNS","구매 고의도 검색자","#올영추천템 #웰니스추천",800000,False),
"커뮤니티 체험단":("뷰티앱·커뮤니티","카테고리 관여 高 사용자","#올리브영 #웰니스어워드",5000000,True),
"챌린저스":("챌린저스 앱","올영 액티브 유저","#올영1위 #올더베러챌린지",5000000,True),
"어필리에이트":("올영 쇼핑 큐레이터","세일즈 경험 크리에이터","#올영세일 #장바구니",100000,False),
"콘텐츠 제작 (KV·영상·이미지)":("제작","—","—",1500000,False),
}
MONTHS=[
("7월 · 착수 / 첫 점화 (Phase 1)","위닝 메시지 발굴 + 일반식품 시딩 도화선 점화 · 타겟: 웰니스 입문형",[
  ("마이크로 (UGC)",24,"멜라나잇(자기 전)·올리브오일(아침 공복) 위닝 메시지 LMF 테스트"),
  ("나노",24,"진성 후기 볼륨 착수 · #오늘도베러 해시태그 자산화"),
  ("KOL 무가시딩",10,"웰니스 코어 KOL 관계 형성·자발 콘텐츠"),
  ("EGC 콘텐츠",2,"직원 피셜 — 올리브오일 ‘이 스펙 이 가격’ 신뢰형"),
  ("Half EGC 콘텐츠",1,"건기식 정보+공감 (루테인 모니터 앞 직장인)"),
  ("어필리에이트",10,"올영 쇼핑 큐레이터 풀 구축·앱 내 동선 세팅"),
  ("콘텐츠 제작 (KV·영상·이미지)",1,"브랜드 KV·이미지·영상 기본 소재 제작"),
]),
("8월 · 세일 사전 부스팅 (1차 웨이브)","올영세일 前 인지·각인 확산 + 선택 증거 확보 · 타겟: 입문형+루틴형",[
  ("마이크로 (UGC)",47,"위닝 메시지 대량 확산 — 세일 1주 전 집중"),
  ("나노",47,"다수 진성 후기로 ‘많이 보이는 제품’ 인식 형성"),
  ("매크로",5,"세일 대세감 점화·도달 확대"),
  ("KOL 무가시딩",10,"오가닉 콘텐츠 확대"),
  ("X (트위터) 시딩",5,"실시간 트렌드·밈 바이럴"),
  ("네이버 블로그",5,"검색 후기·정보 탐색 상위 노출"),
  ("EGC 콘텐츠",2,"제품 RTB 신뢰형 콘텐츠"),
  ("Half EGC 콘텐츠",1,"건기식 인정 기능성 정보+공감"),
  ("파워페이지",3,"‘추천템’ 맥락·정보성 노출"),
  ("커뮤니티 체험단",1,"체험단·어워드로 ‘검증된 후기’ 볼륨"),
  ("챌린저스",1,"구매·리뷰 인증 챌린지 가동(랭킹 빌드업)"),
  ("어필리에이트",10,"앱 내 발견→구매 동선 확대"),
  ("콘텐츠 제작 (KV·영상·이미지)",1,"세일 소재·B&A 컷 제작"),
]),
("9월 · 올영세일 1차 피크","확보한 증거로 구매 전환 + 카테고리 랭킹 상단 · 타겟: 실질 전환형",[
  ("마이크로 (UGC)",28,"세일 전환 소구 + ‘올영 1위’ 랭킹 결합 소재"),
  ("나노",28,"구매 인증·리뷰 볼륨 집중"),
  ("KOL 무가시딩",10,"전환 시점 자발 후기"),
  ("EGC 콘텐츠",2,"세일 클로징 신뢰 소재"),
  ("Half EGC 콘텐츠",1,"건기식 전환 정보 콘텐츠"),
  ("어필리에이트",10,"세일 매출·랭킹 견인"),
  ("콘텐츠 제작 (KV·영상·이미지)",1,"세일 본편·마지막날 소재"),
]),
("10월 · 유지 / 건기식 라인 확장","건기식 라인 확장 + 위닝 소재 낙수 · 타겟: 직장인 루틴형",[
  ("마이크로 (UGC)",24,"건기식(루테인·바나바잎·올인원) 확장 메시지 테스트"),
  ("나노",24,"9월 위닝 소재 낙수·후기 자산 확산"),
  ("KOL 무가시딩",10,"관계형 콘텐츠 유지"),
  ("EGC 콘텐츠",2,"건기식 라인 신뢰 콘텐츠"),
  ("Half EGC 콘텐츠",1,"건기식 정보+공감 확장"),
  ("어필리에이트",10,"상시 전환 동선 유지"),
  ("콘텐츠 제작 (KV·영상·이미지)",1,"건기식 라인 소재 제작"),
]),
("11월 · 블프 (2차 웨이브)","재점화 — 많이 보이고 많이 찾는 브랜드 · 타겟: 입문형+루틴형",[
  ("마이크로 (UGC)",33,"블프 재점화·위닝 소재 재투입"),
  ("나노",33,"진성 후기 재확대·대세감"),
  ("매크로",5,"블프 대세감 점화"),
  ("KOL 무가시딩",10,"오가닉 재가동"),
  ("X (트위터) 시딩",5,"실시간 트렌드 바이럴"),
  ("네이버 블로그",5,"검색 상위·추천 점유"),
  ("EGC 콘텐츠",2,"RTB+혜택 결합 소재"),
  ("Half EGC 콘텐츠",1,"건기식 정보 콘텐츠"),
  ("파워페이지",3,"추천 맥락 재노출"),
  ("커뮤니티 체험단",1,"검증 후기 보강"),
  ("챌린저스",1,"챌린저스 2차·신뢰 자산 보강"),
  ("어필리에이트",10,"블프 전환 견인"),
  ("콘텐츠 제작 (KV·영상·이미지)",1,"블프 프로모션 소재"),
]),
("12월 · 올영세일 2차 피크 / 대표성 착지","웰니스 카테고리 대표 브랜드 착지 · 타겟: 실질 전환형",[
  ("마이크로 (UGC)",41,"전 라인업 대표성 콘텐츠·세일 클로징"),
  ("나노",41,"구매 인증·쟁여두기 소구"),
  ("매크로",5,"세일 대세감·대표성"),
  ("KOL 무가시딩",10,"연말 자발 후기"),
  ("X (트위터) 시딩",5,"실시간 바이럴"),
  ("네이버 블로그",5,"검색 점유 유지"),
  ("EGC 콘텐츠",2,"대표성 신뢰 소재"),
  ("Half EGC 콘텐츠",1,"건기식 정보 콘텐츠"),
  ("파워페이지",3,"추천 맥락 노출"),
  ("커뮤니티 체험단",1,"연말 어워드·랭킹"),
  ("어필리에이트",10,"세일 매출·랭킹 클로징"),
  ("콘텐츠 제작 (KV·영상·이미지)",1,"세일 클로징 소재"),
]),
]
r=hr+1; subtotal_rows=[]
for title,goal,rows in MONTHS:
    # 월 배너
    mrg(ds,r,1,r,9); C(ds,r,1,f"  ▣ {title}   —   {goal}",b=True,col=DGREEN,fill=LGRN,al="left")
    ds.row_dimensions[r].height=22; r+=1
    first=r
    for (item,qty,angle) in rows:
        ch,per,tag,price,mk=META[item]
        C(ds,r,1,item,b=True,al="left",wrap=True); C(ds,r,2,ch); C(ds,r,3,angle,al="left",wrap=True,sz=9)
        C(ds,r,4,per,al="left",wrap=True,sz=9); C(ds,r,5,tag,al="left",wrap=True,sz=9,col=DGREEN)
        C(ds,r,6,qty,b=True); C(ds,r,7,price,fmt=ACC)
        C(ds,r,8,(f"=G{r}*F{r}*0.2" if mk else 0),fmt=ACC,col=(DGREEN if mk else INK))
        C(ds,r,9,(f"=G{r}*F{r}*1.2" if mk else f"=G{r}*F{r}"),fmt=ACC,b=True)
        ds.row_dimensions[r].height=30; r+=1
    last=r-1
    # 월 소계
    mrg(ds,r,1,r,5); C(ds,r,1,f"{title.split(' · ')[0]} 소계",b=True,fill=GREEN_HL,al="right")
    C(ds,r,6,f"=SUM(F{first}:F{last})",b=True,fill=GREEN_HL)
    C(ds,r,7,"",fill=GREEN_HL); C(ds,r,8,f"=SUM(H{first}:H{last})",fmt=ACC,b=True,fill=GREEN_HL)
    C(ds,r,9,f"=SUM(I{first}:I{last})",fmt=ACC,b=True,fill=GREEN_HL)
    subtotal_rows.append(r); r+=1
# 총계
mrg(ds,r,1,r,5); C(ds,r,1,"총 합계 (원고료+마크업, VAT 별도)",b=True,col=WHITE,fill=DARK,al="right")
for c in (6,7): C(ds,r,c,fill=DARK)
C(ds,r,8,"="+"+".join(f"H{x}" for x in subtotal_rows),fmt=ACC,b=True,col=WHITE,fill=DARK)
C(ds,r,9,"="+"+".join(f"I{x}" for x in subtotal_rows),fmt=ACC,b=True,col=WHITE,fill=DGREEN)
grand=r; r+=1
C(ds,r,8,"부가세(10%)",b=True,al="right"); C(ds,r,9,f"=I{grand}*0.1",fmt=ACC,b=True); r+=1
C(ds,r,8,"합계(VAT 포함)",b=True,al="right",col=WHITE,fill=DARK); C(ds,r,9,f"=I{grand}*1.1",fmt=ACC,b=True,col=WHITE,fill=DARK); r+=2

# 무상 지원 (하단)
C(ds,r,1,"■ 무상 지원 (Value-add · 0원)",b=True,col=DGREEN,fill=LGRN,al="left"); mrg(ds,r,1,r,9); r+=1
free=[("스프레이ai 솔루션","대행 종료시점까지 무상 — 월 420만(브랜드 1개) · 연 5,040만원 상당"),
      ("피처링 스탠다드형 솔루션","대행 기간 무상 — 연 378만원 상당"),
      ("고성과 리포트 대시보드","실시간 성과·랭킹 트래킹 제공"),
      ("올리비아 (AI 어시스턴트)","AI 콘텐츠·리포팅 어시스턴트 제공"),
      ("VOC 딥다이브 리포트","리뷰 크롤링·키워드 분석 리포트"),
      ("Claude AI 심의 사전검수","심의 반려율↓·리드타임 단축"),
      ("월간 성과 리포팅","일·주·월간 리포팅 제공"),
      ("위닝 콘텐츠 2차 활용","페이드·온드 미디어 2차 활용 라이선스")]
for nm,desc in free:
    C(ds,r,1,nm,b=True,al="left",wrap=True); mrg(ds,r,3,r,5); C(ds,r,3,desc,al="left",wrap=True,sz=9)
    C(ds,r,2,"무상",col=RED,b=True)
    C(ds,r,6,"-"); C(ds,r,7,"-"); C(ds,r,8,"-"); C(ds,r,9,0,fmt=ACC,col=DGREEN,b=True)
    r+=1

# ===================================== 월별 액션플랜
ms=wb.create_sheet("월별 액션플랜")
for i,w in enumerate([24,13,8.5,11,9.5,8.5,11,11,9,15]): ms.column_dimensions[get_column_letter(i+1)].width=w
C(ms,1,1,"월별 실행 타임라인 · 액션플랜  (Moonshot Rocket Launch)",b=True,sz=14,al="left",border=False)
C(ms,2,1,"단가/금액은 견적상세와 동일 · 합계 = 총 견적 249,450,000 (마크업 포함·VAT 별도)",col="808080",sz=9,al="left",border=False)
mh=["항목","단가","7월","8월\n세일사전","9월\n세일★","10월","11월\n블프★","12월\n세일★","합계","금액"]
mf={9:GREEN_HL,10:DGREEN}; mt={9:INK,10:WHITE}
for i,h in enumerate(mh,1): C(ms,4,i,h,b=True,col=mt.get(i,WHITE),fill=mf.get(i,DARK),al="center",wrap=True)
ms.row_dimensions[4].height=30
def msec(r,t):
    for c in range(1,11): C(ms,r,c,fill=SUB)
    C(ms,r,1,t,b=True,fill=SUB,al="left")
def mrow(r,name,unit,months,mk=False,cum=False):
    C(ms,r,1,name,al="left"); C(ms,r,2,unit,fmt=ACC)
    for i,v in enumerate(months): C(ms,r,3+i,(v if v else "-"))
    if cum:
        C(ms,r,9,"누적 60",b=True); C(ms,r,10,f"=60*B{r}",fmt=ACC,b=True)
    else:
        C(ms,r,9,f"=SUM(C{r}:H{r})",b=True)
        C(ms,r,10,(f"=I{r}*B{r}*1.2" if mk else f"=I{r}*B{r}"),fmt=ACC,b=True)
msec(5,"■ 크리에이터 시딩")
mrow(6,"마이크로 (UGC)",500000,[24,47,28,24,33,41])
mrow(7,"나노",250000,[24,47,28,24,33,41])
mrow(8,"매크로",800000,[0,5,0,0,5,5])
mrow(9,"KOL 무가시딩",100000,[10,10,10,10,10,10])
mrow(10,"X (트위터) 시딩",500000,[0,5,0,0,5,5])
mrow(11,"네이버 블로그",300000,[0,5,0,0,5,5])
msec(12,"■ 콘텐츠")
mrow(13,"EGC 콘텐츠 (마이크로×2)",1000000,[2,2,2,2,2,2])
mrow(14,"Half EGC 콘텐츠 (마이크로×2.5)",1250000,[1,1,1,1,1,1])
mrow(15,"콘텐츠 제작 (KV·영상·이미지)",1500000,[1,1,1,1,1,1])
msec(16,"■ 바이럴 · 전환")
mrow(17,"파워페이지",800000,[0,3,0,0,3,3])
mrow(18,"커뮤니티 체험단 (마크업20%)",5000000,[0,1,0,0,1,1],mk=True)
mrow(19,"챌린저스 (마크업20%)",5000000,[0,1,0,0,1,0],mk=True)
mrow(20,"어필리에이트 (누적)",100000,[10,20,30,40,50,60],cum=True)
C(ms,22,1,"월 시딩 총건수",b=True,fill=SUB,al="left")
for i in range(6):
    cl=get_column_letter(3+i); C(ms,22,3+i,f"=SUM({cl}6:{cl}11)",b=True,fill=SUB)
C(ms,22,2,"",fill=SUB); C(ms,22,9,"=SUM(C22:H22)",b=True,fill=SUB); C(ms,22,10,"",fill=SUB)
C(ms,23,9,"총 견적",b=True,al="right"); C(ms,23,10,"=SUM(J6:J20)",fmt=ACC,b=True,fill=GREEN_HL)
for rr in range(5,21): ms.row_dimensions[rr].height=19

wb.save("proposal/올더베러_견적서_BAT.xlsx")
print("saved · 견적상세 grand row:",grand)
