# -*- coding: utf-8 -*-
"""올더베러 견적서 — 예시 파일 서식(맑은 고딕·그린헤더·회계서식·중앙정렬·병합) 동일 적용.
3탭: 견적제안 / 견적상세 / 월별 액션플랜. 퍼포먼스 항목 없음."""
import openpyxl
from openpyxl.styles import Font, PatternFill, Alignment, Border, Side
from openpyxl.utils import get_column_letter

FONT="맑은 고딕"
ACC='_-* #,##0_-;-* #,##0_-;_-* "-"_-;_-@_-'
DARK="404040"; GREEN_HL="82DC28"; OLIVE="62A81B"; DGREEN="082A08"; SUB="D9D9D9"; WHITE="FFFFFF"; RED="FF0000"; INK="262626"
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

# ===================================== 견적제안 (예시 프레임형 레이아웃)
ws=wb.active; ws.title="견적제안"
GRAY="F2F2F2"
# A 여백 / B 좌측띠 / C 라벨 / D 스페이서 / E 값(넓게) / F 우측띠
for i,w in enumerate([2.6,5.8,23.7,3.0,96,5.8]): ws.column_dimensions[get_column_letter(i+1)].width=w
LC,VC,FL,FR=3,5,2,6
TOP=2
sections=[
 ("· 프로젝트명",[("올리브영 PB 올더베러 브랜드 빌딩 캠페인",True,False)],False),
 ("· 집행 기간",[("2026년 7월 1일 ~ 2026년 12월 31일 (6개월) · 인플루언서 시딩·바이럴 통합",False,False)],False),
 ("· 집행 내용",[("① 인플루언서 시딩(나노·마이크로 코어) 중심 진성 UGC 대량 확산 — ‘선택 증거’ 확보",False,False),
            ("② 바이럴(커뮤니티·파워페이지·챌린저스·어필리에이트·GEO/AEO) — 구매 직전 ‘확신’·올영 랭킹 1위 견인",False,False),
            ("③ 위닝 콘텐츠·시딩 자산을 올영세일 구매 전환으로 직접 연결",False,False)],False),
 ("· 견적 상세",[("· 총 견적 : KRW 249,450,000원 (VAT 별도)   ※ 마크업 포함, 2.5억 내",False,False),
            ("· 크리에이터·집행 원고료 : 244,450,000원",False,False),
            ("· 예상 마크업 : 5,000,000원 (커뮤니티·챌린저스 항목 한정 20%)",False,False),
            ("· 목표 시딩 수량 : IG·TikTok·YT·Blog 등 총 523건 (+ EGC·Half EGC 18건)",False,False)],False),
 ("· A/C 견적 상세",[("· 크리에이터 시딩 : 186,750,000원  (마이크로 209 · 나노 209 · 매크로 15 · KOL 60 · X 15 · 블로그 15)",False,True),
              ("· 콘텐츠 EGC : 19,500,000원  (EGC 12 × 100만 · Half EGC 6 × 125만)",False,True),
              ("· 콘텐츠·바이럴 : 38,200,000원  (파워페이지 9 · 커뮤니티 체험단 3 · 챌린저스 2회 · 어필리에이트 60)",False,True),
              ("· 마크업 (커뮤니티·챌린저스 20%) : 5,000,000원",False,True),
              ("· 총 합계 : 249,450,000원 (VAT 별도)",True,True)],True),
 ("· 원고료/마크업 안내",[("· 마크업은 ‘커뮤니티 체험단’·‘챌린저스’ 2개 항목에만 20% 적용",False,False),
                ("· 그 외 크리에이터 시딩·바이럴·어필리에이트는 원고료(집행 단가) 기준",False,False),
                ("· 나노 25만 / 마이크로 50만 / 매크로 80만 (IG·TikTok 동일 단가)",False,False),
                ("· EGC = 마이크로 ×2 (100만) / Half EGC = 마이크로 ×2.5 (125만) — 별도 항목",False,False)],False),
 ("· 콘텐츠 제작 안내",[("· KV·디자인·영상 등 별도 제작비 미계상 / EGC·Half EGC는 마이크로 시딩 기준 별도 산정 항목으로 포함",False,False),
                ("· 마이크로·나노는 5:5 비율 · 퍼포먼스 항목 없음",False,False)],False),
 ("· 무상 지원 내역",[("· [대행 기간 솔루션 무상 지원] 대행 수행 시 2개 솔루션 이용료를 대행 종료시점까지 무상 지원 예정",False,False),
               ("- 스프레이ai 솔루션 월 이용료 420만원(브랜드 1개 기준)  ※ 연간 5,040만원 상당","RED",False),
               ("- 피처링 스탠다드형 솔루션 (연간 이용료 378만원 상당)",False,False),
               ("· 고성과 리포트 대시보드 / 올리비아 AI 콘텐츠·리포팅 어시스턴트",False,False),
               ("· VOC 딥다이브 리포트 / Claude AI 심의 사전검수 / 월간 성과 리포팅",False,False)],False),
 ("· 특이사항",[("· 시장 단가 변동 및 모집 기간에 따라 총 원고료는 소폭 변동될 수 있습니다.",False,False),
            ("· 캠페인별 상세 전략·일정은 ‘견적상세’ 및 ‘월별 액션플랜’ 시트를 확인 바랍니다.",False,False)],False),
]
r=TOP+1
for label,vals,shadeblk in sections:
    r0=r
    for i,(txt,big,shade) in enumerate(vals):
        red = big=="RED"
        cc=C(ws,r,VC,txt,b=(big is True or red),al="left",sz=(18 if big is True else 10),
             col=(RED if red else INK),border=False)
        if shade: cc.fill=PatternFill("solid",fgColor=GRAY)
        ws.row_dimensions[r].height=(26 if big is True else 18)
        r+=1
    # label (C) at section top, vertical-center across the block
    C(ws,r0,LC,label,b=True,al="left",sz=10,border=False)
    if r-r0>1: mrg(ws,r0,LC,r-1,LC)
    r+=1   # blank spacer row between sections
BOT=r
# ---- 프레임/띠/테두리 ----
for rr in range(TOP,BOT+1):
    for cc in (FL,FR):
        ws.cell(rr,cc).fill=PatternFill("solid",fgColor=GRAY)
# C 라벨열 세로 박스 (좌우 테두리 전 구간) + 값열 좌측 살짝
for rr in range(TOP,BOT+1):
    ws.cell(rr,LC).border=Border(left=Side(style="hair",color="D9D9D9"),right=Side(style="hair",color="D9D9D9"))
# 상/하단 띠
for cc in range(FL,FR+1):
    ws.cell(TOP,cc).fill=PatternFill("solid",fgColor=GRAY)
    ws.cell(BOT,cc).fill=PatternFill("solid",fgColor=GRAY)
ws.row_dimensions[TOP].height=10; ws.row_dimensions[BOT].height=10
# 외곽 테두리 프레임 (B..F, TOP..BOT)
med=Side(style="thin",color="808080")
for rr in range(TOP,BOT+1):
    lc=ws.cell(rr,FL); rc=ws.cell(rr,FR)
    lc.border=Border(left=med, top=(med if rr==TOP else None), bottom=(med if rr==BOT else None))
    rc.border=Border(right=med, top=(med if rr==TOP else None), bottom=(med if rr==BOT else None))
for cc in range(FL,FR+1):
    tcell=ws.cell(TOP,cc); bcell=ws.cell(BOT,cc)
    tb=tcell.border; bb=bcell.border
    tcell.border=Border(top=med,left=tb.left,right=tb.right)
    bcell.border=Border(bottom=med,left=bb.left,right=bb.right)
# 문서 타이틀(상단 띠 위) — 작게
C(ws,TOP,LC,"BAT  |  견적 제안서",b=True,sz=9,al="left",col="808080",border=False)

# ===================================== 견적상세
ds=wb.create_sheet("견적상세")
cols=[("구분",11),("제품 / 대상",16),("기간",9),("채널",12),("캠페인 및 활용 목적",38),
      ("타겟 페르소나",24),("필수 해시태그",24),("목표수량",8),("등급",9),
      ("기본원고료",12),("예상원고료",13),("마크업(20%)",12),("총비용",14)]
for i,(h,w) in enumerate(cols): ds.column_dimensions[get_column_letter(i+1)].width=w
C(ds,1,1,"올더베러 26년 하반기 인플루언서·바이럴 견적 상세   (단위: 원 / VAT 별도)",b=True,sz=13,al="left",border=False)
C(ds,2,1,"※ 마크업은 커뮤니티 체험단·챌린저스 항목에만 20% 적용 · 콘텐츠 제작비는 시딩(마이크로·나노)으로 재배분 · 퍼포먼스 항목 없음",col="808080",sz=9,al="left",border=False)
# 헤더 (예시 색상)
hr=4
hdrfill={8:GREEN_HL,11:GREEN_HL,10:OLIVE,12:OLIVE,13:DGREEN}
hdrtext={8:INK,11:INK,10:WHITE,12:WHITE,13:WHITE}
for i,(h,w) in enumerate(cols,1):
    C(ds,hr,i,h,b=True,col=hdrtext.get(i,WHITE),fill=hdrfill.get(i,DARK),al="center",wrap=True)
ds.row_dimensions[hr].height=30
ROWS=[
("크리에이터 시딩","마이크로 (UGC)","7~12월","IG·TikTok",
 "진성 UGC 대량 확산으로 ‘선택 증거’ 확보 · 위닝 메시지 발굴 및 2차 콘텐츠 활용",
 "직장인 루틴형 2535 — 웰니스 루틴 보유·편의성 중시","#올더베러 #오늘도베러 #데일리웰니스",209,"마이크로",500000,False),
("크리에이터 시딩","나노","7~12월","IG",
 "다수 진성 후기로 ‘익숙함·신뢰’ 형성 · 검색·해시태그 자산 축적(#오늘도베러)",
 "웰니스 입문형 2030 — 선택 장벽·가격 민감","#올더베러 #오늘도베러 #웰니스입문",209,"나노",250000,False),
("크리에이터 시딩","매크로","8·11·12월","IG·TikTok·YT",
 "도달·화제성 점화 — 올영세일/블프 대세감 형성","헬시플레저 라이프 4050 포함 광역","#올더베러 #올영세일 #웰니스대세",15,"매크로",800000,False),
("크리에이터 시딩","KOL 무가시딩","7~12월","IG·YT",
 "유기적 관계 기반 자발 콘텐츠·장기 팬덤 (성과 시 유상 전환)","30대 직장인·웰니스 마이크로~매크로","#올더베러 #오늘도베러",60,"무가",100000,False),
("크리에이터 시딩","X (트위터) 시딩","8·11·12월","X",
 "실시간 트렌드·밈 결합 바이럴 확산","리뷰·밈 특화 2030","#올더베러 #올영템",15,"-",500000,False),
("크리에이터 시딩","네이버 블로그","8·11·12월","Blog",
 "검색 후기·정보 탐색 대응 — 상위 노출 점유","정보성 리뷰 블로거","#올더베러 #웰니스구미추천",15,"-",300000,False),
("콘텐츠 (EGC)","EGC 콘텐츠","7~12월","IG·TikTok",
 "직원 피셜 신뢰형 — 정보·전문성 기반 미디어커머스 콘텐츠 (마이크로 ×2 산정)",
 "정보 신뢰 중시 구매 고려층","#올더베러 #올영PB #직원피셜",12,"-",1000000,False),
("콘텐츠 (EGC)","Half EGC 콘텐츠","7~12월","IG·TikTok",
 "건기식 Half-EGC — 인정 기능성 범위 내 정보+공감(소셜 에비던스, 마이크로 ×2.5 산정)",
 "건기식 관심 직장인","#올더베러 #웰니스루틴 #건강기능식품",6,"-",1250000,False),
("콘텐츠·바이럴","파워페이지","8·11·12월","SNS",
 "SNS 정보성 추천 맥락 진입 — 알고리즘·검색 탭 노출","구매 고의도·‘웰니스 구미 추천’ 검색자","#올영추천템 #웰니스추천",9,"-",800000,False),
("콘텐츠·바이럴","커뮤니티 체험단","8·11·12월","뷰티앱·커뮤니티",
 "‘검증된 후기’ 볼륨·어워드·랭킹 참여로 신뢰 확보","화해 등 카테고리 관여 高","#올리브영 #웰니스어워드",3,"-",5000000,True),
("콘텐츠·바이럴","챌린저스","8·11월","챌린저스 앱",
 "실구매·리뷰 인증 챌린지로 올영 카테고리 랭킹 1위 견인","올영 액티브 유저(실구매 의사)","#올영1위 #올더베러챌린지",2,"-(회)",5000000,True),
("콘텐츠·바이럴","어필리에이트","7~12월","올영 쇼핑 큐레이터",
 "올영 쇼핑 큐레이터 허브 — 앱 내 발견→구매 직결·세일 매출 견인","세일즈 경험 크리에이터","#올영세일 #장바구니",60,"건당",100000,False),
]
r=hr+1; first=r; groups={}
for (grp,prod,term,ch,obj,per,tag,qty,grade,price,mk) in ROWS:
    groups.setdefault(grp,[]).append(r)
    C(ds,r,1,"",border=True)  # 구분(병합용)
    C(ds,r,2,prod,b=True,wrap=True); C(ds,r,3,term); C(ds,r,4,ch,wrap=True)
    C(ds,r,5,obj,al="left",wrap=True,sz=9); C(ds,r,6,per,al="left",wrap=True,sz=9); C(ds,r,7,tag,al="left",wrap=True,sz=9,col=DGREEN)
    C(ds,r,8,qty,b=True); C(ds,r,9,grade,sz=9)
    C(ds,r,10,price,fmt=ACC); C(ds,r,11,f"=H{r}*J{r}",fmt=ACC,b=True)
    C(ds,r,12,(f"=K{r}*0.2" if mk else 0),fmt=ACC,b=mk,col=(DGREEN if mk else INK))
    C(ds,r,13,f"=K{r}+L{r}",fmt=ACC,b=True)
    ds.row_dimensions[r].height=44
    r+=1
last=r-1
# 구분 병합 + 라벨
for i,(grp,rows) in enumerate(groups.items(),1):
    mrg(ds,rows[0],1,rows[-1],1)
    C(ds,rows[0],1,f"{i}. {grp}",b=True,fill=SUB,wrap=True)
# 합계행
C(ds,r,1,"합  계",b=True,fill=SUB); mrg(ds,r,1,r,7)
for c in range(1,8): C(ds,r,c,fill=SUB,border=True)
C(ds,r,1,"합  계",b=True,fill=SUB)
C(ds,r,8,f"=SUM(H{first}:H{last})",b=True,fill=SUB,fmt=ACC)
C(ds,r,9,"",fill=SUB); C(ds,r,10,"",fill=SUB)
C(ds,r,11,f"=SUM(K{first}:K{last})",b=True,fill=SUB,fmt=ACC)
C(ds,r,12,f"=SUM(L{first}:L{last})",b=True,fill=SUB,fmt=ACC,col=DGREEN)
C(ds,r,13,f"=SUM(M{first}:M{last})",b=True,fill=GREEN_HL,fmt=ACC)
tot=r
C(ds,r+2,11,"부가세 (10%)",b=True,al="right",border=False); C(ds,r+2,13,f"=M{tot}*0.1",fmt=ACC,b=True)
C(ds,r+3,11,"총 합계 (VAT 포함)",b=True,al="right",col=WHITE,fill=DARK); C(ds,r+3,13,f"=M{tot}*1.1",fmt=ACC,b=True,col=WHITE,fill=DGREEN)
C(ds,r+5,11,"2.5억 대비 여유(VAT별도)",al="right",sz=9,border=False); C(ds,r+5,13,f"=250000000-M{tot}",fmt=ACC,sz=9,col=DGREEN,al="right")
C(ds,r+6,11,"예상 마크업율",al="right",sz=9,border=False); C(ds,r+6,13,f"=L{tot}/M{tot}",fmt="0.0%",sz=9,col=DGREEN,al="right")

# ===================================== 월별 액션플랜
ms=wb.create_sheet("월별 액션플랜")
for i,w in enumerate([22,13,8.5,11,9.5,8.5,11,11,9,15]): ms.column_dimensions[get_column_letter(i+1)].width=w
C(ms,1,1,"월별 실행 타임라인 · 액션플랜  (Moonshot Rocket Launch)",b=True,sz=14,al="left",border=False)
C(ms,2,1,"Phase 1 (7~9월) 초기 반응·선택증거 확보 → Phase 2 (10~12월) 제품군 확장·재점화 · 단가/금액은 견적상세와 동일",col="808080",sz=9,al="left",border=False)
mh=["항목","단가","7월","8월\n세일사전","9월\n세일★","10월","11월\n블프★","12월\n세일★","합계","금액"]
mf={9:GREEN_HL,10:DGREEN}
mt={9:INK,10:WHITE}
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
mrow(6,"마이크로 (UGC)",500000,[25,50,30,25,35,44])
mrow(7,"나노",250000,[25,50,30,25,35,44])
mrow(8,"매크로",800000,[0,5,0,0,5,5])
mrow(9,"KOL 무가시딩",100000,[10,10,10,10,10,10])
mrow(10,"X (트위터) 시딩",500000,[0,5,0,0,5,5])
mrow(11,"네이버 블로그",300000,[0,5,0,0,5,5])
msec(12,"■ 콘텐츠 (EGC)")
mrow(13,"EGC 콘텐츠 (마이크로×2)",1000000,[2,2,2,2,2,2])
mrow(14,"Half EGC 콘텐츠 (마이크로×2.5)",1250000,[1,1,1,1,1,1])
msec(15,"■ 콘텐츠 · 바이럴")
mrow(16,"파워페이지",800000,[0,3,0,0,3,3])
mrow(17,"커뮤니티 체험단 (마크업20%)",5000000,[0,1,0,0,1,1],mk=True)
mrow(18,"챌린저스 (마크업20%)",5000000,[0,1,0,0,1,0],mk=True)
mrow(19,"어필리에이트 (누적)",100000,[10,20,30,40,50,60],cum=True)
C(ms,21,1,"월 시딩 총건수",b=True,fill=SUB,al="left")
for i in range(6):
    cl=get_column_letter(3+i); C(ms,21,3+i,f"=SUM({cl}6:{cl}11)",b=True,fill=SUB)
C(ms,21,2,"",fill=SUB); C(ms,21,9,"=SUM(C21:H21)",b=True,fill=SUB); C(ms,21,10,"",fill=SUB)
C(ms,22,9,"총 견적",b=True,al="right"); C(ms,22,10,"=SUM(J6:J19)",fmt=ACC,b=True,fill=GREEN_HL)
for r in range(5,20): ms.row_dimensions[r].height=19

wb.save("proposal/올더베러_견적서_BAT.xlsx")
print("saved · tot row:",tot)
