# -*- coding: utf-8 -*-
"""올더베러 견적서 — 공유 폼 기준(견적제안/견적상세) + 월별 액션플랜.
규칙: 콘텐츠 제작비 전액 제거→마이크로·나노 5:5 / 커뮤니티·챌린저스만 마크업20% /
총견적 2.5억 내 / 스프레이AI 무상 최상단. 페르소나·해시태그·목적은 제안 PDF 반영."""
import openpyxl
from openpyxl.styles import Font, PatternFill, Alignment, Border, Side
from openpyxl.utils import get_column_letter

BROWN="3D2B1B"; BEIGE="F3F2E7"; GREEN="0E5A30"; GRAYF="EFEFEC"; WHITE="FFFFFF"; HEAD="2B2B2B"; YEL="FBFBE0"
thin=Side(style="thin",color="CFCcC4"); bd=Border(left=thin,right=thin,top=thin,bottom=thin)
def C(ws,r,c,v=None,b=False,col="141412",fill=None,al="left",fmt=None,sz=10,wrap=False,border=True):
    cc=ws.cell(row=r,column=c)
    if v is not None: cc.value=v
    cc.font=Font(name="Pretendard",size=sz,bold=b,color=col)
    if fill: cc.fill=PatternFill("solid",fgColor=fill)
    cc.alignment=Alignment(horizontal=al,vertical="center",wrap_text=wrap)
    if border: cc.border=bd
    if fmt: cc.number_format=fmt
    return cc
NUM="#,##0"
wb=openpyxl.Workbook()

# ============================================================ 견적제안
ws=wb.active; ws.title="견적제안"
for i,w in enumerate([2,22,2,90]): ws.column_dimensions[get_column_letter(i+1)].width=w
def prop(r,label,vals,lblfill=BEIGE):
    C(ws,r,2,label,b=True,fill=lblfill,al="left",border=True)
    if isinstance(vals,str): vals=[vals]
    for i,v in enumerate(vals):
        C(ws,r+i,4,v,al="left",border=False,wrap=True)
    if len(vals)>1:
        ws.merge_cells(start_row=r,start_column=2,end_row=r+len(vals)-1,end_column=2)
C(ws,1,2,"올리브영 PB 올더베러 브랜드 빌딩 캠페인 — 견적 제안",b=True,sz=15,border=False)
prop(3,"· 프로젝트명","올리브영 PB 올더베러 브랜드 빌딩 캠페인 (인플루언서 시딩·바이럴·퍼포먼스 통합)")
prop(5,"· 집행 기간","2026년 7월 1일 ~ 2026년 12월 31일 (6개월)")
prop(7,"· 집행 내용",["① 인플루언서 시딩(나노·마이크로 코어) 중심의 진성 UGC 콘텐츠 대량 확산 — ‘선택 증거’ 확보",
                   "② 바이럴(커뮤니티·파워페이지·챌린저스·어필리에이트·GEO/AEO) — 구매 직전 ‘확신’ 형성·올영 랭킹 1위 견인",
                   "③ 위닝 콘텐츠 → 퍼포먼스(메타·올영 인앱) 2차 확장 → 올영세일 구매 전환"])
prop(11,"· 견적 상세",["· 총 견적 : KRW 249,450,000원 (VAT 별도)  ※ 마크업 포함, 2.5억 내",
                   "· 크리에이터·집행 원고료 : 244,450,000원",
                   "· 예상 마크업 : 5,000,000원 (커뮤니티·챌린저스 항목 한정 20%)",
                   "· 목표 시딩 수량 : IG·TT·YT·Blog 등 총 575건 (+ 바이럴·어필리에이트 별도)"])
prop(16,"· 원고료/마크업 안내",["· 본 견적의 마크업은 ‘커뮤니티 체험단’·‘챌린저스’ 2개 항목에만 20% 적용",
                       "· 그 외 크리에이터 시딩·바이럴·어필리에이트는 원고료(집행 단가) 기준",
                       "· 나노 25만 / 마이크로 50만 / 매크로 80만 (IG·TikTok 동일 단가)"])
prop(20,"· 콘텐츠 제작 안내",["· 별도 콘텐츠 제작비(KV·디자인·영상·EGC) 미계상 → 전액 인플루언서 시딩(마이크로·나노 5:5)으로 재배분",
                       "· UGC·Half EGC·EGC 등 영상은 크리에이터 시딩 원고료 내에서 제작·확보"])
prop(23,"· 무상 지원 내역",["· [대행 기간 솔루션 무상 지원] 대행 업무 수행 시 2개 솔루션 이용료를 대행 종료시점까지 무상 지원",
                      "   - 스프레이AI 솔루션 월 이용료 420만원(브랜드 1개 기준)  ※ 연간 5,040만원 상당",
                      "   - 피처링 스탠다드형 솔루션 (연간 이용료 378만원 상당)",
                      "· 고성과 리포트 대시보드 (실시간 성과·랭킹 트래킹)",
                      "· 올리비아 — AI 콘텐츠·리포팅 어시스턴트",
                      "· 퍼포먼스 소재 제작 지원 (위닝 소재 PA 2차 고도화)",
                      "· VOC 딥다이브 리포트 / Claude AI 심의 사전검수 / 월간 성과 리포팅"])
prop(31,"· 특이사항",["· 시장 단가 변동 및 모집 기간에 따라 총 원고료는 소폭 변동될 수 있습니다.",
                  "· 캠페인별 상세 전략·일정은 ‘견적상세’ 및 ‘월별 액션플랜’ 시트를 확인 바랍니다."])
for r in range(3,34): ws.row_dimensions[r].height=20

# ============================================================ 견적상세
ds=wb.create_sheet("견적상세")
cols=["구분","제품/대상","기간","채널","캠페인 및 활용 목적","타겟 페르소나","필수 해시태그","목표수량","등급","기본원고료","예상원고료","마크업(20%)","총비용"]
widths=[10,16,8,12,40,26,28,8,9,12,13,12,13]
for i,w in enumerate(widths): ds.column_dimensions[get_column_letter(i+1)].width=w
C(ds,1,1,"올더베러 26년 하반기 인플루언서·바이럴 견적 상세  (단위: 원 / VAT 별도)",b=True,sz=13,border=False)
C(ds,2,1,"※ 마크업은 커뮤니티 체험단·챌린저스 항목에만 20% 적용 · 콘텐츠 제작비는 시딩(마이크로·나노)으로 재배분",col="868680",sz=9,border=False)
hr=4
for c,h in enumerate(cols,1): C(ds,hr,c,h,b=True,col=WHITE,fill=BROWN,al="center",wrap=True)
ds.row_dimensions[hr].height=30
# rows: (구분,제품,기간,채널,목적,페르소나,해시태그,수량,등급,단가,markup20?)
R=[
("크리에이터 시딩","마이크로 (UGC)","7~12월","IG·TikTok",
 "진성 UGC 대량 확산으로 ‘선택 증거’ 확보 · 위닝 메시지 발굴 → 퍼포먼스(PA) 2차 소재화",
 "직장인 루틴형 2535 남녀 — 웰니스 루틴 보유·편의성 중시","#올더베러 #오늘도베러 #데일리웰니스 #올영웰니스",235,"마이크로",500000,False),
("크리에이터 시딩","나노","7~12월","IG",
 "다수 진성 후기로 ‘익숙함·신뢰’ 형성 · 검색·해시태그 자산 축적(#오늘도베러)",
 "웰니스 입문형 2030 남녀 — 선택 장벽·가격 민감","#올더베러 #오늘도베러 #웰니스입문 #내돈내산",235,"나노",250000,False),
("크리에이터 시딩","매크로","8·11·12월","IG·TikTok·YT",
 "도달·화제성 점화 — 올영세일/블프 대세감 형성","헬시플레저 라이프 4050 포함 광역 타겟","#올더베러 #올영세일 #웰니스대세",15,"매크로",800000,False),
("크리에이터 시딩","KOL 무가시딩","7~12월","IG·YT",
 "유기적 관계 기반 자발 콘텐츠·장기 팬덤 — 성과 시 유상 전환","30대 직장인·웰니스 라이프 마이크로~매크로","#올더베러 #오늘도베러",60,"무가",100000,False),
("크리에이터 시딩","X (트위터) 시딩","8·11·12월","X",
 "실시간 트렌드·밈 결합 바이럴 확산","리뷰·밈 특화 2030 크리에이터","#올더베러 #올영템 #웰니스템",15,"-",500000,False),
("크리에이터 시딩","네이버 블로그","8·11·12월","Blog",
 "검색 후기·정보 탐색 대응 — 상위 노출 점유","정보성 리뷰 블로거","#올더베러 #웰니스구미추천 #올리브오일추천",15,"-",300000,False),
("콘텐츠·바이럴","파워페이지","8·11·12월","SNS",
 "SNS 정보성 추천 맥락 진입 — 알고리즘·검색 탭 노출","구매 고의도·‘웰니스 구미 추천’ 검색자","#올영추천템 #웰니스추천 #올더베러",9,"-",800000,False),
("콘텐츠·바이럴","커뮤니티 체험단","8·11·12월","뷰티앱·커뮤니티",
 "‘검증된 후기’ 볼륨·어워드·랭킹 참여로 신뢰 확보","화해 등 카테고리 관여 高 사용자","#올리브영 #웰니스어워드 #올더베러",3,"-",5000000,True),
("콘텐츠·바이럴","챌린저스","8·11월","챌린저스 앱",
 "실구매·리뷰 인증 챌린지로 올영 카테고리 랭킹 1위 견인","올영 액티브 유저(실구매 의사)","#올영1위 #올더베러챌린지 #인증",2,"-(회)",5000000,True),
("콘텐츠·바이럴","어필리에이트","7~12월","올영 쇼핑 큐레이터",
 "올영 쇼핑 큐레이터 허브 — 앱 내 발견→구매 직결·세일 매출 견인","쿠팡파트너스·공구 등 세일즈 경험 크리에이터","#올영세일 #장바구니 #올더베러",60,"건당",100000,False),
]
r=hr+1; first=r
group_prev=None
for (grp,prod,term,ch,obj,per,tag,qty,grade,price,mk) in R:
    C(ds,r,1,grp,b=True,al="center",wrap=True,fill=BEIGE if grp=="콘텐츠·바이럴" else None)
    C(ds,r,2,prod,b=True,wrap=True); C(ds,r,3,term,al="center"); C(ds,r,4,ch,al="center",wrap=True)
    C(ds,r,5,obj,wrap=True,sz=9); C(ds,r,6,per,wrap=True,sz=9); C(ds,r,7,tag,wrap=True,sz=9,col=GREEN)
    C(ds,r,8,qty,al="center",b=True); C(ds,r,9,grade,al="center",sz=9)
    C(ds,r,10,price,fmt=NUM,al="right")
    C(ds,r,11,f"=H{r}*J{r}",fmt=NUM,al="right",b=True)
    if mk: C(ds,r,12,f"=K{r}*0.2",fmt=NUM,al="right",b=True,col=GREEN,fill=YEL)
    else:  C(ds,r,12,0,fmt=NUM,al="right")
    C(ds,r,13,f"=K{r}+L{r}",fmt=NUM,al="right",b=True)
    ds.row_dimensions[r].height=46
    r+=1
last=r-1
# 합계
C(ds,r,2,"합  계",b=True,fill=GRAYF,al="center");
for c in (1,3,4,5,6,7,9,10): C(ds,r,c,fill=GRAYF)
C(ds,r,8,f"=SUM(H{first}:H{last})",b=True,fill=GRAYF,al="center")
C(ds,r,11,f"=SUM(K{first}:K{last})",fmt=NUM,b=True,fill=GRAYF,al="right")
C(ds,r,12,f"=SUM(L{first}:L{last})",fmt=NUM,b=True,fill=GRAYF,al="right",col=GREEN)
C(ds,r,13,f"=SUM(M{first}:M{last})",fmt=NUM,b=True,fill=GRAYF,al="right")
tot=r
C(ds,r+2,10,"부가세 (10%)",b=True,al="right",border=False); C(ds,r+2,13,f"=M{tot}*0.1",fmt=NUM,al="right",b=True)
C(ds,r+3,10,"총 합계 (VAT 포함)",b=True,al="right",col=WHITE,fill=BROWN); C(ds,r+3,13,f"=M{tot}*1.1",fmt=NUM,al="right",b=True,col=WHITE,fill=BROWN)
C(ds,r+5,10,"총견적(VAT별도) 2.5억 대비 여유",al="right",sz=9,border=False); C(ds,r+5,13,f"=250000000-M{tot}",fmt=NUM,al="right",sz=9,col=GREEN)
C(ds,r+6,10,"예상 마크업율",al="right",sz=9,border=False); C(ds,r+6,13,f"=L{tot}/M{tot}",fmt="0.0%",al="right",sz=9,col=GREEN)

# ============================================================ 월별 액션플랜
ms=wb.create_sheet("월별 액션플랜")
for i,w in enumerate([20,13,8,11,10,9,11,11,9,15]): ms.column_dimensions[get_column_letter(i+1)].width=w
C(ms,1,1,"월별 실행 타임라인 · 액션플랜  (Moonshot Rocket Launch)",b=True,sz=14,border=False)
C(ms,2,1,"Phase 1 (7~9월) 초기 반응·선택증거 확보 → Phase 2 (10~12월) 제품군 확장·재점화 · 단가/금액은 견적상세와 동일",col="868680",sz=9,border=False)
for c,h in enumerate(["항목","단가","7월","8월\n세일사전","9월\n세일★","10월","11월\n블프★","12월\n세일★","합계","금액"],1):
    C(ms,4,c,h,b=True,col=WHITE,fill=BROWN,al="center",wrap=True)
ms.row_dimensions[4].height=32
def msec(r,t):
    for c in range(1,11): C(ms,r,c,fill=BEIGE)
    C(ms,r,1,t,b=True,fill=BEIGE)
def mrow(r,name,unit,months,mk=False,cum=False):
    C(ms,r,1,name); C(ms,r,2,unit,fmt=NUM,al="right")
    for i,v in enumerate(months): C(ms,r,3+i,v if v else None,al="center")
    if cum:
        C(ms,r,9,"누적 60",al="center",b=True); C(ms,r,10,f"=60*B{r}",fmt=NUM,al="right",b=True)
    else:
        C(ms,r,9,f"=SUM(C{r}:H{r})",al="center",b=True)
        base=f"=I{r}*B{r}"
        C(ms,r,10,(f"=I{r}*B{r}*1.2" if mk else base),fmt=NUM,al="right",b=True)
msec(5,"■ 크리에이터 시딩")
mrow(6,"마이크로 (UGC)",500000,[30,55,35,30,40,45])
mrow(7,"나노",250000,[30,55,35,30,40,45])
mrow(8,"매크로",800000,[0,5,0,0,5,5])
mrow(9,"KOL 무가시딩",100000,[10,10,10,10,10,10])
mrow(10,"X (트위터) 시딩",500000,[0,5,0,0,5,5])
mrow(11,"네이버 블로그",300000,[0,5,0,0,5,5])
msec(12,"■ 콘텐츠 · 바이럴")
mrow(13,"파워페이지",800000,[0,3,0,0,3,3])
mrow(14,"커뮤니티 체험단 (마크업20%)",5000000,[0,1,0,0,1,1],mk=True)
mrow(15,"챌린저스 (마크업20%)",5000000,[0,1,0,0,1,0],mk=True)
mrow(16,"어필리에이트 (누적)",100000,[10,20,30,40,50,60],cum=True)
# 합계행
C(ms,18,1,"월 시딩 총건수",b=True,fill=GRAYF)
for i in range(6):
    col=get_column_letter(3+i); C(ms,18,3+i,f"=SUM({col}6:{col}11)",al="center",b=True,fill=GRAYF)
C(ms,18,2,"",fill=GRAYF); C(ms,18,9,"=SUM(C18:H18)",al="center",b=True,fill=GRAYF); C(ms,18,10,"",fill=GRAYF)
C(ms,19,9,"총 견적 합계",b=True,al="right"); C(ms,19,10,"=SUM(J6:J16)",fmt=NUM,al="right",b=True,col=GREEN)
C(ms,20,9,"견적상세 일치?",al="right",sz=9); C(ms,20,10,'=IF(ROUND(J19,0)=ROUND(견적상세!M'+str(tot)+',0),"일치","확인")',al="right",sz=9,col=GREEN)
for r in range(5,17): ms.row_dimensions[r].height=20

wb.save("proposal/올더베러_견적서_BAT.xlsx")
print("saved · 견적상세 합계행:",tot)
