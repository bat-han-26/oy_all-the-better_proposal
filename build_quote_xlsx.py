# -*- coding: utf-8 -*-
"""올더베러 견적 + 월별 액션플랜 — 수정 가능한 Excel (산식 포함)."""
import openpyxl
from openpyxl.styles import Font, PatternFill, Alignment, Border, Side
from openpyxl.utils import get_column_letter

BROWN="3D2B1B"; BEIGE="F3F2E7"; GREEN="0E5A30"; DARK="141412"; GRAYF="EFEFEC"
WHITE="FFFFFF"
thin=Side(style="thin",color="D8D5CC")
border=Border(left=thin,right=thin,top=thin,bottom=thin)
def cell(ws,r,c,v=None,bold=False,color="141412",fill=None,align="left",fmt=None,size=10,wrap=False):
    cc=ws.cell(row=r,column=c)
    if v is not None: cc.value=v
    cc.font=Font(name="Pretendard",size=size,bold=bold,color=color)
    if fill: cc.fill=PatternFill("solid",fgColor=fill)
    cc.alignment=Alignment(horizontal=align,vertical="center",wrap_text=wrap)
    cc.border=border
    if fmt: cc.number_format=fmt
    return cc

wb=openpyxl.Workbook()

# ===================== Sheet 1: 견적서 =====================
ws=wb.active; ws.title="견적서"
widths=[26,42,12,14,16,40]
for i,w in enumerate(widths): ws.column_dimensions[get_column_letter(i+1)].width=w
cell(ws,1,1,"올더베러 캠페인 견적서 (Quotation)",bold=True,size=16)
cell(ws,2,1,"올리브영 PB 올더베러 브랜드 빌딩 캠페인 · 단위 원 (VAT 별도)",color="868680",size=10)
cell(ws,3,5,"목표 공급가",bold=True,align="right")
cell(ws,3,6,250000000,bold=True,fmt="#,##0",align="left",color=GREEN)
# header
hdr=["구분","세부 항목","수량","단가","금액","비고"]
for c,h in enumerate(hdr,1): cell(ws,5,c,h,bold=True,color=WHITE,fill=BROWN,align="center" if c>2 else "left")
NUM="#,##0"
def section(r,name):
    for c in range(1,7): cell(ws,r,c,fill=BEIGE)
    cell(ws,r,1,name,bold=True,fill=BEIGE)
# row map
# 6 제작비 / 7 KV / 8 EGC / 9 HalfEGC / 10 크리에이터 / 11~15 / 16 콘텐츠바이럴 /17~19 /20 기타 /21 어필
section(6,"■ 제작비")
cell(ws,7,1,"제작"); cell(ws,7,2,"KV·디자인·영상·콘텐츠 제작 (EGC 제외)")
cell(ws,7,3,"1식",align="center");
cell(ws,7,4,"=E7",fmt=NUM,align="right")
cell(ws,7,5,"=$F$3-SUM(E8:E21)",fmt=NUM,align="right",bold=True)  # 2.5억 자동정산 (잔액)
cell(ws,7,6,"※ 총 공급가 2.5억 맞춤 자동 정산(잔액). 고정하려면 값 직접 입력",color="868680",size=9,wrap=True)
cell(ws,8,2,"EGC 콘텐츠"); cell(ws,8,3,12,align="center"); cell(ws,8,4,1000000,fmt=NUM,align="right"); cell(ws,8,5,"=C8*D8",fmt=NUM,align="right",bold=True); cell(ws,8,6,"월 2개 × 6개월",color="868680",size=9)
cell(ws,9,2,"Half EGC 콘텐츠"); cell(ws,9,3,6,align="center"); cell(ws,9,4,1500000,fmt=NUM,align="right"); cell(ws,9,5,"=C9*D9",fmt=NUM,align="right",bold=True); cell(ws,9,6,"제작 100만 + 출연 섭외 기본 50만(희망 출연자에 따라 상이) · 월 1개 × 6개월",color="868680",size=9,wrap=True)
section(10,"■ 크리에이터 시딩")
seed=[("매크로 (인스타·유튜브)",15,800000),("마이크로",150,400000),("나노",225,250000),("KOL 무가시딩",60,100000),("X (트위터) 시딩",15,500000)]
for i,(n,q,p) in enumerate(seed):
    r=11+i; cell(ws,r,2,n); cell(ws,r,3,q,align="center"); cell(ws,r,4,p,fmt=NUM,align="right"); cell(ws,r,5,f"=C{r}*D{r}",fmt=NUM,align="right",bold=True)
section(16,"■ 콘텐츠 · 바이럴")
viral=[("파워페이지 콘텐츠 바이럴",9,800000),("커뮤니티 체험단 (탑 3건)",3,5000000),("네이버 블로그",15,300000)]
for i,(n,q,p) in enumerate(viral):
    r=17+i; cell(ws,r,2,n); cell(ws,r,3,q,align="center"); cell(ws,r,4,p,fmt=NUM,align="right"); cell(ws,r,5,f"=C{r}*D{r}",fmt=NUM,align="right",bold=True)
    if n.startswith("네이버"): cell(ws,r,6,"단가 25만→30만 변경",color=GREEN,size=9)
section(20,"■ 기타")
cell(ws,21,2,"어필리에이트 (올영 쇼핑 큐레이터)"); cell(ws,21,3,"누적 60",align="center"); cell(ws,21,4,0,fmt=NUM,align="right"); cell(ws,21,5,0,fmt=NUM,align="right",bold=True); cell(ws,21,6,"성과형·무비용",color="868680",size=9)
# totals
cell(ws,22,2,"공급가액 소계",bold=True,fill=GRAYF,align="right"); cell(ws,22,1,"",fill=GRAYF);
for c in (3,4,6): cell(ws,22,c,fill=GRAYF)
cell(ws,22,5,"=SUM(E7:E21)",fmt=NUM,align="right",bold=True,fill=GRAYF)
cell(ws,23,2,"부가가치세 (10%)",bold=True,fill=GRAYF,align="right");
for c in (1,3,4,6): cell(ws,23,c,fill=GRAYF)
cell(ws,23,5,"=E22*0.1",fmt=NUM,align="right",bold=True,fill=GRAYF)
cell(ws,24,2,"합계 금액 (VAT 포함)",bold=True,color=WHITE,fill=BROWN,align="right");
for c in (1,3,4,6): cell(ws,24,c,fill=BROWN)
cell(ws,24,5,"=E22*1.1",fmt=NUM,align="right",bold=True,color=WHITE,fill=BROWN)
# 차액 체크
cell(ws,26,4,"목표 대비 차액",bold=True,align="right"); cell(ws,26,5,"=$F$3-E22",fmt=NUM,align="right",color=GREEN)
cell(ws,26,6,"0 이면 2.5억 정확히 일치",color="868680",size=9)
for r in range(5,25): ws.row_dimensions[r].height=22
ws.row_dimensions[9].height=30; ws.row_dimensions[7].height=30

# ===================== Sheet 2: 월별 액션플랜 =====================
ms=wb.create_sheet("월별 액션플랜")
mw=[22,13,8,11,10,9,11,11,9,15]
for i,w in enumerate(mw): ms.column_dimensions[get_column_letter(i+1)].width=w
cell(ms,1,1,"월별 실행 타임라인 · 액션플랜 (견적 연동)",bold=True,size=15)
cell(ms,2,1,"PHASE 1 (7~9월) 빠른 반응 확보 · PHASE 2 (10~12월) 대표성 확장 · 단가/금액은 견적서와 동일",color="868680",size=10)
hdr2=["항목","단가","7월","8월\n(세일사전)","9월\n세일★","10월","11월\n블프★","12월\n세일★","합계","금액"]
for c,h in enumerate(hdr2,1): cell(ms,4,c,h,bold=True,color=WHITE,fill=BROWN,align="center",wrap=True)
ms.row_dimensions[4].height=34
def msec(r,name):
    for c in range(1,11): cell(ms,r,c,fill=BEIGE)
    cell(ms,r,1,name,bold=True,fill=BEIGE)
def mrow(r,name,unit,months,note_green=False):
    cell(ms,r,1,name);
    if unit is None: cell(ms,r,2,"상시",align="center")
    else: cell(ms,r,2,unit,fmt="#,##0",align="right")
    for i,v in enumerate(months):
        cell(ms,r,3+i,v if v else None,align="center")
    cell(ms,r,9,f"=SUM(C{r}:H{r})",align="center",bold=True)
    if unit is None: cell(ms,r,10,"=견적서!E7",fmt="#,##0",align="right",bold=True)
    else: cell(ms,r,10,f"=I{r}*B{r}",fmt="#,##0",align="right",bold=True)
    if note_green: cell(ms,r,2).font=Font(name="Pretendard",size=10,bold=True,color=GREEN)
# rows
msec(5,"■ 크리에이터 시딩")
mrow(6,"매크로",800000,[0,5,0,0,5,5])
mrow(7,"마이크로",400000,[15,35,20,15,30,35])
mrow(8,"나노",250000,[25,55,35,25,40,45])
mrow(9,"KOL 무가시딩",100000,[10,10,10,10,10,10])
mrow(10,"X (트위터) 시딩",500000,[0,5,0,0,5,5])
msec(11,"■ 콘텐츠 · 바이럴")
mrow(12,"파워페이지",800000,[0,3,0,0,3,3])
mrow(13,"커뮤니티 체험단",5000000,[0,1,0,0,1,1])
mrow(14,"네이버 블로그",300000,[0,5,0,0,5,5],note_green=True)
mrow(15,"EGC 콘텐츠",1000000,[2,2,2,2,2,2])
mrow(16,"Half EGC 콘텐츠",1500000,[1,1,1,1,1,1])
msec(17,"■ 제작")
mrow(18,"KV·디자인·영상 (상시)",None,[None]*6)
msec(19,"■ 기타")
mrow(20,"어필리에이트 (누적)",0,[10,20,30,40,50,60])
cell(ms,20,9,"누적 60",align="center",bold=True)  # 누적값이라 합산 대신 표기
cell(ms,20,10,0,fmt="#,##0",align="right",bold=True)  # 무비용(텍스트 합계로 인한 오류 방지)
# 월 시딩 총건수
cell(ms,21,1,"월 시딩 총건수",bold=True,fill=GRAYF)
cell(ms,21,2,"",fill=GRAYF)
for i in range(6):
    col=get_column_letter(3+i)
    cell(ms,21,3+i,f"=SUM({col}6:{col}10)",align="center",bold=True,fill=GRAYF)
cell(ms,21,9,"=SUM(C21:H21)",align="center",bold=True,fill=GRAYF); cell(ms,21,10,"",fill=GRAYF)
# 월 집행액(제작 KV 제외)
cell(ms,22,1,"월 집행액 (제작 KV 제외)",bold=True,fill=GRAYF); cell(ms,22,2,"",fill=GRAYF)
for i in range(6):
    col=get_column_letter(3+i)
    cell(ms,22,3+i,f"=SUMPRODUCT({col}6:{col}16,$B$6:$B$16)",fmt="#,##0",align="right",bold=True,fill=GRAYF,size=9)
cell(ms,22,9,"",fill=GRAYF)
cell(ms,22,10,"=SUM(C22:H22)",fmt="#,##0",align="right",bold=True,fill=GRAYF)
# 공급가 합계 체크
cell(ms,24,9,"공급가 합계",bold=True,align="right"); cell(ms,24,10,"=SUM(J6:J20)",fmt="#,##0",align="right",bold=True,color=GREEN)
cell(ms,25,9,"견적 소계 일치?",align="right",size=9); cell(ms,25,10,'=IF(J24=견적서!E22,"일치","불일치")',align="right",size=9,color=GREEN)
for r in range(5,23): ms.row_dimensions[r].height=20

wb.save("proposal/올더베러_견적_액션플랜.xlsx")
print("saved proposal/올더베러_견적_액션플랜.xlsx")
