# -*- coding: utf-8 -*-
"""월별 실행 타임라인·액션플랜 — 히트맵 매트릭스 + 상단 마크업 요약 박스 (엑셀)."""
import openpyxl
from openpyxl.styles import Font, PatternFill, Alignment, Border, Side
from openpyxl.utils import get_column_letter

FONT="맑은 고딕"
WON='₩ #,##0'; PCT='0.0%'
DARK="404040"; PH1="333333"; PH2="9C9C97"; RED="B43C3C"; PINK="E8D2D2"
WHITE="FFFFFF"; INK="262626"; GREY="808080"; LINE="D9D9D9"; BLACK="1F1F1F"; BAND="F4F4F1"
thin=Side(style="thin",color="C9C9C4")
bd=Border(left=thin,right=thin,top=thin,bottom=thin)

try:
    from openpyxl.cell.rich_text import CellRichText, TextBlock
    from openpyxl.cell.text import InlineFont
    RICH=True
except Exception:
    RICH=False

wb=openpyxl.Workbook(); ws=wb.active; ws.title="월별 타임라인·액션플랜"
ws.sheet_view.showGridLines=False
for col,w in {1:26,2:11,3:11,4:11,5:11,6:11,7:11,8:12}.items():
    ws.column_dimensions[get_column_letter(col)].width=w

def C(r,c,v=None,b=False,col=INK,fill=None,al="center",fmt=None,sz=10,wrap=False,border=True):
    cell=ws.cell(r,c)
    if v is not None: cell.value=v
    cell.font=Font(name=FONT,size=sz,bold=b,color=col)
    if fill: cell.fill=PatternFill("solid",fgColor=fill)
    cell.alignment=Alignment(horizontal=al,vertical="center",wrap_text=wrap)
    if border: cell.border=bd
    if fmt: cell.number_format=fmt
    return cell

def grouplabel(r,text):
    ws.merge_cells(start_row=r,start_column=1,end_row=r,end_column=8)
    cell=ws.cell(r,1)
    if RICH:
        cell.value=CellRichText(
            TextBlock(InlineFont(rFont=FONT,sz=11,b=True,color="C0392B"),"■ "),
            TextBlock(InlineFont(rFont=FONT,sz=11,b=True,color="262626"),text))
    else:
        cell.value="■ "+text; cell.font=Font(name=FONT,size=11,bold=True,color="C0392B")
    cell.alignment=Alignment(horizontal="left",vertical="center")
    cell.fill=PatternFill("solid",fgColor=BAND)
    for c in range(1,9): ws.cell(r,c).fill=PatternFill("solid",fgColor=BAND); ws.cell(r,c).border=bd

# ===== 타이틀 =====
ws.merge_cells("A1:C2")
C(1,1,"월별 실행 타임라인 · 액션플랜",b=True,col=INK,al="left",sz=16,border=False)

# ===== 상단 요약 박스 (E1:H3) =====
labs=[("할인된\n총 원고료 및 실비",DARK,WHITE),("BAT\n할인율",RED,WHITE),
      ("BAT\n(총) 마크업",RED,WHITE),("올더베러\n(총) 견적",PINK,INK)]
for i,(t,f,tc) in enumerate(labs):
    C(1,5+i,t,b=True,col=tc,fill=f,wrap=True,sz=9)
ws.row_dimensions[1].height=30
C(2,5,216290000,b=True,fmt=WON,sz=10); C(2,6,0.389,b=True,fmt=PCT,sz=10)
C(2,7,33710000,b=True,fmt=WON,sz=10); C(2,8,250000000,b=True,fmt=WON,sz=10)
ws.merge_cells("E3:H3")
C(3,5,"BAT 제안 수수료 : 15.88%",b=True,col=WHITE,fill=RED,al="center",sz=10)
for c in range(5,9): ws.cell(3,c).fill=PatternFill("solid",fgColor=RED); ws.cell(3,c).border=bd

# ===== PHASE 밴드 (row5) =====
ws.merge_cells("B5:D5"); C(5,2,"PHASE 1 · 빠른 반응 확보",b=True,col=WHITE,fill=PH1)
for c in (2,3,4): ws.cell(5,c).fill=PatternFill("solid",fgColor=PH1); ws.cell(5,c).border=bd
ws.merge_cells("E5:G5"); C(5,5,"PHASE 2 · 대표성 확장",b=True,col=WHITE,fill=PH2)
for c in (5,6,7): ws.cell(5,c).fill=PatternFill("solid",fgColor=PH2); ws.cell(5,c).border=bd
C(5,1,"",border=False); C(5,8,"",border=False)

# ===== 헤더 (row6) =====
HDR=["항목  (건당 비용)","7월","8월\n세일 사전","9월\n세일 ★","10월","11월\n블프 ★","12월\n세일 ★","합계"]
for i,h in enumerate(HDR,1):
    C(6,i,h,b=True,col=WHITE,fill=DARK,wrap=True,al=("left" if i==1 else "center"))
ws.row_dimensions[6].height=34

def shade(q,rmax):
    t=q/rmax if rmax else 0
    v=int(228-t*150)
    return f"{v:02X}{v:02X}{v:02X}", (WHITE if v<145 else INK)

def item_row(r,name,dist,total):
    C(r,1,name,al="left",sz=10)
    if dist is None:  # 상시 ●
        for c in range(2,8):
            C(r,c,"●",col=WHITE,fill="8C8C8C")
    else:
        rmax=max(dist) or 1
        for i in range(6):
            q=dist[i]
            if q>0:
                f,tc=shade(q,rmax); C(r,i+2,q,b=True,col=tc,fill=f)
            else:
                C(r,i+2,"-",col=GREY,fill="FAFAF8")
    C(r,8,total,b=True,fill="F2F2EF")

DIST={
"나노":[10,20,12,10,18,20],"마이크로":[16,30,20,16,28,30],"매크로":[0,4,0,0,4,4],
"KOL 무가시딩":[5,5,5,5,5,5],"X (트위터) 시딩":[0,8,0,0,8,8],"네이버 블로그":[0,5,0,0,5,5],
"BATi 마이크로 앰배서더":[10,10,10,10,10,10],
"EGC 콘텐츠":[20,20,20,20,20,20],"Half EGC 콘텐츠":[4,4,4,4,4,4],
"파워페이지":[0,2,0,0,2,2],"커뮤니티 체험단":[0,1,0,0,1,0],"챌린저스":[0,1,0,0,1,0],
}
r=7
grouplabel(r,"크리에이터 시딩"); r+=1
for nm,tot in [("나노",90),("마이크로",140),("매크로",12),("KOL 무가시딩",30),("X (트위터) 시딩",24),("네이버 블로그",15),("BATi 마이크로 앰배서더",60)]:
    item_row(r,nm,DIST[nm],tot); r+=1
# 월 시딩 총건수
seed=[41,82,47,41,78,82]
C(r,1,"월 시딩 총건수",b=True,col=WHITE,fill=BLACK,al="right")
for i,v in enumerate(seed): C(r,i+2,v,b=True,col=WHITE,fill=BLACK)
C(r,8,"371 건",b=True,col=WHITE,fill=RED); r+=1
# 콘텐츠
grouplabel(r,"콘텐츠"); r+=1
item_row(r,"EGC 콘텐츠",DIST["EGC 콘텐츠"],120); r+=1
item_row(r,"Half EGC 콘텐츠",DIST["Half EGC 콘텐츠"],24); r+=1
item_row(r,"콘텐츠 제작 (KV·디자인·영상)",None,"6 (상시)"); r+=1
# 바이럴
grouplabel(r,"바이럴 · 전환"); r+=1
item_row(r,"파워페이지",DIST["파워페이지"],6); r+=1
item_row(r,"커뮤니티 체험단",DIST["커뮤니티 체험단"],2); r+=1
item_row(r,"챌린저스",DIST["챌린저스"],2); r+=1
# 솔루션
grouplabel(r,"솔루션"); r+=1
item_row(r,"스프레이ai 솔루션",None,"6 (상시)"); r+=1
# 월 예산 비중 / 금액
pct=["12%","23%","12%","12%","22%","19%"]; amt=["2,894만","5,714만","3,116만","2,894만","5,576만","4,804만"]
C(r,1,"월 예산 비중 / 금액",b=True,col=WHITE,fill=BLACK,al="right")
for i in range(6):
    C(r,i+2,f"{pct[i]}\n{amt[i]}",b=True,col=WHITE,fill=BLACK,wrap=True,sz=10)
C(r,8,"100%",b=True,col=WHITE,fill=RED)
ws.row_dimensions[r].height=30

# 푸터
C(r+2,1,"© 2026 BAT",col=GREY,sz=9,border=False,al="left")

wb.save("proposal/올더베러_월별타임라인_액션플랜_1p.xlsx")
print("saved rows used:",r)
