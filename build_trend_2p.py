# -*- coding: utf-8 -*-
"""손그림 그대로: ① UGC vs EGC vs Half EGC ② 마케팅 트렌드 변천사."""
from pptx import Presentation
from pptx.util import Inches, Pt
from pptx.enum.text import PP_ALIGN, MSO_ANCHOR
from pptx.enum.shapes import MSO_CONNECTOR
from pptx.oxml.ns import qn
from template_lib import *

prs=Presentation(); prs.slide_width=Inches(13.33); prs.slide_height=Inches(7.5)
BL=prs.slide_layouts[6]

def box(s,x,y,w,h,text,fill=WHITE,line=GREEN,tc=INK,size=11,bold=True,rounded=True,lw=1.3,sub=None):
    rect(s,x,y,w,h,fill=fill,line=line,line_w=lw,rounded=rounded)
    if sub:
        textbox(s,x,y+h*0.12,w,h*0.5,[(text,size,tc,bold)],align=PP_ALIGN.CENTER,anchor=MSO_ANCHOR.MIDDLE)
        textbox(s,x,y+h*0.55,w,h*0.4,[(sub,size-2,GRAY,False)],align=PP_ALIGN.CENTER,anchor=MSO_ANCHOR.MIDDLE)
    else:
        textbox(s,x,y,w,h,[(text,size,tc,bold)],align=PP_ALIGN.CENTER,anchor=MSO_ANCHOR.MIDDLE)

def line2(s,x1,y1,x2,y2,color=GRAY,w=1.2,dash=False):
    c=s.shapes.add_connector(MSO_CONNECTOR.STRAIGHT,Inches(x1),Inches(y1),Inches(x2),Inches(y2))
    c.line.color.rgb=color; c.line.width=Pt(w); c.shadow.inherit=False
    if dash:
        ln=c.line._get_or_add_ln(); pd=ln.makeelement(qn('a:prstDash'),{'val':'dash'}); ln.append(pd)
    return c

def xbox(s,x,y,w,h,label="영상 예시"):
    rect(s,x,y,w,h,fill=WHITE,line=LGRAY,line_w=1.3,rounded=False)
    line2(s,x,y,x+w,y+h,LGRAY,1.0); line2(s,x+w,y,x,y+h,LGRAY,1.0)
    oval(s,x+w/2-0.42,y+h/2-0.42,0.84,fill=WHITE,line=GREEN,line_w=1.5)
    textbox(s,x+w/2-0.42,y+h/2-0.42,0.84,0.84,[("▶",22,GREEN,True)],align=PP_ALIGN.CENTER,anchor=MSO_ANCHOR.MIDDLE)
    textbox(s,x,y+h-0.36,w,0.32,[(label,9.5,GRAY,False)],align=PP_ALIGN.CENTER)

def pagenum(s,n): textbox(s,13.33/2-0.5,7.05,1.0,0.3,[(str(n),10,LGRAY,False)],align=PP_ALIGN.CENTER)

# ===================== S1 =====================
s=prs.slides.add_slide(BL)
content_header(s,"콘텐츠","CONTENT","UGC vs EGC vs Half EGC")
cols=[("UGC","리얼함 · 현장감","3~4주","50만원",None),
      ("EGC","미디어 커머스 (정보·전문성)","1주","100만원",None),
      ("Half EGC","소셜 에비던스 (소비자 신뢰 형성)","2주","150만원","제작 100만 + 출연 섭외 50만(기본)")]
lx,lw=0.55,1.05
xs=[1.78,5.35,8.92]; cw=3.35
# 헤더 박스 + 키워드
for (name,feat,term,cost,note),x in zip(cols,xs):
    box(s,x,1.92,cw,0.6,name,fill=WHITE,line=GREEN,tc=GREEN,size=16)
    textbox(s,x,2.62,cw,0.42,[("+ "+feat,11,INK,True)],align=PP_ALIGN.CENTER,anchor=MSO_ANCHOR.MIDDLE)
    xbox(s,x,3.1,cw,2.25)
# 행 라벨 (시간 / 비용)
box(s,lx,5.5,lw,0.55,"제작기간",fill=BEIGE,line=GREEN,tc=GREEN,size=11)
box(s,lx,6.15,lw,0.62,"제작비",fill=GREEN,line=GREEN,tc=WHITE,size=12)
for (name,feat,term,cost,note),x in zip(cols,xs):
    box(s,x,5.5,cw,0.55,term,fill=WHITE,line=LGRAY,tc=INK,size=14)
    box(s,x,6.15,cw,0.62,cost,fill=BEIGE,line=LGRAY,tc=GREEN,size=16)
    if note: textbox(s,x,6.8,cw,0.22,[(note,8,GRAY,False)],align=PP_ALIGN.CENTER)
textbox(s,0.55,7.08,11.8,0.3,[("※ 제작비·제작기간은 콘텐츠 난이도·출연자 섭외에 따라 상이하며, 본 표는 예상 금액 기준입니다.",8.5,GRAY,False)])
pagenum(s,1)

# ===================== S2 =====================
s=prs.slides.add_slide(BL)
content_header(s,"트렌드","TREND","리테일 마케팅 트렌드 변천사")
# 상단 그로스 라벨 + 화살표
P1=(0.55,2.95); P2=(3.75,5.55); P3=(9.65,3.13)  # (x,w)
def toplabel(x,w,t1,t2,color=GREEN):
    textbox(s,x,1.95,w,0.32,[(t1,12,color,True)],align=PP_ALIGN.CENTER)
    if t2: textbox(s,x,2.26,w,0.28,[(t2,9.5,GRAY,False)],align=PP_ALIGN.CENTER)
toplabel(*P1,"레퍼런스 그로스","자사몰 중심")
toplabel(*P2,"콘텐츠 그로스","B2C 광고·콘텐츠 중심")
toplabel(*P3,"2026.06  What’s Next?",None,color=GREEN)
arrow(s,P1[0]+P1[1]+0.02,2.05,0.28); arrow(s,P2[0]+P2[1]+0.02,2.05,0.28)

cy,ch=2.62,3.95
# --- P1 container ---
x,w=P1; rect(s,x,cy,w,ch,fill=WHITE,line=GREEN,line_w=1.5,rounded=True)
box(s,x+0.2,cy+0.18,w-0.4,0.6,"레퍼런스 마케팅 Era",fill=GREEN,line=GREEN,tc=WHITE,size=11.5)
line2(s,x+w/2,cy+0.85,x+w/2,cy+1.75,GRAY,1.2,dash=True)
box(s,x+0.35,cy+1.75,w-0.7,0.72,"미디어 커머스",fill=BEIGE,line=GREEN,tc=GREEN,size=12)
textbox(s,x+0.25,cy+2.7,w-0.5,0.9,[("자사몰·리테일 레퍼런스\n중심으로 성장하던 시기",9.5,GRAY,False)],align=PP_ALIGN.CENTER)
# --- P2 container ---
x,w=P2; rect(s,x,cy,w,ch,fill=WHITE,line=GREEN,line_w=1.5,rounded=True)
box(s,x+0.25,cy+0.18,w-0.5,0.55,"인플루언서 마케팅 Era",fill=GREEN,line=GREEN,tc=WHITE,size=12)
maps=[("① MCN","공동구매"),("② UGC","어필리에이트"),("③ EGC","Half EGC")]
iw=1.5; ixs=[x+0.3,x+0.3+1.78,x+0.3+3.56]
for i,((it,ev),ix) in enumerate(zip(maps,ixs)):
    box(s,ix,cy+1.0,iw,0.5,it,fill=WHITE,line=GREEN,tc=INK,size=11)
    line2(s,ix+iw/2,cy+1.52,ix+iw/2,cy+2.05,GRAY,1.2,dash=True)
    box(s,ix,cy+2.05,iw,0.5,ev,fill=BEIGE,line=GREEN,tc=GREEN,size=11)
    if i<2: line2(s,ix+iw+0.04,cy+1.25,ixs[i+1]-0.04,cy+1.25,GREEN_B,1.3,dash=True)
textbox(s,x+0.3,cy+2.75,w-0.6,0.3,[("협업 플랫폼 사례",10,INK,True)],align=PP_ALIGN.CENTER)
for j,pf in enumerate(["Meta","TikTok"]):
    ox=x+w/2-1.15+j*1.3
    oval(s,ox,cy+3.1,0.62,fill=WHITE,line=GREEN,line_w=1.3)
    textbox(s,ox-0.3,cy+3.1,1.22,0.62,[(pf,9.5,GREEN,True)],align=PP_ALIGN.CENTER,anchor=MSO_ANCHOR.MIDDLE)
# 큰 화살표 P2->P3
arrow(s,P2[0]+P2[1]+0.04,cy+1.7,0.4)
# --- P3 container (highlight) ---
x,w=P3; rect(s,x,cy,w,ch,fill=GREEN,rounded=True)
textbox(s,x,cy+0.2,w,0.5,[("Micro-IMC",22,WHITE,True)],align=PP_ALIGN.CENTER)
oval(s,x+0.45,cy+0.95,1.3,fill=None,line=YELLOW,line_w=2.0)
oval(s,x+w-1.75,cy+0.95,1.3,fill=None,line=WHITE,line_w=2.0)
textbox(s,x+0.35,cy+1.45,1.3,0.4,[("인플루언서",9,WHITE,True)],align=PP_ALIGN.CENTER)
textbox(s,x+w-1.85,cy+1.45,1.3,0.4,[("퍼포먼스",9,WHITE,True)],align=PP_ALIGN.CENTER)
textbox(s,x,cy+2.3,w,0.3,[("교집합 = 우리의 해답",9.5,YELLOW,True)],align=PP_ALIGN.CENTER)
for i,b in enumerate(["+ 퍼포먼스 광고","+ 히어로컷","+ 미디어커머스 (EGC/UGC)"]):
    textbox(s,x+0.35,cy+2.75+i*0.34,w-0.6,0.3,[(b,10.5,YELLOW,True)])
pagenum(s,2)

prs.save("proposal/올더베러_콘텐츠유형_트렌드_2p.pptx")
print("saved", len(prs.slides._sldIdLst))
