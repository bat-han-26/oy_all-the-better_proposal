# -*- coding: utf-8 -*-
"""멜라나잇 구미 — 인플루언서 UGC 리뷰 시딩 릴스 썸네일 (인물+제품, 9:16, 플랫 일러스트)."""
from PIL import Image, ImageDraw, ImageFont, ImageFilter
W,H=1080,1920
F="/usr/share/fonts/truetype/wqy/wqy-zenhei.ttc"
def fnt(s): return ImageFont.truetype(F,s)
def T(d,xy,s,f,fill,anchor="la",bold=0):
    for dx in range(-bold,bold+1):
        for dy in range(-bold,bold+1):
            d.text((xy[0]+dx,xy[1]+dy),s,font=f,fill=fill,anchor=anchor)

img=Image.new("RGB",(W,H))
# ---- cozy night bedroom gradient ----
top=(74,46,86); bot=(38,28,40)
dd=ImageDraw.Draw(img)
for y in range(H):
    t=y/H; r=int(top[0]+(bot[0]-top[0])*t); g=int(top[1]+(bot[1]-top[1])*t); b=int(top[2]+(bot[2]-top[2])*t)
    dd.line([(0,y),(W,y)],fill=(r,g,b))
# window + moon (top-left)
dd.rounded_rectangle([70,150,360,540],18,fill=(34,30,58))
dd.line([215,150,215,540],fill=(70,60,95),width=8); dd.line([70,345,360,345],fill=(70,60,95),width=8)
dd.ellipse([250,210,330,290],fill=(245,238,205))
dd.ellipse([268,206,326,272],fill=(34,30,58))  # crescent
for sx,sy in [(120,250),(160,430),(300,470),(110,420)]:
    dd.ellipse([sx-4,sy-4,sx+4,sy+4],fill=(240,235,210))
# warm lamp glow (right)
glow=Image.new("L",(W,H),0); ImageDraw.Draw(glow).ellipse([W-360,120,W+260,740],fill=130)
glow=glow.filter(ImageFilter.GaussianBlur(150))
img=Image.composite(Image.new("RGB",(W,H),(252,224,150)),img,glow)
dd=ImageDraw.Draw(img)
# bed / blanket bottom
dd.rounded_rectangle([-40,1480,W+40,H+60],60,fill=(214,196,214))
dd.rounded_rectangle([-40,1480,W+40,1560],40,fill=(228,212,228))
# pillow behind
dd.rounded_rectangle([120,1360,960,1620],70,fill=(236,224,236))

# ================= PERSON (flat) =================
cx=540
SK=(244,206,180); SK2=(232,188,160); HAIR=(62,42,40); HAIR2=(48,32,32)
SWEAT=(150,120,168)
# hair back
dd.ellipse([cx-250,720,cx+250,1300],fill=HAIR)
# shoulders/sweater
dd.rounded_rectangle([cx-330,1230,cx+330,1640],120,fill=SWEAT)
dd.rounded_rectangle([cx-300,1300,cx+300,1640],90,fill=(162,132,180))
# neck
dd.rounded_rectangle([cx-58,1120,cx+58,1280],40,fill=SK2)
# face
dd.ellipse([cx-180,760,cx+180,1200],fill=SK)
# hair front (bangs + sides framing)
dd.chord([cx-200,690,cx+200,1060],185,355,fill=HAIR)
dd.ellipse([cx-205,820,cx-120,1180],fill=HAIR2)   # left side lock
dd.ellipse([cx+120,820,cx+205,1180],fill=HAIR2)   # right side lock
dd.pieslice([cx-185,720,cx+185,1120],195,345,fill=HAIR)  # top hair cap
# face features
ey=965
dd.arc([cx-118,ey-8,cx-44,ey+58],190,350,fill=(60,42,40),width=9)   # left eye (happy ^)
dd.arc([cx+44,ey-8,cx+118,ey+58],190,350,fill=(60,42,40),width=9)   # right eye
dd.arc([cx-122,ey-46,cx-44,ey+6],200,340,fill=HAIR2,width=8)        # brow L
dd.arc([cx+44,ey-46,cx+122,ey+6],200,340,fill=HAIR2,width=8)        # brow R
dd.ellipse([cx-118,1035,cx-66,1075],fill=(244,168,158))            # blush L
dd.ellipse([cx+66,1035,cx+118,1075],fill=(244,168,158))            # blush R
dd.line([cx-8,1010,cx-14,1052],fill=SK2,width=6); dd.line([cx-14,1052,cx+8,1058],fill=SK2,width=6)  # nose
dd.arc([cx-52,1052,cx+52,1118],20,160,fill=(196,96,92),width=10)   # smile

# ---- raised hand holding pouch (right side, near cheek) ----
# arm
dd.rounded_rectangle([cx+150,1250,cx+330,1560],70,fill=SWEAT)
# product pouch (held), tilted region near face right
ppx,ppy,ppw,pph=cx+150,830,300,470
pouch=Image.new("RGBA",(ppw+40,pph+40),(0,0,0,0)); pd=ImageDraw.Draw(pouch)
pd.rounded_rectangle([20,20,20+ppw,20+pph],34,fill=(243,239,227,255))
pd.rounded_rectangle([20,20,20+ppw,20+150],34,fill=(120,40,140,255))
pd.rectangle([20,110,20+ppw,170],fill=(120,40,140,255))
pd.ellipse([20+ppw/2-16,40,20+ppw/2+16,72],fill=(243,239,227,255))
pdf=ImageDraw.Draw(pouch)
T(pdf,(20+ppw/2,118),"ALL THE BETTER",fnt(30),(252,255,200,255),anchor="mm",bold=1)
gyy=20+255
for i,gxx in enumerate([20+78,20+ppw/2,20+ppw-78]):
    o=[-8,-20,-8][i]
    pdf.rounded_rectangle([gxx-44,gyy+o-34,gxx+44,gyy+o+40],26,fill=(226,58,70,255))
    pdf.rounded_rectangle([gxx-44,gyy+o-34,gxx+44,gyy+o+6],26,fill=(238,84,94,255))
T(pdf,(20+ppw/2,20+pph-118),"멜라나잇 구미",fnt(28),(60,45,55,255),anchor="mm",bold=1)
T(pdf,(20+ppw/2,20+pph-78),"식물성 멜라토닌 함유",fnt(20),(120,100,110,255),anchor="mm")
T(pdf,(20+ppw/2,20+pph-44),"Tart Cherry · 42g",fnt(18),(150,135,145,255),anchor="mm")
pouch=pouch.rotate(-10,expand=True,resample=Image.BICUBIC)
img.paste(pouch,(ppx-10,ppy-10),pouch)
dd=ImageDraw.Draw(img)
# hand/fingers over pouch bottom
hx=ppx+40
dd.rounded_rectangle([hx,ppy+pph-150,hx+260,ppy+pph+30],50,fill=SK)
for k in range(4):
    dd.rounded_rectangle([hx+30+k*56,ppy+pph-150,hx+72+k*56,ppy+pph+10],22,fill=SK2)

# ================= overlay text (casual UGC) =================
# top hook handwriting-ish
T(dd,(70,250),"요즘 내 ‘자기 전’ 루틴",fnt(58),(255,255,255),bold=2)
T(dd,(70,330),"솔직 후기 들려줄게",fnt(58),(252,255,161),bold=2)
# subtitle (reel style) lower-center
sub="“불 끄기 전 마지막으로 챙기는 거 하나”"
f2=fnt(40)
tw=dd.textlength(sub,font=f2)
dd.rounded_rectangle([(W-tw)//2-26,1690,(W+tw)//2+26,1762],18,fill=(0,0,0))
T(dd,(W//2,1726),sub,f2,(255,255,255),anchor="mm",bold=1)

# ================= reel UI =================
# 광고 badge
dd.rounded_rectangle([W-250,80,W-60,150],26,fill=(252,255,161))
T(dd,(W-155,115),"협찬·광고",fnt(30),(60,45,30),anchor="mm",bold=1)
# right action icons
ix=W-92
def heart(x,y):
    dd.ellipse([x-22,y-18,x-2,y+2],fill=(255,255,255)); dd.ellipse([x+2,y-18,x+22,y+2],fill=(255,255,255))
    dd.polygon([(x-21,y-4),(x+21,y-4),(x,y+26)],fill=(255,255,255))
heart(ix,980); T(dd,(ix,1030),"5.7천",fnt(26),(255,255,255),anchor="mm",bold=1)
dd.ellipse([ix-24,1090,ix+24,1138],outline=(255,255,255),width=8); T(dd,(ix,1170),"482",fnt(26),(255,255,255),anchor="mm",bold=1)
dd.polygon([(ix-24,1230),(ix+24,1212),(ix-2,1256)],outline=(255,255,255),width=2,fill=(255,255,255)); T(dd,(ix,1290),"공유",fnt(26),(255,255,255),anchor="mm",bold=1)
# music disc
dd.ellipse([ix-26,1340,ix+26,1392],fill=(40,30,45),outline=(255,255,255),width=5); dd.ellipse([ix-7,1359,ix+7,1373],fill=(255,255,255))
# bottom: handle + audio + progress
T(dd,(70,H-178),"@yujin.nightlog · 팔로우",fnt(34),(255,255,255),bold=1)
T(dd,(70,H-128),"♫  원본 오디오 · yujin.nightlog",fnt(28),(225,218,230))
T(dd,(70,H-84),"#오늘도배러 #자기전루틴 #멜라나잇구미 #타트체리",fnt(26),(252,255,161),bold=1)
dd.rounded_rectangle([60,H-44,W-60,H-36],4,fill=(120,110,120))
dd.rounded_rectangle([60,H-44,520,H-36],4,fill=(255,255,255))
# corner note
T(dd,(70,H-228),"인플루언서 UGC 리뷰 시딩 — 콘셉트 썸네일(일러스트) · 일반식품 효능 표현 배제",fnt(22),(200,190,205))

img.save("/tmp/melanight_ugc.png"); print("saved",img.size)
